-- ============================================================
-- FW DAILY LEADS QUERY
-- Database : LEADS
-- BU       : FW  (LEAD_DATA.PSRC = 'FW')
--
-- Returns one row per (date, partner) combination within the
-- requested date range, plus a daily total of leads collected.
--
-- Parameters injected by Python (replace at runtime):
--   :start_date  -- first date of range  e.g. DATE '2026-05-01'
--   :end_date    -- last date of range   e.g. DATE '2026-05-14'
--   :utc_offset  -- ET offset as a number e.g. -4 (EDT) or -5 (EST)
--
-- Leads Collected : LEAD_DATA.CREATED_DATE converted from UTC to ET,
--                   filtered where PSRC = 'FW', independent of partner
-- Leads Sent      : LEAD_FOR_PARTNER where SEND_FLAG=1, joined to
--                   LEAD_DATA where PSRC='FW', grouped by SEND_DATE
-- ============================================================

WITH

-- ── Date spine: one row per date in the requested range ───────────────
date_spine AS (
    SELECT TRUNC(:start_date) + LEVEL - 1 AS report_date
    FROM dual
    CONNECT BY LEVEL <= TRUNC(:end_date) - TRUNC(:start_date) + 1
),

-- ── Active contracts: one per partner, current date window ───────────
partner_contract AS (
    SELECT PARTNER_ID, cost_per_lead, scrub_rate
    FROM (
        SELECT
            ltp.PARTNER_ID,
            ltc.PRICE_FLAT            AS cost_per_lead,
            ltc.BAD_LEADS_PERCENT / 100 AS scrub_rate,
            ROW_NUMBER() OVER (
                PARTITION BY ltp.PARTNER_ID
                ORDER BY ltc.START_DATE DESC
            ) AS rn
        FROM LEADS.LEAD_TRACKING_CONTRACTS ltc
        JOIN LEADS.LEAD_TRACKING_PARTNER   ltp ON ltp.CONTRACT_ID = ltc.CONTRACT_ID
        WHERE TRUNC(SYSDATE, 'MONTH') >= TRUNC(ltc.START_DATE, 'MONTH')
          AND (TRUNC(SYSDATE, 'MONTH') <= ltc.END_DATE OR ltc.END_DATE IS NULL)
    )
    WHERE rn = 1
),

-- ── FW partner list: active contract + at least 1 FW send ────────────
fw_partners AS (
    SELECT DISTINCT lfp.PARTNER_ID
    FROM LEADS.LEAD_FOR_PARTNER lfp
    JOIN LEADS.LEAD_DATA        ld  ON ld.LEAD_ID  = lfp.LEAD_ID
    JOIN partner_contract       pc  ON pc.PARTNER_ID = lfp.PARTNER_ID
    WHERE lfp.SEND_FLAG = 1
      AND ld.PSRC       = 'FW'
),

-- ── Leads collected per day (UTC → ET, PSRC='FW', all leads) ─────────
collected_daily AS (
    SELECT
        TRUNC(ld.CREATED_DATE + (:utc_offset / 24)) AS report_date,
        COUNT(*)                                     AS leads_collected
    FROM LEADS.LEAD_DATA ld
    WHERE ld.PSRC = 'FW'
      AND TRUNC(ld.CREATED_DATE + (:utc_offset / 24))
              BETWEEN TRUNC(:start_date) AND TRUNC(:end_date)
    GROUP BY TRUNC(ld.CREATED_DATE + (:utc_offset / 24))
),

-- ── Leads sent per day per partner (PSRC='FW') ────────────────────────
sent_daily AS (
    SELECT
        TRUNC(lfp.SEND_DATE) AS report_date,
        lfp.PARTNER_ID,
        COUNT(*)             AS leads_sent
    FROM LEADS.LEAD_FOR_PARTNER lfp
    JOIN LEADS.LEAD_DATA        ld ON ld.LEAD_ID = lfp.LEAD_ID
    WHERE lfp.SEND_FLAG = 1
      AND ld.PSRC       = 'FW'
      AND TRUNC(lfp.SEND_DATE) BETWEEN TRUNC(:start_date) AND TRUNC(:end_date)
    GROUP BY TRUNC(lfp.SEND_DATE), lfp.PARTNER_ID
),

-- ── Daily total leads sent (across all FW partners) ───────────────────
sent_daily_total AS (
    SELECT
        report_date,
        SUM(leads_sent) AS leads_sent_total
    FROM sent_daily
    GROUP BY report_date
)

-- ── Final SELECT: one row per (date, partner) ─────────────────────────
SELECT
    ds.report_date,
    lp.PARTNER_ID,
    lp.INTERNAL_NAME                  AS partner_name,
    NVL(sd.leads_sent,            0)  AS leads_sent,
    pc.cost_per_lead,
    pc.scrub_rate,
    -- Collected and total sent are the same for every partner row on that
    -- date; Excel will display them only in the summary columns (A-D)
    NVL(cd.leads_collected,       0)  AS leads_collected,
    NVL(st.leads_sent_total,      0)  AS leads_sent_total

FROM date_spine ds
CROSS JOIN (
    SELECT lp.PARTNER_ID, lp.INTERNAL_NAME
    FROM LEADS.LEAD_PARTNER lp
    JOIN fw_partners fp ON fp.PARTNER_ID = lp.PARTNER_ID
) lp
JOIN  partner_contract       pc ON pc.PARTNER_ID  = lp.PARTNER_ID
LEFT JOIN sent_daily         sd ON sd.report_date = ds.report_date
                               AND sd.PARTNER_ID  = lp.PARTNER_ID
LEFT JOIN collected_daily    cd ON cd.report_date = ds.report_date
LEFT JOIN sent_daily_total   st ON st.report_date = ds.report_date

ORDER BY ds.report_date, lp.INTERNAL_NAME
