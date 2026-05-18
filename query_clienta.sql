-- ============================================================
-- MIL PARTNER LEADS QUERY
-- Database : LEADS
-- BU       : MIL  (identified by LEAD_DATA.FSRC IN ('default-flow','default-flow-v2'))
--
-- Tables used:
--   LEADS.LEAD_DATA              : LEAD_ID, GCL_ID, MSCLICK_ID, FSRC
--   LEADS.LEAD_FOR_PARTNER       : LEAD_ID, PARTNER_ID, SEND_FLAG, SEND_DATE
--   LEADS.LEAD_PARTNER           : PARTNER_ID, INTERNAL_NAME
--   LEADS.LEAD_TRACKING_CONTRACTS: CONTRACT_ID, PRICE_FLAT, BAD_LEADS_PERCENT, START_DATE, END_DATE
--   LEADS.LEAD_TRACKING_PARTNER  : CONTRACT_ID, PARTNER_ID
--
-- Counting unit : unique (LEAD_ID, PARTNER_ID) where SEND_FLAG = 1
-- MIL filter    : LEAD_DATA.FSRC IN ('default-flow', 'default-flow-v2')
-- Source logic  : GCL_ID IS NOT NULL    → Google
--                 MSCLICK_ID IS NOT NULL → Bing
--                 both NULL             → Organic
-- Version logic : FSRC = 'default-flow'    → V1
--                 FSRC = 'default-flow-v2' → V2
-- ============================================================

WITH

-- ── Date spine ────────────────────────────────────────────────────────
date_params AS (
    SELECT
        TRUNC(SYSDATE, 'MM')                          AS mtd_start,
        TRUNC(SYSDATE)                                AS today,
        TRUNC(SYSDATE) - TRUNC(SYSDATE, 'MM') + 1    AS days_elapsed_mtd,
        LAST_DAY(SYSDATE) - TRUNC(SYSDATE, 'MM') + 1 AS total_days_in_month,
        TRUNC(ADD_MONTHS(SYSDATE, -1), 'MM')          AS prior_month_start,
        LAST_DAY(ADD_MONTHS(SYSDATE, -1))             AS prior_month_end,
        TRUNC(SYSDATE) - 29                           AS last30_start,
        TRUNC(SYSDATE)                                AS last30_end
    FROM dual
),

-- ── Contract rates ────────────────────────────────────────────────────
partner_contract AS (
    SELECT PARTNER_ID, cost_per_lead, scrub_rate
    FROM (
        SELECT
            ltp.PARTNER_ID,
            ltc.PRICE_FLAT        AS cost_per_lead,
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

-- ── MIL sends: FSRC-qualified, with source & version flags ───────────
sends AS (
    SELECT
        lfp.PARTNER_ID,
        lfp.SEND_DATE,
        CASE WHEN ld.GCL_ID     IS NOT NULL THEN 'GOOGLE'
             WHEN ld.MSCLICK_ID IS NOT NULL THEN 'BING'
             ELSE                                'ORGANIC'
        END AS lead_source,
        CASE WHEN ld.FSRC = 'default-flow'    THEN 'V1'
             WHEN ld.FSRC = 'default-flow-v2' THEN 'V2'
             ELSE                                   'OTHER'
        END AS lead_version
    FROM LEADS.LEAD_FOR_PARTNER lfp
    JOIN LEADS.LEAD_DATA        ld  ON ld.LEAD_ID = lfp.LEAD_ID
    WHERE lfp.SEND_FLAG = 1
      AND ld.FSRC IN ('default-flow', 'default-flow-v2')
),

-- ── MIL partner list: only partners with at least 1 MIL send ever ────
mil_partners AS (
    SELECT DISTINCT PARTNER_ID
    FROM sends
),

-- ── MTD aggregation ───────────────────────────────────────────────────
agg_mtd AS (
    SELECT
        s.PARTNER_ID,
        COUNT(*)                                                    AS leads_mtd,
        SUM(CASE WHEN s.lead_source  = 'GOOGLE' THEN 1 ELSE 0 END) AS google_mtd,
        SUM(CASE WHEN s.lead_source  = 'BING'   THEN 1 ELSE 0 END) AS bing_mtd,
        SUM(CASE WHEN s.lead_version = 'V1'     THEN 1 ELSE 0 END) AS v1_mtd,
        SUM(CASE WHEN s.lead_version = 'V2'     THEN 1 ELSE 0 END) AS v2_mtd
    FROM sends s
    CROSS JOIN date_params dp
    WHERE TRUNC(s.SEND_DATE) >= dp.mtd_start
      AND TRUNC(s.SEND_DATE) <  dp.today
    GROUP BY s.PARTNER_ID
),

-- ── Prior month aggregation ───────────────────────────────────────────
agg_prior AS (
    SELECT
        s.PARTNER_ID,
        COUNT(*)                                                    AS leads_prior_month,
        SUM(CASE WHEN s.lead_source  = 'GOOGLE' THEN 1 ELSE 0 END) AS google_prior_month,
        SUM(CASE WHEN s.lead_source  = 'BING'   THEN 1 ELSE 0 END) AS bing_prior_month,
        SUM(CASE WHEN s.lead_version = 'V1'     THEN 1 ELSE 0 END) AS v1_prior_month,
        SUM(CASE WHEN s.lead_version = 'V2'     THEN 1 ELSE 0 END) AS v2_prior_month
    FROM sends s
    CROSS JOIN date_params dp
    WHERE TRUNC(s.SEND_DATE) BETWEEN dp.prior_month_start AND dp.prior_month_end
    GROUP BY s.PARTNER_ID
),

-- ── Last 30 days aggregation ──────────────────────────────────────────
agg_last30 AS (
    SELECT
        s.PARTNER_ID,
        COUNT(*)                                                    AS leads_last_30,
        SUM(CASE WHEN s.lead_source  = 'GOOGLE' THEN 1 ELSE 0 END) AS google_last_30,
        SUM(CASE WHEN s.lead_source  = 'BING'   THEN 1 ELSE 0 END) AS bing_last_30,
        SUM(CASE WHEN s.lead_version = 'V1'     THEN 1 ELSE 0 END) AS v1_last_30,
        SUM(CASE WHEN s.lead_version = 'V2'     THEN 1 ELSE 0 END) AS v2_last_30
    FROM sends s
    CROSS JOIN date_params dp
    WHERE TRUNC(s.SEND_DATE) BETWEEN dp.last30_start AND dp.last30_end
    GROUP BY s.PARTNER_ID
)

-- ── Final SELECT ──────────────────────────────────────────────────────
SELECT
    lp.PARTNER_ID,
    lp.INTERNAL_NAME             AS partner_name,
    NVL(m.leads_mtd,         0)  AS leads_mtd,
    NVL(p.leads_prior_month, 0)  AS leads_prior_month,
    NVL(l.leads_last_30,     0)  AS leads_last_30,
    pc.cost_per_lead,
    pc.scrub_rate,
    NVL(m.v1_mtd,            0)  AS v1_mtd,
    NVL(p.v1_prior_month,    0)  AS v1_prior_month,
    NVL(l.v1_last_30,        0)  AS v1_last_30,
    NVL(m.v2_mtd,            0)  AS v2_mtd,
    NVL(p.v2_prior_month,    0)  AS v2_prior_month,
    NVL(l.v2_last_30,        0)  AS v2_last_30,
    NVL(m.google_mtd,         0) AS google_mtd,
    NVL(p.google_prior_month, 0) AS google_prior_month,
    NVL(l.google_last_30,     0) AS google_last_30,
    NVL(m.bing_mtd,           0) AS bing_mtd,
    NVL(p.bing_prior_month,   0) AS bing_prior_month,
    NVL(l.bing_last_30,       0) AS bing_last_30,
    dp.days_elapsed_mtd,
    dp.total_days_in_month

FROM LEADS.LEAD_PARTNER lp
CROSS JOIN date_params dp
JOIN mil_partners      mp ON mp.PARTNER_ID = lp.PARTNER_ID
JOIN partner_contract  pc ON pc.PARTNER_ID = lp.PARTNER_ID
LEFT JOIN agg_mtd      m  ON m.PARTNER_ID  = lp.PARTNER_ID
LEFT JOIN agg_prior    p  ON p.PARTNER_ID  = lp.PARTNER_ID
LEFT JOIN agg_last30   l  ON l.PARTNER_ID  = lp.PARTNER_ID

ORDER BY lp.INTERNAL_NAME
