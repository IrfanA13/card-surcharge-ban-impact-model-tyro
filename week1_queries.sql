-- My SQL analysis for the Tyro project
-- database is data/tyro_project.db, run it with python src/run_sql.py
-- periods follow Tyro's year, H1 is Jul to Dec and H2 is Jan to Jun


-- Q1. Tyro's share of all card spending at Australian merchants
-- first I add up the RBA monthly numbers into half years, then divide Tyro's TTV by it
-- I only keep half years where I have all 6 months
WITH market AS (
  SELECT period,
         SUM(total_acquired_au_m) AS market_m,
         COUNT(*) AS months
  FROM rba_cards_monthly
  GROUP BY period
)
SELECT h.period,
       ROUND(h.ttv_m / 1000.0, 1) AS tyro_ttv_bn,
       ROUND(m.market_m / 1000.0, 1) AS market_bn,
       ROUND(100.0 * h.ttv_m / m.market_m, 2) AS tyro_share_pct
FROM tyro_halfyear h
JOIN market m USING (period)
WHERE m.months = 6
ORDER BY h.start_date;


-- Q2. How the payments margin moved over time
-- LAG(x) gives me the value from the row before, LAG(x, 2) goes back 2 rows which is the same half last year
SELECT period,
       calc_margin_bps AS margin_bps,
       ROUND(calc_margin_bps - LAG(calc_margin_bps) OVER (ORDER BY start_date), 1) AS chg_vs_prior_half_bps,
       payments_gp_m,
       ROUND(100.0 * (payments_gp_m / LAG(payments_gp_m, 2) OVER (ORDER BY start_date) - 1), 1) AS gp_yoy_pct,
       ROUND(100.0 * (ttv_m / LAG(ttv_m, 2) OVER (ORDER BY start_date) - 1), 1) AS ttv_yoy_pct
FROM tyro_halfyear
ORDER BY start_date;


-- Q3. How much card spending is on credit vs debit
-- this matters because the interchange cut is much bigger on credit cards
-- H1FY27 only has Jul and Aug 2026 so far
SELECT period,
       ROUND(100.0 * SUM(credit_acquired_au_m) / SUM(total_acquired_au_m), 1) AS credit_share_pct,
       ROUND(SUM(credit_acquired_au_m) / 1000.0, 1) AS credit_bn,
       ROUND(SUM(debit_acquired_au_m) / 1000.0, 1) AS debit_bn,
       COUNT(*) AS months
FROM rba_cards_monthly
WHERE period IN (SELECT period FROM tyro_halfyear) OR period = 'H1FY27'
GROUP BY period
ORDER BY MIN(month);


-- Q4. Is Tyro more expensive than the market
-- Tyro's merchant fee vs the RBA average merchant fee for each year
-- for the market number I blend the debit and credit averages using how much is spent on credit
WITH rba_fy AS (
  SELECT 'FY' || SUBSTR(period, 5, 2) AS fy,
         AVG((eftpos + visa_debit_dom + mc_debit_dom) / 3.0) AS debit_pct,
         AVG((visa_credit_dom + mc_credit_dom) / 2.0) AS credit_pct,
         COUNT(*) AS quarters
  FROM rba_merchant_fees_quarterly
  GROUP BY 1
),
mix AS (
  SELECT 'FY' || SUBSTR(period, 5, 2) AS fy,
         SUM(credit_acquired_au_m) / SUM(total_acquired_au_m) AS credit_share
  FROM rba_cards_monthly
  GROUP BY 1
)
SELECT a.fy,
       a.msf_bps AS tyro_msf_bps,
       ROUND(r.debit_pct * 100, 0) AS rba_debit_bps,
       ROUND(r.credit_pct * 100, 0) AS rba_credit_bps,
       ROUND(100 * (x.credit_share * r.credit_pct + (1 - x.credit_share) * r.debit_pct), 0) AS rba_blended_bps,
       ROUND(a.msf_bps - 100 * (x.credit_share * r.credit_pct + (1 - x.credit_share) * r.debit_pct), 0) AS tyro_premium_bps,
       a.direct_cost_bps AS tyro_direct_cost_bps,
       a.payments_margin_bps AS tyro_margin_bps
FROM tyro_annual a
JOIN rba_fy r USING (fy)
JOIN mix x USING (fy)
ORDER BY a.fy;


-- Q5. Which industries are growing
-- PARTITION BY vertical means the LAG only looks back within the same industry
-- I multiply by 1.0 so SQLite doesnt do whole number division
WITH v AS (
  SELECT t.period, t.vertical, t.ttv_m,
         ROUND(100.0 * t.ttv_m / SUM(t.ttv_m) OVER (PARTITION BY t.period), 1) AS share_pct,
         ROUND(100.0 * (1.0 * t.ttv_m / LAG(t.ttv_m, 2) OVER (PARTITION BY t.vertical ORDER BY h.start_date) - 1), 1) AS yoy_pct
  FROM tyro_vertical t
  JOIN tyro_halfyear h USING (period)
)
SELECT period, vertical, share_pct, yoy_pct
FROM v
WHERE period IN ('H1FY26', 'H2FY26')
ORDER BY period, share_pct DESC;


-- Q6. What 1 basis point of margin is worth to Tyro
-- 1 bp of a year's TTV, and how big that is compared to total gross profit
SELECT fy,
       ROUND(ttv_m / 1e4, 2) AS gp_per_1bp_m,
       ROUND(100.0 * (ttv_m / 1e4) / total_gp_m, 1) AS pct_of_total_gp,
       ROUND(total_gp_m - LAG(total_gp_m) OVER (ORDER BY fy), 1) AS gp_growth_m
FROM tyro_annual
ORDER BY fy;


-- Q7. How much of Tyro's gross profit comes from payments
SELECT period,
       ROUND(100.0 * payments_gp_m / total_gp_m, 1) AS payments_pct,
       ROUND(100.0 * banking_gp_m / total_gp_m, 1) AS banking_pct,
       ROUND(100.0 * corporate_other_gp_m / total_gp_m, 1) AS corporate_other_pct
FROM tyro_halfyear
ORDER BY start_date;
