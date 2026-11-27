-- =====================================================================
-- P x Q CE/RJ - captive customers, SAMP document base, AGGREGATED extraction
-- Source : bi_brrj_bill.bt_brrj_zisu_t_samp_m_dt (RJ)
--          CE: same table in the CE schema (name to be confirmed by IT)
-- Grain  : one row per accounting month x SAMP file type x class x modality x subgroup
-- Period : replace 202610 with the accounting month (AAAAMM)
-- Scope  : file types 5 and 6 (free customers, already in the SAMP LV base) and 7 (others) are
--          excluded, as instructed by the billing area
-- Output : CSV, all records (no LIMIT)
-- Removed from the original extraction: mandt, document date, applied tariffs (tarifa_*: revenue /
--          quantity, recomputed in the P x Q), contracted demand, injected energy and credit balances
-- =====================================================================
SELECT
    -- grouping keys
     data_yyyymm                      AS mes_contabil
    ,tipo_arquivo                     -- SAMP file: captive / GD, billing / re-billing
    ,classe_consumo                   -- SAMP class code
    ,modal_tarifaria                  -- SAMP tariff modality code
    ,subgrupo_tarifa                  -- SAMP subgroup code (incl. social tariff bands)
    -- next delivery (requested to IT): reference month, CNR flag, reading lot
    -- ,mes_referencia
    -- ,flag_cnr
    -- ,lote_leitura

    -- counts
    ,COUNT(*)                         AS qtd_documentos
    ,SUM(quantidade)                  AS qtd_consumidores

    -- billed energy (kWh), by time-of-use period
    ,SUM(cons_ponta_tusd)             AS cons_ponta_tusd
    ,SUM(cons_fponta_tusd)            AS cons_fponta_tusd
    ,SUM(cons_intermed_tusd)          AS cons_intermed_tusd
    ,SUM(cons_reservado_tusd)         AS cons_reservado_tusd    -- special time slot (irrigation / aquaculture)
    ,SUM(cons_nsa_tusd)               AS cons_nsa_tusd          -- single period (conventional)

    -- energy revenue (R$, without taxes), by period
    ,SUM(receita_ponta)               AS receita_ponta
    ,SUM(receita_fponta)              AS receita_fponta
    ,SUM(receita_intermed)            AS receita_intermed
    ,SUM(receita_reservado)           AS receita_reservado
    ,SUM(receita_nsa)                 AS receita_nsa

    -- billed demand and overrun (kW)
    ,SUM(dem_ponta)                   AS dem_ponta
    ,SUM(dem_fponta)                  AS dem_fponta
    ,SUM(dem_nsa)                     AS dem_nsa
    ,SUM(ult_dem_ponta)               AS ult_dem_ponta
    ,SUM(ult_dem_fponta)              AS ult_dem_fponta
    ,SUM(ult_dem_nsa)                 AS ult_dem_nsa

    -- demand and overrun revenue (R$, without taxes)
    ,SUM(receita_dem_ponta)           AS receita_dem_ponta
    ,SUM(receita_dem_fponta)          AS receita_dem_fponta
    ,SUM(receita_dem_nsa)             AS receita_dem_nsa
    ,SUM(receita_ult_dem_ponta)       AS receita_ult_dem_ponta
    ,SUM(receita_ult_dem_fponta)      AS receita_ult_dem_fponta
    ,SUM(receita_ult_dem_nsa)         AS receita_ult_dem_nsa

    -- net metering (SCEE): compensated energy (kWh) and credits used
    ,SUM(energia_scee_p)              AS energia_scee_p
    ,SUM(energia_scee_fp)             AS energia_scee_fp
    ,SUM(energia_scee_int)            AS energia_scee_int
    ,SUM(energia_scee_res)            AS energia_scee_res
    ,SUM(energia_scee_nsa)            AS energia_scee_nsa
    ,SUM(creditosutilizados)          AS creditosutilizados

    -- tariff flag revenue (R$, without taxes): total and by period
    ,SUM(receita_bandeira)            AS receita_bandeira
    ,SUM(receitabandeirasnp)          AS receitabandeirasnp
    ,SUM(receitabandeirasfp)          AS receitabandeirasfp
    ,SUM(receitabandeirasint)         AS receitabandeirasint
    ,SUM(receitabandeirasrv)          AS receitabandeirasrv
    ,SUM(receitabandeirasnsa)         AS receitabandeirasnsa

    -- taxes and invoice total (R$)
    ,SUM(pis)                         AS pis
    ,SUM(cofins)                      AS cofins
    ,SUM(icms)                        AS icms
    ,SUM(total_fat)                   AS total_fat
FROM bi_brrj_bill.bt_brrj_zisu_t_samp_m_dt
WHERE data_yyyymm = '202610'
  AND tipo_arquivo NOT IN (5, 6, 7)
GROUP BY data_yyyymm
        ,tipo_arquivo
        ,classe_consumo
        ,modal_tarifaria
        ,subgrupo_tarifa
        -- ,mes_referencia
        -- ,flag_cnr
        -- ,lote_leitura
;
