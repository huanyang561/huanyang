const form = document.getElementById("model-form");
const statusEl = document.getElementById("status");
const runMeta = document.getElementById("runMeta");
const cards = document.getElementById("cards");
const summaryTable = document.getElementById("summaryTable");
const resultsTable = document.getElementById("resultsTable");
const validationTable = document.getElementById("validationTable");
const channelTable = document.getElementById("channelTable");
const inputMap = document.getElementById("inputMap");
const resetBtn = document.getElementById("resetBtn");
const excelFile = document.getElementById("excelFile");
const uploadExcelBtn = document.getElementById("uploadExcelBtn");
const downloadTemplateBtn = document.getElementById("downloadTemplateBtn");
const batchDownload = document.getElementById("batchDownload");
const batchTable = document.getElementById("batchTable");
const manualRowForm = document.getElementById("manual-row-form");
const runManualBatchBtn = document.getElementById("runManualBatchBtn");
const clearManualRowsBtn = document.getElementById("clearManualRowsBtn");
const manualRowCount = document.getElementById("manualRowCount");
const xlsxStatus = document.getElementById("xlsxStatus");
const manualRows = [];

const INPUT_MAP = [
  { 参数: "PD0", 主要来源: "银行", 取值范围: "(0%, 100%)；实务上企业一年期PD通常远低于100%", 说明: "不显式考虑物理气候风险的一年违约概率" },
  { 参数: "PD", 主要来源: "银行或第三方信用-气候模型", 取值范围: "PD0 ≤ PD < (1-q)PD0 + q", 说明: "考虑物理气候风险后的违约概率；这是论文自洽条件" },
  { 参数: "LGD0", 主要来源: "银行", 取值范围: "(0%, 100%]", 说明: "基础违约损失率" },
  { 参数: "EAD", 主要来源: "银行", 取值范围: "> 0", 说明: "违约暴露，用于转换为金额或RWA口径" },
  { 参数: "rho", 主要来源: "银行/监管公式", 取值范围: "(0%, 100%)，可留空自动计算", 说明: "资产相关性；默认用 Basel corporate correlation" },
  { 参数: "Q", 主要来源: "银行/监管", 取值范围: "(0%, 100%)；资本计量常用99.9%", 说明: "VaR置信水平，IRB常用99.9%" },
  { 参数: "q", 主要来源: "巨灾模型/气象地震等外部模型", 取值范围: "(0%, 100%]", 说明: "一年内发生定义好的灾害事件的概率" },
  { 参数: "loss threshold", 主要来源: "巨灾模型", 取值范围: "[0%, 100%]", 说明: "情景物理损害率达到该阈值则定义为巨灾事件" },
  { 参数: "intensity threshold", 主要来源: "气象/地震/巨灾模型", 取值范围: "有限数字，单位需与情景强度一致", 说明: "情景灾害强度达到该阈值则定义为巨灾事件" },
  { 参数: "LGD1", 主要来源: "银行+巨灾模型", 取值范围: "LGD0 ≤ LGD1 ≤ 100%", 说明: "气候事件发生后的违约损失率，可由损害率推导或外部指定" },
  { 参数: "alpha_hat", 主要来源: "模型求解", 取值范围: "由模型反推", 说明: "由PD0、PD、q通过自洽条件反推" },
  { 参数: "EL/VaR/UL", 主要来源: "模型输出", 取值范围: "由模型输出", 说明: "预期损失、尾部分位损失、非预期损失" },
];

const ALIASES = {
  id: ["id", "name", "客户", "借款人", "borrower", "obligor"],
  pd0Pct: ["pd0_pct", "pd0", "pd_0_pct", "基础pd", "基础pd_pct"],
  pdMode: ["pd_mode", "pd定义", "pd模式"],
  pdClimatePct: ["pd_climate_pct", "climate_pd_pct", "气候pd", "气候调整pd"],
  longRunQPct: ["long_run_q_pct", "长期q_pct", "长期频率_pct"],
  defaultUpliftWhenFrequencyDoublesPct: ["default_uplift_when_frequency_doubles_pct", "default_uplift_pct", "pd上升_pct"],
  qMode: ["q_mode", "q定义", "q模式"],
  qDirectPct: ["q_direct_pct", "q_pct", "q", "巨灾概率_pct"],
  returnPeriodYears: ["return_period_years", "return_period", "重现期", "重现期_年"],
  horizonYears: ["horizon_years", "horizon", "期限_年"],
  eventRatePerYear: ["event_rate_per_year", "annual_rate", "年发生率"],
  lossThresholdPct: ["loss_threshold_pct", "损失率阈值_pct"],
  lossRatesText: ["loss_rates", "scenario_loss_rates", "情景损失率"],
  intensityThreshold: ["intensity_threshold", "强度阈值"],
  intensityValuesText: ["intensities", "scenario_intensities", "情景强度"],
  lgd0Pct: ["lgd0_pct", "lgd0", "基础lgd"],
  confidenceLevelPct: ["confidence_level_pct", "q_confidence_pct", "置信水平_pct"],
  assetVolPct: ["asset_vol_pct", "asset_volatility_pct", "资产波动率_pct"],
  rhoPct: ["rho_pct", "asset_correlation_pct", "相关性_pct"],
  ead: ["ead", "exposure", "敞口"],
  maturityAdjustment: ["maturity_adjustment", "ma", "期限调整"],
  capitalMultiplier: ["capital_multiplier", "资本乘数"],
  lgd1OverridePct: ["lgd1_override_pct", "lgd1_pct", "气候lgd_pct"],
};

function erf(x) {
  const sign = x < 0 ? -1 : 1;
  const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741;
  const a4 = -1.453152027, a5 = 1.061405429, p = 0.3275911;
  x = Math.abs(x);
  const t = 1 / (1 + p * x);
  const y = 1 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
  return sign * y;
}

function Phi(x) {
  return 0.5 * (1 + erf(x / Math.SQRT2));
}

function PhiInv(p) {
  if (!(p > 0 && p < 1)) throw new Error(`Probability must be in (0,1), got ${p}`);
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.3577518672690, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const plow = 0.02425;
  const phigh = 1 - plow;
  let q, r, x;
  if (p < plow) {
    q = Math.sqrt(-2 * Math.log(p));
    x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p <= phigh) {
    q = p - 0.5;
    r = q * q;
    x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q /
      (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  } else {
    q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  for (let i = 0; i < 2; i++) {
    const err = Phi(x) - p;
    x -= err / (Math.exp(-0.5 * x * x) / Math.sqrt(2 * Math.PI));
  }
  return x;
}

function phi(x) {
  return Math.exp(-0.5 * x * x) / Math.sqrt(2 * Math.PI);
}

function pct(v) {
  return Number(v) / 100;
}

function asNumber(obj, key, fallback) {
  const v = obj[key];
  return v === "" || v === null || v === undefined ? fallback : Number(v);
}

function assertRange(label, value, low = null, high = null, options = {}) {
  const { lowInclusive = false, highInclusive = false, unit = "" } = options;
  if (!Number.isFinite(Number(value))) throw new Error(`${label} must be a finite number.`);
  value = Number(value);
  if (low !== null) {
    const bad = lowInclusive ? value < low : value <= low;
    if (bad) throw new Error(`${label}=${value}${unit} is out of range; expected ${lowInclusive ? ">=" : ">"} ${low}${unit}.`);
  }
  if (high !== null) {
    const bad = highInclusive ? value > high : value >= high;
    if (bad) throw new Error(`${label}=${value}${unit} is out of range; expected ${highInclusive ? "<=" : "<"} ${high}${unit}.`);
  }
}

function assertPct(label, value, options = {}) {
  assertRange(label, Number(value), 0, 100, {
    lowInclusive: !!options.allowZero,
    highInclusive: !!options.allow100,
    unit: "%",
  });
}

function assertProb(label, value, options = {}) {
  assertRange(label, Number(value), 0, 1, {
    lowInclusive: !!options.allowZero,
    highInclusive: !!options.allowOne,
  });
}

function parseNumberList(text) {
  const values = String(text || "").split(/[\s,;，；]+/).filter(Boolean).map(Number);
  if (!values.length || values.some((x) => Number.isNaN(x))) throw new Error("Scenario values must be numeric.");
  return values;
}

function qFromReturnPeriod(returnPeriodYears, horizonYears = 1) {
  return 1 - Math.exp(-horizonYears / returnPeriodYears);
}

function qFromPoissonRate(rate, horizonYears = 1) {
  return 1 - Math.exp(-rate * horizonYears);
}

function qFromLossThreshold(lossRates, threshold) {
  const normalized = lossRates.map((x) => x > 1 ? x / 100 : x);
  return normalized.filter((x) => x >= threshold).length / normalized.length;
}

function qFromIntensityThreshold(intensities, threshold) {
  return intensities.filter((x) => x >= threshold).length / intensities.length;
}

function pdFromFrequencyUplift(pd0, q, longRunQ, upliftWhenFrequencyDoubles) {
  return pd0 * (1 + upliftWhenFrequencyDoubles * ((q - longRunQ) / longRunQ));
}

function baselCorporateRho(pd0) {
  const ratio = (1 - Math.exp(-50 * pd0)) / (1 - Math.exp(-50));
  return 0.12 * ratio + 0.24 * (1 - ratio);
}

function inferAlphaHat(pd0, pdClimate, qClimate) {
  const c = PhiInv(pd0);
  const target = (pdClimate - (1 - qClimate) * pd0) / qClimate;
  if (!(target > 0 && target < 1)) throw new Error("Self-consistency is infeasible. Check PD0, climate PD and q.");
  return PhiInv(target) - c;
}

function lgdAfterJump(lgd0, alpha) {
  return lgd0 + (1 - Math.exp(-alpha)) * (1 - lgd0);
}

function cvVasicek(pd0, rho, qLevel) {
  const x = (Math.sqrt(rho) * PhiInv(qLevel) + PhiInv(pd0)) / Math.sqrt(1 - rho);
  return Phi(x);
}

function cvFirstOrder(pd0, rho, qLevel, qClimate, alphaHat) {
  const x = (Math.sqrt(rho) * PhiInv(qLevel) + PhiInv(pd0)) / Math.sqrt(1 - rho);
  return Phi(x) + qClimate * alphaHat * phi(x) / Math.sqrt(1 - rho);
}

function cvExactDefaultMixture(pd0, rho, qLevel, qClimate, alphaHat) {
  const x = (Math.sqrt(rho) * PhiInv(qLevel) + PhiInv(pd0)) / Math.sqrt(1 - rho);
  return (1 - qClimate) * Phi(x) + qClimate * Phi(x + alphaHat / Math.sqrt(1 - rho));
}

function lossCdfAsrf(loss, pd0, qClimate, alphaHat, rho, lgd0, lgd1) {
  if (loss <= 0) return 0;
  const theta0 = Math.min(Math.max(loss / lgd0, 1e-15), 1 - 1e-15);
  const theta1 = Math.min(Math.max(loss / lgd1, 1e-15), 1 - 1e-15);
  const term0 = Phi((Math.sqrt(1 - rho) * PhiInv(theta0) - PhiInv(pd0)) / Math.sqrt(rho));
  const term1 = Phi((Math.sqrt(1 - rho) * PhiInv(theta1) - PhiInv(pd0) - alphaHat) / Math.sqrt(rho));
  return (1 - qClimate) * term0 + qClimate * term1;
}

function lossVarBisection(pd0, qClimate, alphaHat, rho, lgd0, lgd1, qLevel) {
  let lo = 0;
  let hi = Math.max(lgd0, lgd1);
  for (let i = 0; i < 100; i++) {
    const mid = (lo + hi) / 2;
    if (lossCdfAsrf(mid, pd0, qClimate, alphaHat, rho, lgd0, lgd1) < qLevel) lo = mid;
    else hi = mid;
  }
  return (lo + hi) / 2;
}

function computeQ(payload) {
  const mode = payload.qMode || "direct";
  if (mode === "direct") {
    const qPct = asNumber(payload, "qDirectPct", 3);
    assertPct("q (%)", qPct, { allow100: true });
    return [pct(qPct), "direct"];
  }
  if (mode === "return_period") {
    assertRange("Return period", asNumber(payload, "returnPeriodYears", 33), 0);
    assertRange("Horizon", asNumber(payload, "horizonYears", 1), 0);
    return [qFromReturnPeriod(asNumber(payload, "returnPeriodYears", 33), asNumber(payload, "horizonYears", 1)), "return period"];
  }
  if (mode === "poisson_rate") {
    assertRange("Poisson event rate per year", asNumber(payload, "eventRatePerYear", 0.03), 0, null, { lowInclusive: true });
    assertRange("Horizon", asNumber(payload, "horizonYears", 1), 0);
    return [qFromPoissonRate(asNumber(payload, "eventRatePerYear", 0.03), asNumber(payload, "horizonYears", 1)), "Poisson rate"];
  }
  if (mode === "loss_threshold") {
    const losses = parseNumberList(payload.lossRatesText);
    losses.forEach((loss, i) => assertRange(`Scenario loss rate #${i + 1}`, loss, 0, null, { lowInclusive: true }));
    const thresholdPct = asNumber(payload, "lossThresholdPct", 10);
    assertPct("Loss threshold (%)", thresholdPct, { allowZero: true, allow100: true });
    return [qFromLossThreshold(losses, pct(thresholdPct)), "loss threshold"];
  }
  if (mode === "intensity_threshold") {
    const intensities = parseNumberList(payload.intensityValuesText);
    intensities.forEach((x, i) => assertRange(`Scenario intensity #${i + 1}`, x));
    assertRange("Intensity threshold", asNumber(payload, "intensityThreshold", 3));
    return [qFromIntensityThreshold(intensities, asNumber(payload, "intensityThreshold", 3)), "intensity threshold"];
  }
  throw new Error(`Unknown q mode: ${mode}`);
}

function computePd(payload, pd0, qClimate) {
  const mode = payload.pdMode || "direct";
  if (mode === "direct") {
    const pdPct = asNumber(payload, "pdClimatePct", 0.337);
    assertPct("Climate-adjusted PD (%)", pdPct);
    return [pct(pdPct), "direct"];
  }
  if (mode === "frequency_uplift") {
    assertPct("Long-run q (%)", asNumber(payload, "longRunQPct", 1.7));
    assertRange("Default uplift when frequency doubles (%)", asNumber(payload, "defaultUpliftWhenFrequencyDoublesPct", 16.1), -100, 1000, { unit: "%" });
    return [pdFromFrequencyUplift(pd0, qClimate, pct(asNumber(payload, "longRunQPct", 1.7)), pct(asNumber(payload, "defaultUpliftWhenFrequencyDoublesPct", 16.1))), "frequency uplift"];
  }
  throw new Error(`Unknown PD mode: ${mode}`);
}

function runModel(config) {
  const rho = config.rhoPct === "" || config.rhoPct === undefined ? baselCorporateRho(config.pd0) : pct(config.rhoPct);
  const alphaHat = inferAlphaHat(config.pd0, config.pdClimate, config.qClimate);
  const alpha = config.assetVol * alphaHat;
  const lgd1Model = lgdAfterJump(config.lgd0, alpha);
  const lgd1 = config.lgd1OverridePct === "" || config.lgd1OverridePct === undefined ? lgd1Model : pct(config.lgd1OverridePct);
  const cv0 = cvVasicek(config.pd0, rho, config.confidenceLevel);
  const cvPaper = cvFirstOrder(config.pd0, rho, config.confidenceLevel, config.qClimate, alphaHat);
  const cvMix = cvExactDefaultMixture(config.pd0, rho, config.confidenceLevel, config.qClimate, alphaHat);
  const pdEvent = Phi(PhiInv(config.pd0) + alphaHat);
  const baselineEl = config.lgd0 * config.pd0;
  const baselineVar = config.lgd0 * cv0;
  const baselineUl = baselineVar - baselineEl;
  const lgdMultiplier = 1 + ((lgd1 - config.lgd0) / config.lgd0) * config.qClimate;
  const paperEl = config.lgd0 * config.pdClimate * lgdMultiplier;
  const paperVar = config.lgd0 * cvPaper * lgdMultiplier;
  const paperUl = paperVar - paperEl;
  const exactEl = (1 - config.qClimate) * config.lgd0 * config.pd0 + config.qClimate * lgd1 * pdEvent;
  const exactVar = lossVarBisection(config.pd0, config.qClimate, alphaHat, rho, config.lgd0, lgd1, config.confidenceLevel);
  const exactUl = exactVar - exactEl;
  const scale = config.ead * config.maturityAdjustment * config.capitalMultiplier;

  const summary = {
    rho,
    alpha_hat: alphaHat,
    alpha,
    asset_drop: 1 - Math.exp(-alpha),
    lgd0: config.lgd0,
    lgd1_model: lgd1Model,
    lgd1_used: lgd1,
    pd_event: pdEvent,
    cv0,
    cv_paper: cvPaper,
    cv_exact_default_mixture: cvMix,
  };
  const results = [
    { 口径: "baseline Vasicek", EL: baselineEl, VaR: baselineVar, UL: baselineUl, RWA_like: scale * baselineUl, "UL增幅": 0 },
    { 口径: "paper first-order", EL: paperEl, VaR: paperVar, UL: paperUl, RWA_like: scale * paperUl, "UL增幅": paperUl / baselineUl - 1 },
    { 口径: "exact mixture CDF", EL: exactEl, VaR: exactVar, UL: exactUl, RWA_like: scale * exactUl, "UL增幅": exactUl / baselineUl - 1 },
  ];
  return { summary, results };
}

function payloadToConfig(payload) {
  const pd0Pct = asNumber(payload, "pd0Pct", 0.3);
  assertPct("PD0 (%)", pd0Pct);
  const pd0 = pct(pd0Pct);
  const [qClimate, qSource] = computeQ(payload);
  try {
    assertProb("q", qClimate, { allowOne: true });
  } catch (err) {
    throw new Error(`${err.message} Current q mode is ${qSource}; if a loss/intensity threshold produced q=0, no scenario exceeded the threshold, so lower the threshold or add more catastrophe scenarios.`);
  }
  const [pdClimate, pdSource] = computePd(payload, pd0, qClimate);
  assertProb("Climate-adjusted PD", pdClimate);
  if (pdClimate < pd0) {
    throw new Error(`Climate-adjusted PD=${fmtPct(pdClimate)} cannot be lower than PD0=${fmtPct(pd0)}; the physical risk model assumes climate risk does not reduce PD.`);
  }
  const maxPd = (1 - qClimate) * pd0 + qClimate;
  if (pdClimate >= maxPd) {
    throw new Error(`PD0/Climate PD/q fails the self-consistency condition: Climate PD must be below ${fmtPct(maxPd)}, currently ${fmtPct(pdClimate)}.`);
  }
  assertPct("LGD0 (%)", asNumber(payload, "lgd0Pct", 10), { allow100: true });
  assertPct("Confidence level Q (%)", asNumber(payload, "confidenceLevelPct", 99.9));
  assertRange("Asset volatility (%)", asNumber(payload, "assetVolPct", 30), 0, 500, { highInclusive: true, unit: "%" });
  assertRange("EAD", asNumber(payload, "ead", 1), 0);
  assertRange("Maturity adjustment", asNumber(payload, "maturityAdjustment", 1), 0);
  assertRange("Capital multiplier", asNumber(payload, "capitalMultiplier", 12.5), 0, 1000, { lowInclusive: true });
  assertRange("Simulation scenarios", Math.floor(asNumber(payload, "nScenarios", 50000)), 100, 2000000, { lowInclusive: true, highInclusive: true });
  assertRange("Loans per scenario", Math.floor(asNumber(payload, "nLoans", 20000)), 1, 10000000, { lowInclusive: true, highInclusive: true });
  if (payload.rhoPct !== "" && payload.rhoPct !== undefined) assertPct("rho override (%)", Number(payload.rhoPct));
  if (payload.lgd1OverridePct !== "" && payload.lgd1OverridePct !== undefined) assertPct("LGD1 override (%)", Number(payload.lgd1OverridePct), { allow100: true });
  if (payload.lgd1OverridePct !== "" && payload.lgd1OverridePct !== undefined && pct(Number(payload.lgd1OverridePct)) < pct(asNumber(payload, "lgd0Pct", 10))) {
    throw new Error(`LGD1 override=${fmtPct(pct(Number(payload.lgd1OverridePct)))} cannot be lower than LGD0=${fmtPct(pct(asNumber(payload, "lgd0Pct", 10)))}; physical damage reduces recoverable asset value, so event LGD should not be below baseline LGD.`);
  }
  return {
    pd0,
    qClimate,
    qSource,
    pdClimate,
    pdSource,
    lgd0: pct(asNumber(payload, "lgd0Pct", 10)),
    confidenceLevel: pct(asNumber(payload, "confidenceLevelPct", 99.9)),
    assetVol: pct(asNumber(payload, "assetVolPct", 30)),
    rhoPct: payload.rhoPct,
    ead: asNumber(payload, "ead", 1),
    maturityAdjustment: asNumber(payload, "maturityAdjustment", 1),
    capitalMultiplier: asNumber(payload, "capitalMultiplier", 12.5),
    nScenarios: Math.max(100, Math.floor(asNumber(payload, "nScenarios", 50000))),
    nLoans: Math.max(1, Math.floor(asNumber(payload, "nLoans", 20000))),
    seed: Math.floor(asNumber(payload, "seed", 20260603)),
    lgd1OverridePct: payload.lgd1OverridePct,
  };
}

function calculatePayload(payload) {
  const config = payloadToConfig(payload);
  const model = runModel(config);
  const summaryFormatted = [formatRecord(model.summary, ["rho", "alpha", "asset_drop", "lgd0", "lgd1_model", "lgd1_used", "pd_event", "cv0", "cv_paper", "cv_exact_default_mixture"], ["alpha_hat"])];
  const resultsFormatted = model.results.map((r) => formatRecord(r, ["EL", "VaR", "UL", "RWA_like", "UL增幅"], []));
  const response = {
    config,
    model,
    qPct: fmtPct(config.qClimate),
    qSource: config.qSource,
    pdClimatePct: fmtPct(config.pdClimate),
    pdSource: config.pdSource,
    summary: summaryFormatted,
    results: resultsFormatted,
  };
  if (payload.runSimulation) {
    const sim = simulatePortfolio(config, model);
    response.validation = validationTableRows(config, model, sim).map((r) => formatRecord(r, ["理论值", "模拟值", "绝对误差"], []));
    response.channelStats = channelStatsRows(sim).map((r) => formatRecord(r, ["mean_default_rate", "mean_loss_rate", "p99_loss", "p999_loss"], []));
  }
  return response;
}

function mulberry32(seed) {
  let a = seed >>> 0;
  return function () {
    a += 0x6D2B79F5;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomNormal(rng) {
  const u1 = Math.max(rng(), 1e-12);
  const u2 = rng();
  return Math.sqrt(-2 * Math.log(u1)) * Math.cos(2 * Math.PI * u2);
}

function poisson(lambda, rng) {
  if (lambda <= 0) return 0;
  if (lambda < 30) {
    const l = Math.exp(-lambda);
    let k = 0, p = 1;
    do { k++; p *= rng(); } while (p > l);
    return k - 1;
  }
  return Math.max(0, Math.round(lambda + Math.sqrt(lambda) * randomNormal(rng)));
}

function binomialApprox(n, p, rng) {
  const mean = n * p;
  if (mean < 30) return Math.min(n, poisson(mean, rng));
  const sd = Math.sqrt(n * p * (1 - p));
  return Math.min(n, Math.max(0, Math.round(mean + sd * randomNormal(rng))));
}

function quantile(values, q) {
  const arr = [...values].sort((a, b) => a - b);
  const idx = Math.min(arr.length - 1, Math.max(0, Math.floor(q * (arr.length - 1))));
  return arr[idx];
}

function mean(values) {
  return values.reduce((a, b) => a + b, 0) / values.length;
}

function simulatePortfolio(config, model) {
  const rng = mulberry32(config.seed);
  const rho = model.summary.rho;
  const alphaHat = model.summary.alpha_hat;
  const lgd1 = model.summary.lgd1_used;
  const out = [];
  const c = PhiInv(config.pd0);
  const denom = Math.sqrt(1 - rho);
  for (let i = 0; i < config.nScenarios; i++) {
    const s = randomNormal(rng);
    const event = rng() < config.qClimate ? 1 : 0;
    const threshold = (c - Math.sqrt(rho) * s + event * alphaHat) / denom;
    const conditionalPd = Phi(threshold);
    const defaults = binomialApprox(config.nLoans, conditionalPd, rng);
    const defaultRate = defaults / config.nLoans;
    const lgd = event ? lgd1 : config.lgd0;
    out.push({ climate_event: event, conditional_pd: conditionalPd, default_rate: defaultRate, loss_rate: lgd * defaultRate });
  }
  return out;
}

function validationTableRows(config, model, sim) {
  const exact = model.results.find((r) => r.口径 === "exact mixture CDF");
  const defaultRates = sim.map((r) => r.default_rate);
  const losses = sim.map((r) => r.loss_rate);
  const elSim = mean(losses);
  const varSim = quantile(losses, config.confidenceLevel);
  return [
    { 指标: "PD", 理论值: config.pdClimate, 模拟值: mean(defaultRates), 绝对误差: mean(defaultRates) - config.pdClimate },
    { 指标: "EL", 理论值: exact.EL, 模拟值: elSim, 绝对误差: elSim - exact.EL },
    { 指标: `VaR ${(config.confidenceLevel * 100).toFixed(1)}%`, 理论值: exact.VaR, 模拟值: varSim, 绝对误差: varSim - exact.VaR },
    { 指标: "UL", 理论值: exact.UL, 模拟值: varSim - elSim, 绝对误差: (varSim - elSim) - exact.UL },
  ];
}

function channelStatsRows(sim) {
  return [0, 1].map((event) => {
    const rows = sim.filter((r) => r.climate_event === event);
    if (!rows.length) return { climate_event: event ? "Cat event" : "No event", scenarios: 0, mean_default_rate: 0, mean_loss_rate: 0, p99_loss: 0, p999_loss: 0 };
    const losses = rows.map((r) => r.loss_rate);
    return {
      climate_event: event ? "Cat event" : "No event",
      scenarios: rows.length,
      mean_default_rate: mean(rows.map((r) => r.default_rate)),
      mean_loss_rate: mean(losses),
      p99_loss: quantile(losses, 0.99),
      p999_loss: quantile(losses, 0.999),
    };
  });
}

function fmtPct(x) {
  return `${(Number(x) * 100).toFixed(4)}%`;
}

function fmtDec(x) {
  return Number(x).toFixed(4);
}

function formatRecord(record, percentCols, decimalCols) {
  const out = { ...record };
  for (const c of percentCols) if (c in out) out[c] = fmtPct(out[c]);
  for (const c of decimalCols) if (c in out) out[c] = fmtDec(out[c]);
  return out;
}

function formPayload() {
  const data = new FormData(form);
  const payload = {};
  for (const [key, value] of data.entries()) payload[key] = value;
  payload.runSimulation = form.elements.runSimulation.checked;
  return payload;
}

function showMode(prefix, mode) {
  document.querySelectorAll(`[id^="${prefix}-"]`).forEach((el) => {
    el.classList.toggle("hidden", el.id !== `${prefix}-${mode}`);
  });
}

function table(rows) {
  if (!rows || rows.length === 0) return `<p class="empty">No data</p>`;
  const cols = Object.keys(rows[0]);
  const head = cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("");
  const body = rows.map((row) => `<tr>${cols.map((c) => `<td>${escapeHtml(row[c] ?? "")}</td>`).join("")}</tr>`).join("");
  return `<div class="table-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}

function escapeHtml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#039;");
}

function renderCards(data) {
  const summary = data.summary?.[0] || {};
  const exact = (data.results || []).find((r) => r.口径 === "exact mixture CDF") || {};
  const paper = (data.results || []).find((r) => r.口径 === "paper first-order") || {};
  const items = [
    ["q", data.qPct],
    ["Climate PD", data.pdClimatePct],
    ["alpha_hat", summary.alpha_hat],
    ["Exact UL increase", exact["UL增幅"]],
    ["Paper UL", paper.UL],
    ["Exact UL", exact.UL],
    ["LGD1 used", summary.lgd1_used],
    ["Asset drop", summary.asset_drop],
  ];
  cards.innerHTML = items.map(([label, value]) => `<article class="card"><span>${label}</span><strong>${value ?? "-"}</strong></article>`).join("");
}

function calculate() {
  statusEl.textContent = "running";
  try {
    const data = calculatePayload(formPayload());
    renderCards(data);
    summaryTable.innerHTML = table(data.summary);
    resultsTable.innerHTML = table(data.results);
    validationTable.innerHTML = table(data.validation || []);
    channelTable.innerHTML = table(data.channelStats || []);
    runMeta.textContent = `q: ${data.qSource}; climate PD: ${data.pdSource}; computed locally in browser`;
    statusEl.textContent = "ready";
  } catch (err) {
    cards.innerHTML = "";
    summaryTable.innerHTML = `<p class="empty error">${escapeHtml(err.message)}</p>`;
    resultsTable.innerHTML = "";
    validationTable.innerHTML = "";
    channelTable.innerHTML = "";
    runMeta.textContent = "Input validation error";
    statusEl.textContent = "error";
  }
}

function normKey(key) {
  return String(key).trim().toLowerCase().replace(/[%％]/g, "pct").replace(/[^a-z0-9\u4e00-\u9fff]+/g, "_").replace(/^_+|_+$/g, "");
}

function aliasedValue(row, canonical, fallback = "") {
  const normalized = {};
  for (const [k, v] of Object.entries(row)) normalized[normKey(k)] = v ?? "";
  for (const alias of ALIASES[canonical] || [canonical]) {
    const value = normalized[normKey(alias)];
    if (value !== "" && value !== null && value !== undefined) return value;
  }
  return fallback;
}

function batchRowToPayload(row, defaults) {
  const payload = { ...defaults, runSimulation: false };
  for (const canonical of Object.keys(ALIASES)) {
    const v = aliasedValue(row, canonical, "");
    if (v !== "") payload[canonical] = v;
  }
  if (!payload.pdMode) payload.pdMode = payload.pdClimatePct !== undefined && payload.pdClimatePct !== "" ? "direct" : "frequency_uplift";
  if (!payload.qMode) payload.qMode = "direct";
  return payload;
}

function resultColumns(calc) {
  const summary = calc.model.summary;
  const raw = Object.fromEntries(calc.model.results.map((r) => [r.口径, r]));
  return {
    calc_q_pct: calc.config.qClimate * 100,
    calc_pd_climate_pct: calc.config.pdClimate * 100,
    alpha_hat: summary.alpha_hat,
    alpha_pct: summary.alpha * 100,
    asset_drop_pct: summary.asset_drop * 100,
    rho_pct: summary.rho * 100,
    lgd1_used_pct: summary.lgd1_used * 100,
    baseline_el_pct: raw["baseline Vasicek"].EL * 100,
    baseline_var_pct: raw["baseline Vasicek"].VaR * 100,
    baseline_ul_pct: raw["baseline Vasicek"].UL * 100,
    paper_el_pct: raw["paper first-order"].EL * 100,
    paper_var_pct: raw["paper first-order"].VaR * 100,
    paper_ul_pct: raw["paper first-order"].UL * 100,
    paper_ul_increase_pct: raw["paper first-order"]["UL增幅"] * 100,
    exact_el_pct: raw["exact mixture CDF"].EL * 100,
    exact_var_pct: raw["exact mixture CDF"].VaR * 100,
    exact_ul_pct: raw["exact mixture CDF"].UL * 100,
    exact_ul_increase_pct: raw["exact mixture CDF"]["UL增幅"] * 100,
    exact_rwa_like: raw["exact mixture CDF"].RWA_like,
  };
}

function calculateBatch(rows, defaults) {
  return rows.map((row, index) => {
    const excelRow = index + 2;
    try {
      const calc = calculatePayload(batchRowToPayload(row, defaults));
      return { ...row, calc_excel_row: excelRow, calc_status: "ok", calc_error: "", ...resultColumns(calc) };
    } catch (err) {
      return { ...row, calc_excel_row: excelRow, calc_status: "error", calc_error: `Data row ${index + 1} (Excel row ${excelRow} if uploaded): ${err.message}` };
    }
  });
}

function downloadRowsAsXlsx(rows, filename) {
  if (!window.XLSX) throw new Error("SheetJS is not loaded. Excel download is unavailable.");
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "results");
  XLSX.writeFile(wb, filename);
}

function renderBatch(rows) {
  batchDownload.innerHTML = rows.length
    ? `<button class="ghost" type="button" id="downloadBatchBtn">Download Result Excel</button><span class="muted"> ${rows.length} rows calculated</span>`
    : "";
  batchTable.innerHTML = table(rows.slice(0, 50));
  const btn = document.getElementById("downloadBatchBtn");
  if (btn) btn.addEventListener("click", () => downloadRowsAsXlsx(rows, `climate_credit_results_${Date.now()}.xlsx`));
}

function templateRows() {
  return [
    { id: "A001", pd0_pct: 0.3, pd_mode: "direct", pd_climate_pct: 0.3367, q_mode: "direct", q_direct_pct: 3, lgd0_pct: 10, confidence_level_pct: 99.9, asset_vol_pct: 30, ead: 1 },
    { id: "A002", pd0_pct: 0.3, pd_mode: "frequency_uplift", long_run_q_pct: 1.7, default_uplift_pct: 16.1, q_mode: "loss_threshold", loss_threshold_pct: 10, loss_rates: "1, 3, 8, 12, 20, 0, 5, 16, 2, 30", lgd0_pct: 10, confidence_level_pct: 99.9, asset_vol_pct: 30, ead: 1 },
  ];
}

async function uploadExcel() {
  if (!window.XLSX) {
    batchTable.innerHTML = `<p class="empty error">SheetJS CDN is not loaded, so Excel upload is unavailable.</p>`;
    return;
  }
  if (!excelFile.files.length) {
    batchTable.innerHTML = `<p class="empty error">Please choose an Excel file first.</p>`;
    return;
  }
  const file = excelFile.files[0];
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const resultRows = calculateBatch(rows, { ...formPayload(), runSimulation: false });
  renderBatch(resultRows);
}

function updateManualCount() {
  manualRowCount.textContent = `${manualRows.length} rows`;
}

document.getElementById("qMode").addEventListener("change", (event) => showMode("q", event.target.value));
document.getElementById("pdMode").addEventListener("change", (event) => showMode("pd", event.target.value === "direct" ? "direct" : "uplift"));

resetBtn.addEventListener("click", () => {
  form.reset();
  showMode("q", "direct");
  showMode("pd", "direct");
  calculate();
});

form.addEventListener("submit", (event) => {
  event.preventDefault();
  calculate();
});

uploadExcelBtn.addEventListener("click", () => uploadExcel().catch((err) => {
  batchTable.innerHTML = `<p class="empty error">${escapeHtml(err.message)}</p>`;
}));

downloadTemplateBtn.addEventListener("click", () => {
  try {
    downloadRowsAsXlsx(templateRows(), "climate_credit_batch_template.xlsx");
  } catch (err) {
    batchTable.innerHTML = `<p class="empty error">${escapeHtml(err.message)}</p>`;
  }
});

manualRowForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const data = new FormData(manualRowForm);
  const row = {
    id: data.get("id") || `manual_${manualRows.length + 1}`,
    pd0_pct: data.get("pd0_pct"),
    pd_mode: "direct",
    pd_climate_pct: data.get("pd_climate_pct"),
    q_mode: "direct",
    q_direct_pct: data.get("q_direct_pct"),
    lgd0_pct: data.get("lgd0_pct"),
  };
  manualRows.push(row);
  manualRowForm.reset();
  updateManualCount();
  batchTable.innerHTML = table(manualRows);
});

runManualBatchBtn.addEventListener("click", () => {
  if (!manualRows.length) {
    batchTable.innerHTML = `<p class="empty error">No manual rows added.</p>`;
    return;
  }
  renderBatch(calculateBatch(manualRows, { ...formPayload(), runSimulation: false }));
});

clearManualRowsBtn.addEventListener("click", () => {
  manualRows.length = 0;
  updateManualCount();
  batchDownload.innerHTML = "";
  batchTable.innerHTML = "";
});

if (window.XLSX) xlsxStatus.textContent = "Excel support is ready. Calculations and files stay in this browser.";
else xlsxStatus.textContent = "Excel library was not loaded. Single-record calculation still works.";

showMode("q", "direct");
showMode("pd", "direct");
inputMap.innerHTML = table(INPUT_MAP);
updateManualCount();
calculate();
