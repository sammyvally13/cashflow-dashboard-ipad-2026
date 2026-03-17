const inputTypes = {};
let expensesPieChart = null;
let savingsProjectionChart = null;
let latestMetrics = null;
let isSelfEmployed = false;
let isPensioner = false;

const TEMPLATE_FILE = "Cashflow Summary template.xlsx";

const BASE_PERSONAL_FIELDS = [
  "food",
  "apparel",
  "transport",
  "subscriptions",
  "bills",
  "allowances",
  "recreation",
  "holidays"
];

const BASE_LOAN_FIELDS = ["loanHouse", "loanCar", "loanEducation", "loanBusiness", "loanPersonal"];
const BASE_INSURANCE_FIELDS = ["hospPremium", "wholePremium", "termPremium", "ciPremium"];
const BASE_SAVINGS_FIELDS = ["endowment", "investments", "fixedDeposits", "stocks", "otherSavings"];

function fmt(value) {
  const num = Number(value) || 0;
  return "$" + num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function toNumber(value) {
  const parsed = parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getMonthlyValue(fieldId, value) {
  const type = inputTypes[fieldId] || "Monthly";
  const num = toNumber(value);
  return type === "Annual" ? num / 12 : num;
}

function setType(fieldId, type) {
  inputTypes[fieldId] = type;
  document
    .querySelectorAll(`.toggle-btn[data-field="${fieldId}"]`)
    .forEach((btn) => btn.classList.toggle("active", btn.dataset.type === type));
  updateDashboard();
}

// CPF ceilings (SGD)
const MONTHLY_OW_CEILING = 8000;
const ANNUAL_OW_CEILING = 102000;

// Age-tiered CPF contribution rates.
// oa, sa, ma are the employee-only allocation as a fraction of gross wage
// (i.e. they sum to the employee rate for each tier).
// Source: CPF Board published rates.
const CPF_RATE_TABLE = [
  // age ≤ 55
  { maxAge: 55,       employee: 0.20,  employer: 0.17,  oa: 0.130, sa: 0.015, ma: 0.055 },
  // age > 55 to 60
  { maxAge: 60,       employee: 0.15,  employer: 0.15,  oa: 0.080, sa: 0.010, ma: 0.060 },
  // age > 60 to 65
  { maxAge: 65,       employee: 0.095, employer: 0.11,  oa: 0.030, sa: 0.005, ma: 0.060 },
  // age > 65 to 70
  { maxAge: 70,       employee: 0.07,  employer: 0.085, oa: 0.010, sa: 0.000, ma: 0.060 },
  // age > 70
  { maxAge: Infinity, employee: 0.05,  employer: 0.075, oa: 0.005, sa: 0.000, ma: 0.045 },
];

// ─── SEP Medisave contribution rates (CPF MARates2026_amended) ───────────────
// Contributions are computed on the full Net Trade Income (NTI).
// Age is assessed as at 1 January of the contribution year.
// Both tables share the same NTI band thresholds:
//   ≤ $6,000         → No contribution required
//   $6,001–$12,000   → Band 1: flat % of full NTI (by age)
//   $12,001–$18,000  → Band 2: phase-in formula ($ amount, linear interpolation)
//   Above $18,000    → Band 3: flat % of full NTI, table/age-specific cap
// All caps are reached at NTI = $96,000.

const SEP_MA_EXEMPT      = 6000;
const SEP_MA_BAND1_UPPER = 12000;
const SEP_MA_BAND2_UPPER = 18000;

// Age as at 1 January of the current year (CPF's convention for SEP MA)
function getAgeAsAtJan1(dob) {
  const birthDate = dob ? new Date(dob) : null;
  if (!birthDate || Number.isNaN(birthDate.valueOf())) return 0;
  const jan1 = new Date(new Date().getFullYear(), 0, 1);
  let age = jan1.getFullYear() - birthDate.getFullYear();
  const m = jan1.getMonth() - birthDate.getMonth();
  if (m < 0 || (m === 0 && jan1.getDate() < birthDate.getDate())) age -= 1;
  return age;
}

// annualNTI  – estimated annual Net Trade Income (gross income × 12)
// age        – age as at 1 Jan of contribution year
// pensioner  – use Table 2 (Pensioner) rates when true,
//              Table 1 (Non-Pensioner) rates when false
function calculateSEPMedisave(annualNTI, age, pensioner) {
  if (annualNTI <= SEP_MA_EXEMPT) {
    return { annualMA: 0, monthlyMA: 0, effectiveRate: 0, annualCap: 0,
             aboveExempt: false, atCeiling: false, age, pensioner };
  }

  let annualMA = 0;
  let effectiveRate = 0;
  let annualCap = 0;

  if (pensioner) {
    // ── Table 2: Pensioners (CPF MARates2026_amended) ─────────────────
    //
    // Band 1  ($6,001–$12,000) – % of full NTI
    //   Below 35     : 4.00%
    //   35 to < 45   : 4.50%
    //   45 to < 50   : 5.00%
    //   50 and above : 5.25%
    //
    // Band 2  ($12,001–$18,000) – phase-in to 6% ($ formula)
    //   Below 35     : 480 + 0.10  × (NTI − 12,000)
    //   35 to < 45   : 540 + 0.09  × (NTI − 12,000)
    //   45 to < 50   : 600 + 0.08  × (NTI − 12,000)
    //   50 and above : 630 + 0.075 × (NTI − 12,000)
    //
    // Band 3  (above $18,000) – 6.00% of NTI, maximum $5,760

    annualCap = 5760;

    if (annualNTI <= SEP_MA_BAND1_UPPER) {
      effectiveRate = age < 35 ? 0.04 : age < 45 ? 0.045 : age < 50 ? 0.05 : 0.0525;
      annualMA = effectiveRate * annualNTI;
    } else if (annualNTI <= SEP_MA_BAND2_UPPER) {
      if      (age < 35) annualMA = 480 + 0.10  * (annualNTI - SEP_MA_BAND1_UPPER);
      else if (age < 45) annualMA = 540 + 0.09  * (annualNTI - SEP_MA_BAND1_UPPER);
      else if (age < 50) annualMA = 600 + 0.08  * (annualNTI - SEP_MA_BAND1_UPPER);
      else               annualMA = 630 + 0.075 * (annualNTI - SEP_MA_BAND1_UPPER);
      effectiveRate = annualMA / annualNTI;
    } else {
      annualMA = Math.min(0.06 * annualNTI, annualCap);
      effectiveRate = 0.06;
    }

  } else {
    // ── Table 1: Non-Pensioners (CPF MARates2026_amended) ─────────────
    //
    // Band 1  ($6,001–$12,000) – % of full NTI (same base rates as Table 2)
    //   Below 35     : 4.00%
    //   35 to < 45   : 4.50%
    //   45 to < 50   : 5.00%
    //   50 and above : 5.25%
    //
    // Band 2  ($12,001–$18,000) – steeper phase-in to Band 3 rate ($ formula)
    //   Below 35     : 480 + 0.1600 × (NTI − 12,000)   → 4% → 8%
    //   35 to < 45   : 540 + 0.1800 × (NTI − 12,000)   → 4.5% → 9%
    //   45 to < 50   : 600 + 0.2000 × (NTI − 12,000)   → 5% → 10%
    //   50 and above : 630 + 0.2100 × (NTI − 12,000)   → 5.25% → 10.5%
    //
    // Band 3  (above $18,000) – age-specific rate & cap, all cap at NTI $96,000
    //   Below 35     :  8.00%, maximum  $7,680
    //   35 to < 45   :  9.00%, maximum  $8,640
    //   45 to < 50   : 10.00%, maximum  $9,600
    //   50 and above : 10.50%, maximum $10,080

    const band3 = age < 35
      ? { rate: 0.08,  cap: 7680  }
      : age < 45
      ? { rate: 0.09,  cap: 8640  }
      : age < 50
      ? { rate: 0.10,  cap: 9600  }
      : { rate: 0.105, cap: 10080 };

    annualCap = band3.cap;

    if (annualNTI <= SEP_MA_BAND1_UPPER) {
      effectiveRate = age < 35 ? 0.04 : age < 45 ? 0.045 : age < 50 ? 0.05 : 0.0525;
      annualMA = effectiveRate * annualNTI;
    } else if (annualNTI <= SEP_MA_BAND2_UPPER) {
      if      (age < 35) annualMA = 480 + 0.1600 * (annualNTI - SEP_MA_BAND1_UPPER);
      else if (age < 45) annualMA = 540 + 0.1800 * (annualNTI - SEP_MA_BAND1_UPPER);
      else if (age < 50) annualMA = 600 + 0.2000 * (annualNTI - SEP_MA_BAND1_UPPER);
      else               annualMA = 630 + 0.2100 * (annualNTI - SEP_MA_BAND1_UPPER);
      effectiveRate = annualMA / annualNTI;
    } else {
      annualMA = Math.min(band3.rate * annualNTI, annualCap);
      effectiveRate = band3.rate;
    }
  }

  const atCeiling = annualNTI > SEP_MA_BAND2_UPPER && annualMA >= annualCap;

  return {
    annualMA,
    monthlyMA: annualMA / 12,
    effectiveRate,
    annualCap,
    aboveExempt: true,
    atCeiling,
    age,
    pensioner
  };
}

// owMonthly   – monthly Ordinary Wage (employment field, already monthly-normalised)
// awAnnual    – total Additional Wages for the year (bonus field, annual total)
// monthsWorked – number of months OW is paid (defaults to 12)
// dob         – date-of-birth string used to derive age and select rate tier
function calculateCPF(owMonthly, awAnnual, monthsWorked, dob) {
  // 1. Derive age from DOB
  const birthDate = dob ? new Date(dob) : null;
  const today = new Date();
  let age = 0;
  if (birthDate && !Number.isNaN(birthDate.valueOf())) {
    age = today.getFullYear() - birthDate.getFullYear();
    const monthDiff = today.getMonth() - birthDate.getMonth();
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
      age -= 1;
    }
  }

  // 2. Select rate tier
  const rates = CPF_RATE_TABLE.find(r => age <= r.maxAge) || CPF_RATE_TABLE[CPF_RATE_TABLE.length - 1];

  // 3. Ordinary Wages
  const owSubjectPerMonth = Math.min(owMonthly, MONTHLY_OW_CEILING);
  const annualOWSubject = owSubjectPerMonth * monthsWorked;

  // 4. Additional Wages ceiling and subject amount
  const awCeiling = Math.max(0, ANNUAL_OW_CEILING - annualOWSubject);
  const awSubject = Math.min(awAnnual, awCeiling);
  const awExcess = awAnnual - awSubject;

  // 5. Employee CPF on OW (monthly deduction from take-home)
  const monthlyEmployeeCPF = owSubjectPerMonth * rates.employee;

  // 6. Annual CPF – OW and AW combined
  const annualOWEmployeeCPF = annualOWSubject * rates.employee;
  const annualOWEmployerCPF = annualOWSubject * rates.employer;
  const awEmployeeCPF       = awSubject * rates.employee;
  const awEmployerCPF       = awSubject * rates.employer;

  const totalEmployeeCPF = annualOWEmployeeCPF + awEmployeeCPF;
  const totalEmployerCPF = annualOWEmployerCPF + awEmployerCPF;

  // 7. OA / SA / MA breakdown (employee portion only, as fraction of subject wages)
  const monthlyOA = owSubjectPerMonth * rates.oa;
  const monthlySA = owSubjectPerMonth * rates.sa;
  const monthlyMA = owSubjectPerMonth * rates.ma;

  const totalWageSubject = annualOWSubject + awSubject;
  const annualOA = totalWageSubject * rates.oa;
  const annualSA = totalWageSubject * rates.sa;
  const annualMA = totalWageSubject * rates.ma;

  const eeRatePct = Math.round(rates.employee * 100);
  const erRatePct = Math.round(rates.employer * 100);

  return {
    age,
    rateGroup: `Ee ${eeRatePct}% / Er ${erRatePct}%`,
    owSubjectPerMonth,
    annualOWSubject,
    awCeiling,
    awSubject,
    awExcess,
    bonusBreakdown: awAnnual > 0
      ? [{ amount: awAnnual, subject: awSubject, excess: awExcess }]
      : [],
    monthlyEmployeeCPF,
    monthlyOA,
    monthlySA,
    monthlyMA,
    monthlyTotal: monthlyEmployeeCPF,   // used for net monthly income
    totalEmployeeCPF,
    totalEmployerCPF,
    annualOA,
    annualSA,
    annualMA,
    annualTotal: totalEmployeeCPF,      // employee annual total
    annualTotalRaw: totalEmployeeCPF,
    monthlyCapped: owMonthly >= MONTHLY_OW_CEILING,
    annualCapped:  annualOWSubject >= ANNUAL_OW_CEILING,
  };
}

function valueById(id) {
  const el = document.getElementById(id);
  return el ? toNumber(el.value) : 0;
}

function sumContainerMonthly(containerId) {
  let total = 0;
  // Exclude [data-current] inputs (current portfolio value fields, not contributions)
  document.querySelectorAll(`#${containerId} input[type="number"]:not([data-current])`).forEach((el) => {
    total += getMonthlyValue(el.id, el.value);
  });
  return total;
}

function updateText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = fmt(value);
}

function updateDashboard() {
  const employment = getMonthlyValue("employment", valueById("employment"));
  const bonus = getMonthlyValue("bonus", valueById("bonus"));
  const otherIncome = getMonthlyValue("otherIncome", valueById("otherIncome"));
  const grossIncome = employment + bonus + otherIncome;

  const dob = document.getElementById("dob")?.value;
  let cpfResult = null;
  let sepResult  = null;
  let netAnnual, netIncome;

  if (isSelfEmployed) {
    // SEP: no standard CPF; only Medisave estimate
    const ageJan1 = getAgeAsAtJan1(dob);
    sepResult  = calculateSEPMedisave(grossIncome * 12, ageJan1, isPensioner);
    netAnnual  = grossIncome * 12 - sepResult.annualMA;
    netIncome  = netAnnual / 12;
  } else {
    // Employed: full CPF calculation
    // OW = employment monthly; AW = bonus converted to annual total
    const awAnnual = getMonthlyValue("bonus", valueById("bonus")) * 12;
    cpfResult = calculateCPF(employment, awAnnual, 12, dob);
    // Net must be computed annually so the AW CPF deduction (lump-sum, not monthly)
    // is correctly spread across the year alongside the monthly-averaged bonus income.
    netAnnual = grossIncome * 12 - cpfResult.totalEmployeeCPF;
    netIncome = netAnnual / 12;
  }

  const totalCPF = isSelfEmployed ? (sepResult?.monthlyMA || 0) : (cpfResult?.monthlyTotal || 0);

  const food = getMonthlyValue("food", valueById("food"));
  const apparel = getMonthlyValue("apparel", valueById("apparel"));
  const transport = getMonthlyValue("transport", valueById("transport"));
  const subscriptions = getMonthlyValue("subscriptions", valueById("subscriptions"));
  const bills = getMonthlyValue("bills", valueById("bills"));
  const allowances = getMonthlyValue("allowances", valueById("allowances"));
  const recreation = getMonthlyValue("recreation", valueById("recreation"));
  const holidays = getMonthlyValue("holidays", valueById("holidays"));
  const additionalPersonalMonthly = sumContainerMonthly("additionalPersonal");
  const totalPersonal =
    food + apparel + transport + subscriptions + bills + allowances + recreation + holidays + additionalPersonalMonthly;

  const loanHouse = getMonthlyValue("loanHouse", valueById("loanHouse"));
  const loanCar = getMonthlyValue("loanCar", valueById("loanCar"));
  const loanEducation = getMonthlyValue("loanEducation", valueById("loanEducation"));
  const loanBusiness = getMonthlyValue("loanBusiness", valueById("loanBusiness"));
  const loanPersonal = getMonthlyValue("loanPersonal", valueById("loanPersonal"));
  const additionalLoansMonthly = sumContainerMonthly("additionalLoans");
  const totalLoans = loanHouse + loanCar + loanEducation + loanBusiness + loanPersonal + additionalLoansMonthly;

  const totalExpenses = totalPersonal + totalLoans;

  const hosp = getMonthlyValue("hospPremium", valueById("hospPremium"));
  const whole = getMonthlyValue("wholePremium", valueById("wholePremium"));
  const term = getMonthlyValue("termPremium", valueById("termPremium"));
  const ci = getMonthlyValue("ciPremium", valueById("ciPremium"));
  const additionalInsuranceMonthly = sumContainerMonthly("additionalInsurance");
  const totalInsurance = hosp + whole + term + ci + additionalInsuranceMonthly;

  const endo = getMonthlyValue("endowment", valueById("endowment"));
  const inv = getMonthlyValue("investments", valueById("investments"));
  const fd = getMonthlyValue("fixedDeposits", valueById("fixedDeposits"));
  const stk = getMonthlyValue("stocks", valueById("stocks"));
  const oth = getMonthlyValue("otherSavings", valueById("otherSavings"));
  const additionalSavingsMonthly = sumContainerMonthly("additionalSavings");
  const totalSavings = endo + inv + fd + stk + oth + additionalSavingsMonthly;

  const endowmentCurrent = valueById("endowmentCurrent");
  const investmentsCurrent = valueById("investmentsCurrent");
  const liquidCashCurrent = valueById("liquidCashCurrent");
  const fixedDepositsCurrent = valueById("fixedDepositsCurrent");
  const stocksCurrent = valueById("stocksCurrent");
  const otherSavingsCurrent = valueById("otherSavingsCurrent");
  const additionalSavingsCurrentValue = (() => {
    let total = 0;
    document.querySelectorAll('#additionalSavings input[data-current="true"]').forEach((el) => {
      total += toNumber(el.value);
    });
    return total;
  })();

  const portfolioCurrentValue =
    endowmentCurrent + investmentsCurrent + fixedDepositsCurrent + stocksCurrent + otherSavingsCurrent + additionalSavingsCurrentValue;

  const netCashflow = netIncome - totalExpenses - totalInsurance - totalSavings;
  const savingsRate = netIncome ? (totalSavings / netIncome) * 100 : 0;

  updateText("grossIncomeM", grossIncome);
  updateText("grossIncomeA", grossIncome * 12);
  updateText("netIncomeM", netIncome);
  updateText("netIncomeA", netIncome * 12);

  if (!isSelfEmployed && cpfResult) {
    // ── Employed CPF display ──────────────────────────────────────────
    const cpfRateGroupEl = document.getElementById("cpfRateGroup");
    if (cpfRateGroupEl) cpfRateGroupEl.textContent = cpfResult.rateGroup;

    updateText("cpfOAM", cpfResult.monthlyOA);
    updateText("cpfOAA", cpfResult.annualOA);
    updateText("cpfSAM", cpfResult.monthlySA);
    updateText("cpfSAA", cpfResult.annualSA);
    updateText("cpfMAM", cpfResult.monthlyMA);
    updateText("cpfMAA", cpfResult.annualMA);
    updateText("totalCPFM", cpfResult.monthlyTotal);
    updateText("totalCPFA", cpfResult.annualTotal);

    const cpfNoteSection = document.getElementById("cpfNoteSection");
    const cpfNote = document.getElementById("cpfNote");
    if (cpfNoteSection && cpfNote) {
      const noteParts = [];
      if (cpfResult.monthlyCapped) {
        noteParts.push(`OW capped at S$${MONTHLY_OW_CEILING.toLocaleString()}/month for CPF.`);
      }
      if (cpfResult.annualCapped) {
        noteParts.push(`Annual OW ceiling of S$${ANNUAL_OW_CEILING.toLocaleString()} reached.`);
      }
      if (cpfResult.awExcess > 0) {
        noteParts.push(`S${fmt(cpfResult.awExcess)} of bonus exceeds AW ceiling — not subject to CPF.`);
      }
      if (noteParts.length > 0) {
        cpfNoteSection.style.display = "block";
        cpfNote.textContent = noteParts.join(" ");
      } else {
        cpfNoteSection.style.display = "none";
      }
    }

  } else if (isSelfEmployed && sepResult) {
    // ── SEP Medisave display ──────────────────────────────────────────
    updateText("sepMAM",    sepResult.monthlyMA);
    updateText("sepMAA",    sepResult.annualMA);
    updateText("sepTotalM", sepResult.monthlyMA);
    updateText("sepTotalA", sepResult.annualMA);

    const sepCaption = document.getElementById("sepCaption");
    if (sepCaption) {
      const ntiAnnual  = grossIncome * 12;
      const ratePct    = (sepResult.effectiveRate * 100).toFixed(2);
      const tableLabel = isPensioner ? "Table 2 — Pensioner" : "Table 1 — Non-Pensioner";
      // Determine which NTI band the client falls in
      let band;
      if (ntiAnnual <= SEP_MA_EXEMPT)       band = "Below exemption threshold";
      else if (ntiAnnual <= SEP_MA_BAND1_UPPER) band = `Band 1 ($6,001–$12,000)`;
      else if (ntiAnnual <= SEP_MA_BAND2_UPPER) band = `Band 2 ($12,001–$18,000, phase-in)`;
      else                                   band = `Band 3 (above $18,000)`;

      sepCaption.textContent = [
        `NTI est.: ${fmt(ntiAnnual)} p.a.`,
        `Age (as at 1 Jan): ${sepResult.age}`,
        band,
        `Eff. rate: ${ratePct}%`,
        tableLabel
      ].join(" · ");
    }

    const sepInfoNote = document.getElementById("sepInfoNote");
    if (sepInfoNote) {
      const ntiAnnual = grossIncome * 12;
      if (ntiAnnual <= SEP_MA_EXEMPT) {
        sepInfoNote.style.display = "block";
        sepInfoNote.innerHTML = `<strong>No Contribution Required:</strong> Estimated NTI of ${fmt(ntiAnnual)} p.a. is at or below the $6,000 exemption threshold.`;
      } else if (sepResult.atCeiling) {
        sepInfoNote.style.display = "block";
        sepInfoNote.innerHTML = `<strong>Maximum Contribution Reached:</strong> Annual Medisave contribution is capped at ${fmt(sepResult.annualCap)}.`;
      } else {
        sepInfoNote.style.display = "none";
      }
    }
  }

  const totalCPFBalance = valueById("cpfOABalance") + valueById("cpfSABalance") + valueById("cpfMABalance");
  updateText("totalCPFBalance", totalCPFBalance);

  updateText("totalExpensesM", totalExpenses);
  updateText("totalExpensesA", totalExpenses * 12);
  updateText("totalInsuranceM", totalInsurance);
  updateText("totalInsuranceA", totalInsurance * 12);
  updateText("totalSavingsM", totalSavings);
  updateText("totalSavingsA", totalSavings * 12);

  updateText("totalLiquidCash", liquidCashCurrent);
  updateText("totalPortfolioCurrent", portfolioCurrentValue);

  updateText("summaryNetIncomeM", netIncome);
  updateText("summaryNetIncomeY", netIncome * 12);
  updateText("summaryInsuranceM", totalInsurance);
  updateText("summaryInsuranceY", totalInsurance * 12);
  updateText("summarySavingsM", totalSavings);
  updateText("summarySavingsY", totalSavings * 12);
  updateText("summaryExpensesM", totalExpenses);
  updateText("summaryExpensesY", totalExpenses * 12);
  updateText("summaryNetCashflowM", netCashflow);
  updateText("summaryNetCashflowY", netCashflow * 12);

  updateExpensesChart(
    food,
    apparel,
    transport,
    subscriptions,
    bills,
    allowances,
    recreation,
    holidays,
    totalInsurance,
    totalSavings,
    netCashflow,
    additionalPersonalMonthly,
    totalLoans
  );
  updateSavingsProjectionChart(netCashflow, liquidCashCurrent);

  latestMetrics = {
    grossIncome,
    netIncome,
    isSelfEmployed,
    totalCPF,
    totalExpenses,
    totalInsurance,
    totalSavings,
    liquidCashCurrent,
    portfolioCurrentValue,
    netCashflow,
    savingsRate,
    cpf: cpfResult,
    sep: sepResult,
    details: {
      income: {
        employment,
        bonus,
        otherIncome
      },
      expenses: {
        food,
        apparel,
        transport,
        subscriptions,
        bills,
        allowances,
        recreation,
        holidays,
        loans: totalLoans,
        personal: totalPersonal,
        additionalPersonalMonthly,
        additionalLoansMonthly
      },
      insurance: {
        hosp,
        whole,
        term,
        ci,
        additionalInsuranceMonthly
      },
      savings: {
        endo,
        inv,
        fd,
        stk,
        oth,
        additionalSavingsMonthly,
        endowmentCurrent,
        investmentsCurrent,
        fixedDepositsCurrent,
        stocksCurrent,
        otherSavingsCurrent
      }
    }
  };
}

function updateExpensesChart(
  food,
  apparel,
  transport,
  subscriptions,
  bills,
  allowances,
  recreation,
  holidays,
  totalInsurance,
  totalSavings,
  netCashflow,
  otherExpenses,
  totalLoans
) {
  if (!window.Chart) return;
  const canvas = document.getElementById("expensesPieChart");
  if (!canvas) return;

  if (expensesPieChart) expensesPieChart.destroy();

  const baseValues = [
    food,
    apparel,
    transport,
    subscriptions,
    bills,
    allowances,
    recreation,
    holidays,
    totalInsurance,
    totalSavings
  ];
  const baseLabels = [
    "Food",
    "Apparel",
    "Transport",
    "Subscriptions",
    "Bills",
    "Allowances",
    "Recreation",
    "Holidays",
    "Insurance",
    "Savings/Investments"
  ];
  const baseColors = [
    "#FFBCB5", /* Food          – pastel coral   */
    "#FFD9B5", /* Apparel        – pastel peach   */
    "#FFF3B5", /* Transport      – pastel yellow  */
    "#C9F0C5", /* Subscriptions  – pastel green   */
    "#B5E0F0", /* Bills          – pastel sky     */
    "#B5C8F0", /* Allowances     – pastel blue    */
    "#CCBCF0", /* Recreation     – pastel violet  */
    "#F0BCDC", /* Holidays       – pastel pink    */
    "#a8c8a0", /* Insurance      – sage green     */
    "#a0c4e8"  /* Savings        – pastel blue    */
  ];

  // Include optional slices only when they have a value
  const dataValues = [...baseValues];
  const labels     = [...baseLabels];
  const colors     = [...baseColors];
  if (otherExpenses > 0) {
    dataValues.push(otherExpenses);
    labels.push("Other Expenses");
    colors.push("#d4b5f0"); /* Other Expenses – pastel purple */
  }
  if (totalLoans > 0) {
    dataValues.push(totalLoans);
    labels.push("Loans/Liabilities");
    colors.push("#f0c8a0"); /* Loans/Liabilities – pastel orange */
  }
  if (netCashflow > 0) {
    dataValues.push(netCashflow);
    labels.push("Net Cashflow");
    colors.push("#f5d78e"); /* Net Cashflow – warm gold */
  }

  const total = dataValues.reduce((sum, n) => sum + n, 0) || 1;

  expensesPieChart = new Chart(canvas.getContext("2d"), {
    type: "doughnut",
    data: {
      labels,
      datasets: [
        {
          data: dataValues,
          backgroundColor: colors,
          borderColor: "#fff",
          borderWidth: 2
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            font: {
              family: "Source Sans 3",
              size: 11
            },
            usePointStyle: true,
            padding: 12
          }
        },
        tooltip: {
          callbacks: {
            label(context) {
              const value = Number(context.parsed) || 0;
              const percentage = ((value / total) * 100).toFixed(1);
              return `${context.label}: ${fmt(value)} (${percentage}%)`;
            }
          }
        }
      }
    }
  });
}

function updateSavingsProjectionChart(monthlySurplus, currentLiquidCash) {
  if (!window.Chart) return;
  const canvas = document.getElementById("savingsProjectionChart");
  if (!canvas) return;

  if (savingsProjectionChart) savingsProjectionChart.destroy();

  const months = [];
  const projectedSavings = [];
  let running = currentLiquidCash;

  for (let i = 0; i < 12; i += 1) {
    const date = new Date(new Date().getFullYear(), new Date().getMonth() + i, 1);
    months.push(date.toLocaleString("default", { month: "short" }));
    running += monthlySurplus;
    projectedSavings.push(Math.round(running));
  }

  savingsProjectionChart = new Chart(canvas.getContext("2d"), {
    type: "line",
    data: {
      labels: months,
      datasets: [
        {
          label: "Projected Savings (from Net Cashflow)",
          data: projectedSavings,
          borderColor: "#7abfa0",
          backgroundColor: "rgba(122, 191, 160, 0.15)",
          borderWidth: 3,
          fill: true,
          tension: 0.35,
          pointRadius: 4,
          pointBackgroundColor: "#7abfa0",
          pointBorderColor: "#fff",
          pointBorderWidth: 2
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          labels: {
            font: {
              family: "Source Sans 3",
              size: 11
            },
            usePointStyle: true
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: {
            callback(value) {
              return "$" + Number(value).toLocaleString();
            },
            font: {
              family: "Source Sans 3"
            }
          }
        },
        x: {
          ticks: {
            font: {
              family: "Source Sans 3"
            }
          }
        }
      }
    }
  });
}

function createDynamicRow({ idPrefix, title, notesPlaceholder }) {
  const id = `${idPrefix}${Date.now()}`;
  const wrapper = document.createElement("div");
  wrapper.className = "item-row";
  wrapper.innerHTML = `
    <div class="item-label">${title}</div>
    <div class="input-group">
      <button type="button" class="toggle-btn active" data-field="${id}" data-type="Monthly">Monthly</button>
      <button type="button" class="toggle-btn" data-field="${id}" data-type="Annual">Annual</button>
      <input type="number" id="${id}" class="item-input" value="0" />
    </div>
    <textarea class="item-note" placeholder="${notesPlaceholder}"></textarea>
    <button type="button" class="remove-btn" onclick="this.closest('.item-row').remove(); updateDashboard();">Remove</button>
  `;
  return wrapper;
}

function addPersonalExpense() {
  const name = prompt("Expense name:", "Other Personal Expense");
  if (!name) return;
  document
    .getElementById("additionalPersonal")
    ?.appendChild(createDynamicRow({ idPrefix: "persExp", title: name, notesPlaceholder: "Details" }));
  updateDashboard();
}

function addLoan() {
  const name = prompt("Loan/Liability name:", "Other Liability");
  if (!name) return;
  document
    .getElementById("additionalLoans")
    ?.appendChild(createDynamicRow({ idPrefix: "loan", title: name, notesPlaceholder: "Loan details" }));
  updateDashboard();
}

function addInsurance() {
  const name = prompt("Plan name:", "Other Plan");
  if (!name) return;
  document
    .getElementById("additionalInsurance")
    ?.appendChild(createDynamicRow({ idPrefix: "ins", title: name, notesPlaceholder: "Plan details" }));
  updateDashboard();
}

function createSavingsDynamicRow(title) {
  const id        = `sav${Date.now()}`;
  const currentId = `${id}Current`;
  const row       = document.createElement("tr");
  row.className   = "item-row";
  row.innerHTML   = `
    <td>
      <div class="item-label">${title}</div>
      <div class="input-group tight">
        <button type="button" class="toggle-btn active" data-field="${id}" data-type="Monthly">Monthly</button>
        <button type="button" class="toggle-btn" data-field="${id}" data-type="Annual">Annual</button>
      </div>
      <input type="number" id="${id}" class="item-input" value="" />
    </td>
    <td>
      <input type="number" id="${currentId}" class="item-input" data-current="true" placeholder="0" value="" />
    </td>
    <td>
      <textarea class="item-note" placeholder="Savings details"></textarea>
      <button type="button" class="remove-btn" onclick="this.closest('tr').remove(); updateDashboard();">Remove</button>
    </td>
  `;
  return row;
}

function addSavings() {
  const name = prompt("Savings type:", "Other Savings");
  if (!name) return;
  document
    .getElementById("additionalSavings")
    ?.appendChild(createSavingsDynamicRow(name));
  updateDashboard();
}

function getFieldMeta(fieldId) {
  const input = document.getElementById(fieldId);
  if (!input) {
    return {
      id: fieldId,
      label: fieldId,
      rawText: "",
      rawValue: 0,
      frequency: "Monthly",
      monthly: 0,
      annual: 0,
      note: ""
    };
  }

  const row = input.closest(".item-row, tr");
  const note = row?.querySelector("textarea")?.value?.trim() || "";
  const label = row?.querySelector(".item-label")?.textContent?.trim() || fieldId;
  const rawText = String(input.value || "").trim();
  const rawValue = toNumber(rawText);
  const frequency = inputTypes[fieldId] || "Monthly";
  const monthly = getMonthlyValue(fieldId, rawText);

  return {
    id: fieldId,
    label,
    rawText,
    rawValue,
    frequency,
    monthly,
    annual: monthly * 12,
    note
  };
}

function collectDynamicRows(containerId) {
  const rows = [];
  document.querySelectorAll(`#${containerId} .item-row`).forEach((row) => {
    const input = row.querySelector('input[type="number"]');
    if (!input) return;

    const label = row.querySelector(".item-label")?.textContent?.trim() || input.id;
    const rawText = String(input.value || "").trim();
    const rawValue = toNumber(rawText);
    const frequency = inputTypes[input.id] || "Monthly";
    const monthly = getMonthlyValue(input.id, rawText);
    const note = row.querySelector("textarea")?.value?.trim() || "";

    rows.push({
      id: input.id,
      label,
      rawText,
      rawValue,
      frequency,
      monthly,
      annual: monthly * 12,
      note
    });
  });

  return rows;
}

function summarizeDynamicItemNotes(items) {
  return items
    .filter((item) => item.note)
    .map((item) => `${item.label}: ${item.note}`)
    .join(" | ");
}

function remarkForField(meta) {
  return meta.note || "";
}

function decodeBase64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function roundToTwo(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function writeCell(ws, address, value, options = {}) {
  const cell = ws.getCell(address);
  const numFmt = options.numFmt || "#,##0.00";
  const applyNumFmt = options.applyNumFmt !== false;

  if (value === null || value === undefined || value === "") {
    cell.value = null;
    return;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    cell.value = roundToTwo(value);
    if (applyNumFmt) cell.numFmt = numFmt;
    return;
  }

  cell.value = value;
}

async function loadTemplateWorkbook() {
  if (!window.ExcelJS) throw new Error("ExcelJS library is not loaded");

  const workbook = new window.ExcelJS.Workbook();

  if (window.CASHFLOW_TEMPLATE_BASE64) {
    const arrayBuffer = decodeBase64ToArrayBuffer(window.CASHFLOW_TEMPLATE_BASE64);
    await workbook.xlsx.load(arrayBuffer);
    return workbook;
  }

  const response = await fetch(encodeURI(TEMPLATE_FILE));
  if (!response.ok) throw new Error(`Template fetch failed with status ${response.status}`);
  const arrayBuffer = await response.arrayBuffer();
  await workbook.xlsx.load(arrayBuffer);
  return workbook;
}

function addCommentsSheet(workbook, commentsRows) {
  const existing = workbook.getWorksheet("Comments");
  if (existing) workbook.removeWorksheet(existing.id);

  const commentsSheet = workbook.addWorksheet("Comments");
  commentsRows.forEach((row) => commentsSheet.addRow(row));
  commentsSheet.columns = [
    { width: 20 },
    { width: 28 },
    { width: 12 },
    { width: 14 },
    { width: 14 },
    { width: 14 },
    { width: 60 }
  ];
  commentsSheet.getRow(1).font = { bold: true };
}

function downloadArrayBufferAsFile(buffer, filename) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function exportToExcel() {
  if (!window.ExcelJS) {
    alert("Excel export library not loaded.");
    return;
  }

  try {
    if (!latestMetrics) updateDashboard();

    const clientName = document.getElementById("clientName")?.value?.trim() || "Client";
    const dob = document.getElementById("dob")?.value?.trim() || "";
    const consultantName = document.getElementById("consultantName")?.value?.trim() || "";
    const presentationDate = document.getElementById("presentationDate")?.value?.trim() || new Date().toISOString().split("T")[0];

    const employmentMeta = getFieldMeta("employment");
    const bonusMeta = getFieldMeta("bonus");
    const otherIncomeMeta = getFieldMeta("otherIncome");

    const foodMeta = getFieldMeta("food");
    const apparelMeta = getFieldMeta("apparel");
    const transportMeta = getFieldMeta("transport");
    const subscriptionsMeta = getFieldMeta("subscriptions");
    const billsMeta = getFieldMeta("bills");
    const allowancesMeta = getFieldMeta("allowances");
    const recreationMeta = getFieldMeta("recreation");
    const holidaysMeta = getFieldMeta("holidays");

    const loanHouseMeta = getFieldMeta("loanHouse");
    const loanCarMeta = getFieldMeta("loanCar");
    const loanEducationMeta = getFieldMeta("loanEducation");
    const loanBusinessMeta = getFieldMeta("loanBusiness");
    const loanPersonalMeta = getFieldMeta("loanPersonal");

    const hospMeta = getFieldMeta("hospPremium");
    const termMeta = getFieldMeta("termPremium");
    const wholeMeta = getFieldMeta("wholePremium");
    const ciMeta = getFieldMeta("ciPremium");

    const endowmentMeta = getFieldMeta("endowment");
    const investmentsMeta = getFieldMeta("investments");
    const fixedDepositsMeta = getFieldMeta("fixedDeposits");
    const stocksMeta = getFieldMeta("stocks");
    const otherSavingsMeta = getFieldMeta("otherSavings");

    const endowmentCurrentMeta = getFieldMeta("endowmentCurrent");
    const investmentsCurrentMeta = getFieldMeta("investmentsCurrent");
    const liquidCashCurrentMeta = getFieldMeta("liquidCashCurrent");
    const fixedDepositsCurrentMeta = getFieldMeta("fixedDepositsCurrent");
    const stocksCurrentMeta = getFieldMeta("stocksCurrent");
    const otherSavingsCurrentMeta = getFieldMeta("otherSavingsCurrent");

    const additionalPersonal = collectDynamicRows("additionalPersonal");
    const additionalLoans = collectDynamicRows("additionalLoans");
    const additionalInsurance = collectDynamicRows("additionalInsurance");
    const additionalSavings = collectDynamicRows("additionalSavings");

    const workbook = await loadTemplateWorkbook();
    const ws = workbook.worksheets[0] || workbook.addWorksheet("Sheet1");

    const additionalPersonalMonthly = additionalPersonal.reduce((sum, item) => sum + item.monthly, 0);
    const additionalLoansMonthly = additionalLoans.reduce((sum, item) => sum + item.monthly, 0);
    const additionalInsuranceMonthly = additionalInsurance.reduce((sum, item) => sum + item.monthly, 0);
    const additionalSavingsMonthly = additionalSavings.reduce((sum, item) => sum + item.monthly, 0);

    const insuranceOthersMonthly = ciMeta.monthly + additionalInsuranceMonthly;
    const savingsOthersMonthly = otherSavingsMeta.monthly + additionalSavingsMonthly;

    writeCell(ws, "A1", `${clientName}'s Cashflow`);

    if (dob) {
      const [y, m, d] = dob.split("-");
      const dobFormatted = new Date(+y, +m - 1, +d)
        .toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
      writeCell(ws, "E3", `DOB: ${dobFormatted}`);
    }

    writeCell(ws, "B6", employmentMeta.monthly);
    writeCell(ws, "C6", employmentMeta.annual);
    writeCell(ws, "D6", remarkForField(employmentMeta));

    writeCell(ws, "B7", bonusMeta.monthly);
    writeCell(ws, "C7", bonusMeta.annual);
    writeCell(ws, "D7", remarkForField(bonusMeta));

    writeCell(ws, "B8", otherIncomeMeta.monthly);
    writeCell(ws, "C8", otherIncomeMeta.annual);
    writeCell(ws, "D8", remarkForField(otherIncomeMeta));

    if (latestMetrics.isSelfEmployed && latestMetrics.sep) {
      // SEP mode: OA and SA are zero; MA carries the estimated Medisave contribution
      writeCell(ws, "B10", 0);
      writeCell(ws, "C10", valueById("cpfOABalance"));

      writeCell(ws, "B11", 0);
      writeCell(ws, "C11", valueById("cpfSABalance"));

      writeCell(ws, "B12", latestMetrics.sep.monthlyMA);
      writeCell(ws, "C12", valueById("cpfMABalance"));
    } else if (latestMetrics.cpf) {
      writeCell(ws, "B10", latestMetrics.cpf.monthlyOA);
      writeCell(ws, "C10", valueById("cpfOABalance"));

      writeCell(ws, "B11", latestMetrics.cpf.monthlySA);
      writeCell(ws, "C11", valueById("cpfSABalance"));

      writeCell(ws, "B12", latestMetrics.cpf.monthlyMA);
      writeCell(ws, "C12", valueById("cpfMABalance"));
    }

    writeCell(ws, "B13", latestMetrics.netIncome);
    writeCell(ws, "C13", latestMetrics.netIncome * 12);
    writeCell(ws, "D13", "");

    writeCell(ws, "G6", foodMeta.monthly);
    writeCell(ws, "H6", foodMeta.annual);
    writeCell(ws, "I6", remarkForField(foodMeta));

    writeCell(ws, "G7", apparelMeta.monthly);
    writeCell(ws, "H7", apparelMeta.annual);
    writeCell(ws, "I7", remarkForField(apparelMeta));

    writeCell(ws, "G8", transportMeta.monthly);
    writeCell(ws, "H8", transportMeta.annual);
    writeCell(ws, "I8", remarkForField(transportMeta));

    writeCell(ws, "G9", subscriptionsMeta.monthly);
    writeCell(ws, "H9", subscriptionsMeta.annual);
    writeCell(ws, "I9", remarkForField(subscriptionsMeta));

    writeCell(ws, "G10", billsMeta.monthly);
    writeCell(ws, "H10", billsMeta.annual);
    writeCell(ws, "I10", remarkForField(billsMeta));

    writeCell(ws, "G11", allowancesMeta.monthly);
    writeCell(ws, "H11", allowancesMeta.annual);
    writeCell(ws, "I11", remarkForField(allowancesMeta));

    writeCell(ws, "G12", recreationMeta.monthly);
    writeCell(ws, "H12", recreationMeta.annual);
    writeCell(ws, "I12", remarkForField(recreationMeta));

    writeCell(ws, "G13", holidaysMeta.monthly);
    writeCell(ws, "H13", holidaysMeta.annual);
    writeCell(ws, "I13", remarkForField(holidaysMeta));

    writeCell(ws, "G14", additionalPersonalMonthly);
    writeCell(ws, "H14", additionalPersonalMonthly * 12);
    writeCell(ws, "I14", summarizeDynamicItemNotes(additionalPersonal));

    writeCell(ws, "B17", hospMeta.monthly);
    writeCell(ws, "C17", hospMeta.annual);
    writeCell(ws, "D17", remarkForField(hospMeta));

    writeCell(ws, "B18", termMeta.monthly);
    writeCell(ws, "C18", termMeta.annual);
    writeCell(ws, "D18", remarkForField(termMeta));

    writeCell(ws, "B19", wholeMeta.monthly);
    writeCell(ws, "C19", wholeMeta.annual);
    writeCell(ws, "D19", remarkForField(wholeMeta));

    writeCell(ws, "B20", insuranceOthersMonthly);
    writeCell(ws, "C20", insuranceOthersMonthly * 12);
    writeCell(
      ws,
      "D20",
      [
        ciMeta.note,
        summarizeDynamicItemNotes(additionalInsurance)
      ]
        .filter(Boolean)
        .join(" | ")
    );

    writeCell(ws, "B21", latestMetrics.totalInsurance);
    writeCell(ws, "C21", latestMetrics.totalInsurance * 12);

    writeCell(ws, "G16", loanHouseMeta.monthly);
    writeCell(ws, "H16", loanHouseMeta.annual);
    writeCell(ws, "I16", remarkForField(loanHouseMeta));

    writeCell(ws, "G17", loanCarMeta.monthly);
    writeCell(ws, "H17", loanCarMeta.annual);
    writeCell(ws, "I17", remarkForField(loanCarMeta));

    writeCell(ws, "G18", loanEducationMeta.monthly);
    writeCell(ws, "H18", loanEducationMeta.annual);
    writeCell(ws, "I18", remarkForField(loanEducationMeta));

    writeCell(ws, "G19", loanBusinessMeta.monthly);
    writeCell(ws, "H19", loanBusinessMeta.annual);
    writeCell(ws, "I19", remarkForField(loanBusinessMeta));

    writeCell(ws, "G20", loanPersonalMeta.monthly);
    writeCell(ws, "H20", loanPersonalMeta.annual);
    writeCell(ws, "I20", remarkForField(loanPersonalMeta));

    writeCell(ws, "G21", additionalLoansMonthly);
    writeCell(ws, "H21", additionalLoansMonthly * 12);
    writeCell(ws, "I21", summarizeDynamicItemNotes(additionalLoans));

    writeCell(ws, "G22", latestMetrics.totalExpenses);
    writeCell(ws, "H22", latestMetrics.totalExpenses * 12);

    writeCell(ws, "B26", endowmentMeta.monthly);
    writeCell(ws, "C26", endowmentCurrentMeta.rawValue);
    writeCell(ws, "D26", remarkForField(endowmentMeta));

    writeCell(ws, "B27", investmentsMeta.monthly);
    writeCell(ws, "C27", investmentsCurrentMeta.rawValue);
    writeCell(ws, "D27", remarkForField(investmentsMeta));

    writeCell(ws, "B28", 0);
    writeCell(ws, "C28", liquidCashCurrentMeta.rawValue);
    writeCell(ws, "D28", liquidCashCurrentMeta.note || "");

    writeCell(ws, "B29", fixedDepositsMeta.monthly);
    writeCell(ws, "C29", fixedDepositsCurrentMeta.rawValue);
    writeCell(ws, "D29", remarkForField(fixedDepositsMeta));

    writeCell(ws, "B30", stocksMeta.monthly);
    writeCell(ws, "C30", stocksCurrentMeta.rawValue);
    writeCell(ws, "D30", remarkForField(stocksMeta));

    writeCell(ws, "B31", savingsOthersMonthly);
    writeCell(ws, "C31", otherSavingsCurrentMeta.rawValue);
    writeCell(
      ws,
      "D31",
      [remarkForField(otherSavingsMeta), summarizeDynamicItemNotes(additionalSavings)].filter(Boolean).join(" | ")
    );

    writeCell(ws, "B32", latestMetrics.totalSavings);
    writeCell(
      ws,
      "C32",
      endowmentCurrentMeta.rawValue +
        investmentsCurrentMeta.rawValue +
        liquidCashCurrentMeta.rawValue +
        fixedDepositsCurrentMeta.rawValue +
        stocksCurrentMeta.rawValue +
        otherSavingsCurrentMeta.rawValue
    );

    writeCell(ws, "G26", latestMetrics.netIncome);
    writeCell(ws, "H26", latestMetrics.netIncome * 12);

    writeCell(ws, "G27", latestMetrics.totalInsurance);
    writeCell(ws, "H27", latestMetrics.totalInsurance * 12);

    writeCell(ws, "G28", latestMetrics.totalSavings);
    writeCell(ws, "H28", latestMetrics.totalSavings * 12);

    writeCell(ws, "G29", latestMetrics.totalExpenses);
    writeCell(ws, "H29", latestMetrics.totalExpenses * 12);

    writeCell(ws, "G30", latestMetrics.netCashflow);
    writeCell(ws, "H30", latestMetrics.netCashflow * 12);

    const sepModeLabel = latestMetrics.isSelfEmployed
      ? `Self-Employed (SEP) — Medisave est. ${fmt(latestMetrics.sep?.annualMA || 0)} p.a. · Eff. rate: ${((latestMetrics.sep?.effectiveRate || 0) * 100).toFixed(2)}% · ${latestMetrics.sep?.pensioner ? "Pensioner (Table 2)" : "Non-Pensioner (Table 1)"} · CPF MARates2026`
      : "Employed";

    const commentsRows = [
      ["Category", "Item", "Frequency", "Raw Input", "Monthly Value", "Annual Value", "Comment"],
      ["Client", "Client Name", "-", "-", "-", "-", clientName],
      ["Client", "Date of Birth", "-", "-", "-", "-", dob],
      ["Client", "Consultant", "-", "-", "-", "-", consultantName],
      ["Client", "Presentation Date", "-", "-", "-", "-", presentationDate],
      ["Client", "Employment Status", "-", "-", "-", "-", sepModeLabel]
    ];

    function addCommentRow(category, meta) {
      commentsRows.push([
        category,
        meta.label,
        meta.frequency,
        meta.rawText || "",
        Number.isFinite(meta.monthly) ? meta.monthly : "",
        Number.isFinite(meta.annual) ? meta.annual : "",
        meta.note || ""
      ]);
    }

    [employmentMeta, bonusMeta, otherIncomeMeta].forEach((m) => addCommentRow("Income", m));
    [foodMeta, apparelMeta, transportMeta, subscriptionsMeta, billsMeta, allowancesMeta, recreationMeta, holidaysMeta].forEach((m) =>
      addCommentRow("Personal Expense", m)
    );
    [loanHouseMeta, loanCarMeta, loanEducationMeta, loanBusinessMeta, loanPersonalMeta].forEach((m) =>
      addCommentRow("Loan Liability", m)
    );
    [hospMeta, termMeta, wholeMeta, ciMeta].forEach((m) => addCommentRow("Insurance", m));
    [endowmentMeta, investmentsMeta, fixedDepositsMeta, stocksMeta, otherSavingsMeta].forEach((m) =>
      addCommentRow("Savings Contribution", m)
    );

    [
      { category: "Savings Current", meta: endowmentCurrentMeta },
      { category: "Savings Current", meta: investmentsCurrentMeta },
      { category: "Savings Current", meta: liquidCashCurrentMeta },
      { category: "Savings Current", meta: fixedDepositsCurrentMeta },
      { category: "Savings Current", meta: stocksCurrentMeta },
      { category: "Savings Current", meta: otherSavingsCurrentMeta }
    ].forEach(({ category, meta }) => {
      commentsRows.push([category, meta.label, "Current", meta.rawText || "", "", "", meta.note || ""]);
    });

    additionalPersonal.forEach((m) => addCommentRow("Additional Personal", m));
    additionalLoans.forEach((m) => addCommentRow("Additional Loan", m));
    additionalInsurance.forEach((m) => addCommentRow("Additional Insurance", m));
    additionalSavings.forEach((m) => addCommentRow("Additional Savings", m));

    commentsRows.push(["Summary", "Net Income", "Monthly", "", latestMetrics.netIncome, latestMetrics.netIncome * 12, ""]);
    commentsRows.push([
      "Summary",
      "Total Expenses",
      "Monthly",
      "",
      latestMetrics.totalExpenses,
      latestMetrics.totalExpenses * 12,
      ""
    ]);
    commentsRows.push([
      "Summary",
      "Total Insurance",
      "Monthly",
      "",
      latestMetrics.totalInsurance,
      latestMetrics.totalInsurance * 12,
      ""
    ]);
    commentsRows.push([
      "Summary",
      "Total Savings",
      "Monthly",
      "",
      latestMetrics.totalSavings,
      latestMetrics.totalSavings * 12,
      ""
    ]);
    commentsRows.push([
      "Summary",
      "Net Cashflow",
      "Monthly",
      "",
      latestMetrics.netCashflow,
      latestMetrics.netCashflow * 12,
      `Savings rate: ${latestMetrics.savingsRate.toFixed(2)}%`
    ]);

    addCommentsSheet(workbook, commentsRows);

    const filename = `Cashflow_${clientName.replace(/[^a-z0-9]/gi, "_")}_${presentationDate}.xlsx`;
    const output = await workbook.xlsx.writeBuffer();
    downloadArrayBufferAsFile(output, filename);
    alert("Excel file exported using template layout.");
  } catch (error) {
    console.error("Excel export error:", error);
    alert("Error exporting Excel file.");
  }
}

// ─── Employment mode toggles ─────────────────────────────────────────────────

function setEmploymentMode(selfEmployed) {
  isSelfEmployed = selfEmployed;

  const btnEmployed     = document.getElementById("btnEmployed");
  const btnSelfEmployed = document.getElementById("btnSelfEmployed");
  const cpfEmployedView = document.getElementById("cpfEmployedView");
  const cpfSEPView      = document.getElementById("cpfSEPView");
  const cpfPanelTitle   = document.getElementById("cpfPanelTitle");

  if (btnEmployed)     btnEmployed.classList.toggle("active", !selfEmployed);
  if (btnSelfEmployed) btnSelfEmployed.classList.toggle("active", selfEmployed);
  if (cpfEmployedView) cpfEmployedView.style.display = selfEmployed ? "none"  : "block";
  if (cpfSEPView)      cpfSEPView.style.display      = selfEmployed ? "block" : "none";
  if (cpfPanelTitle)   cpfPanelTitle.textContent      = selfEmployed
    ? "CPF (Self-Employed)"
    : "CPF (Employee)";

  updateDashboard();
}

function setPensionerMode(pensioner) {
  isPensioner = pensioner;

  const btnNonPensioner    = document.getElementById("btnNonPensioner");
  const btnPensioner       = document.getElementById("btnPensioner");
  const sepPensionerNote   = document.getElementById("sepPensionerNote");

  if (btnNonPensioner)  btnNonPensioner.classList.toggle("active", !pensioner);
  if (btnPensioner)     btnPensioner.classList.toggle("active", pensioner);
  if (sepPensionerNote) sepPensionerNote.style.display = pensioner ? "block" : "none";

  updateDashboard();
}

function registerEvents() {
  document.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    if (target.classList.contains("toggle-btn")) {
      const fieldId = target.dataset.field;
      const type = target.dataset.type;
      if (fieldId && type) setType(fieldId, type);
    }
  });

  document.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    if (target.matches('input[type="number"], input[type="date"]')) {
      updateDashboard();
    }
  });
}

function initApp() {
  registerEvents();
  updateDashboard();
}

// ─── Client Profile Save / Load ──────────────────────────────────────────────

const PROFILE_VERSION = 1;

// All static input IDs to capture
const PROFILE_STATIC_FIELDS = [
  "clientName", "dob", "presentationDate", "consultantName",
  "employment", "bonus", "otherIncome",
  "cpfOABalance", "cpfSABalance", "cpfMABalance",
  "food", "apparel", "transport", "subscriptions", "bills",
  "allowances", "recreation", "holidays",
  "loanHouse", "loanCar", "loanEducation", "loanBusiness", "loanPersonal",
  "hospPremium", "wholePremium", "termPremium", "ciPremium",
  "liquidCashCurrent",
  "endowment", "endowmentCurrent",
  "investments", "investmentsCurrent",
  "fixedDeposits", "fixedDepositsCurrent",
  "stocks", "stocksCurrent",
  "otherSavings", "otherSavingsCurrent"
];

// Fields that have an associated textarea note in the same row/tr
const PROFILE_NOTE_FIELDS = [
  "employment", "bonus", "otherIncome",
  "food", "apparel", "transport", "subscriptions", "bills",
  "allowances", "recreation", "holidays",
  "loanHouse", "loanCar", "loanEducation", "loanBusiness", "loanPersonal",
  "hospPremium", "wholePremium", "termPremium", "ciPremium",
  "liquidCashCurrent", "endowment", "investments", "fixedDeposits", "stocks", "otherSavings"
];

function getRowNote(fieldId) {
  const input = document.getElementById(fieldId);
  if (!input) return "";
  const row = input.closest(".item-row") || input.closest("tr");
  return row?.querySelector("textarea.item-note")?.value || "";
}

function setRowNote(fieldId, note) {
  const input = document.getElementById(fieldId);
  if (!input) return;
  const row = input.closest(".item-row") || input.closest("tr");
  const ta = row?.querySelector("textarea.item-note");
  if (ta) ta.value = note;
}

function collectDynamicRowsForExport(containerId) {
  const rows = [];
  document.querySelectorAll(`#${containerId} .item-row`).forEach((row) => {
    const input = row.querySelector('input[type="number"]:not([data-current])');
    if (!input) return;
    rows.push({
      title:     row.querySelector(".item-label")?.textContent?.trim() || "",
      value:     input.value,
      frequency: inputTypes[input.id] || "Monthly",
      note:      row.querySelector("textarea")?.value || ""
    });
  });
  return rows;
}

function collectSavingsDynamicRowsForExport() {
  const rows = [];
  document.querySelectorAll("#additionalSavings .item-row").forEach((row) => {
    const contribInput = row.querySelector('input[type="number"]:not([data-current])');
    const currentInput = row.querySelector('input[data-current="true"]');
    if (!contribInput) return;
    rows.push({
      title:        row.querySelector(".item-label")?.textContent?.trim() || "",
      value:        contribInput.value,
      frequency:    inputTypes[contribInput.id] || "Monthly",
      note:         row.querySelector("textarea")?.value || "",
      currentValue: currentInput?.value || ""
    });
  });
  return rows;
}

function saveClientProfile() {
  const profile = {
    version:  PROFILE_VERSION,
    savedAt:  new Date().toISOString(),
    fields:   {},
    notes:    {},
    toggles:  { ...inputTypes },
    mode:     { isSelfEmployed, isPensioner },
    dynamic:  {
      additionalPersonal:  collectDynamicRowsForExport("additionalPersonal"),
      additionalLoans:     collectDynamicRowsForExport("additionalLoans"),
      additionalInsurance: collectDynamicRowsForExport("additionalInsurance"),
      additionalSavings:   collectSavingsDynamicRowsForExport()
    }
  };

  PROFILE_STATIC_FIELDS.forEach((id) => {
    const el = document.getElementById(id);
    if (el) profile.fields[id] = el.value;
  });

  PROFILE_NOTE_FIELDS.forEach((id) => {
    profile.notes[id] = getRowNote(id);
  });

  const name     = (document.getElementById("clientName")?.value?.trim() || "client")
                     .replace(/[^a-z0-9_\-]/gi, "_");
  const date     = new Date().toISOString().slice(0, 10);
  const filename = `${name}_cashflow_${date}.json`;

  const blob = new Blob([JSON.stringify(profile, null, 2)], { type: "application/json" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function loadClientProfile() {
  document.getElementById("profileFileInput")?.click();
}

function onProfileFileSelected(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const profile = JSON.parse(ev.target.result);
      applyClientProfile(profile);
    } catch {
      alert("Could not load profile — please select a valid .json profile file.");
    }
    event.target.value = ""; // reset so the same file can be re-selected
  };
  reader.readAsText(file);
}

function applyClientProfile(profile) {
  // 1. Restore static field values
  Object.entries(profile.fields || {}).forEach(([id, value]) => {
    const el = document.getElementById(id);
    if (el) el.value = value;
  });

  // 2. Restore textarea notes
  Object.entries(profile.notes || {}).forEach(([id, note]) => {
    setRowNote(id, note);
  });

  // 3. Restore Monthly/Annual toggle states + re-highlight buttons
  Object.assign(inputTypes, profile.toggles || {});
  Object.entries(inputTypes).forEach(([fieldId, type]) => {
    document.querySelectorAll(`.toggle-btn[data-field="${fieldId}"]`).forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.type === type);
    });
  });

  // 4. Restore employment / pensioner mode
  if (profile.mode) {
    setEmploymentMode(!!profile.mode.isSelfEmployed);
    setPensionerMode(!!profile.mode.isPensioner);
  }

  // 5. Rebuild dynamic rows
  const dynamic = profile.dynamic || {};
  const rowConfig = {
    additionalPersonal:  { prefix: "persExp", notes: "Details" },
    additionalLoans:     { prefix: "loan",    notes: "Loan details" },
    additionalInsurance: { prefix: "ins",     notes: "Plan details" }
  };

  Object.entries(rowConfig).forEach(([containerId, cfg]) => {
    const container = document.getElementById(containerId);
    if (container) container.innerHTML = "";
    (dynamic[containerId] || []).forEach((item) => {
      const row   = createDynamicRow({ idPrefix: cfg.prefix, title: item.title, notesPlaceholder: cfg.notes });
      const input = row.querySelector('input[type="number"]');
      if (input) {
        input.value = item.value;
        inputTypes[input.id] = item.frequency || "Monthly";
        row.querySelectorAll(".toggle-btn").forEach((btn) => {
          btn.classList.toggle("active", btn.dataset.type === (item.frequency || "Monthly"));
        });
      }
      const ta = row.querySelector("textarea");
      if (ta) ta.value = item.note || "";
      container?.appendChild(row);
    });
  });

  // Additional savings rows (have the extra Current Value column)
  const savingsContainer = document.getElementById("additionalSavings");
  if (savingsContainer) savingsContainer.innerHTML = "";
  (dynamic.additionalSavings || []).forEach((item) => {
    const row         = createSavingsDynamicRow(item.title);
    const contribInput = row.querySelector('input[type="number"]:not([data-current])');
    const currentInput = row.querySelector('input[data-current="true"]');
    if (contribInput) {
      contribInput.value = item.value;
      inputTypes[contribInput.id] = item.frequency || "Monthly";
      row.querySelectorAll(".toggle-btn").forEach((btn) => {
        btn.classList.toggle("active", btn.dataset.type === (item.frequency || "Monthly"));
      });
    }
    if (currentInput) currentInput.value = item.currentValue || "";
    const ta = row.querySelector("textarea");
    if (ta) ta.value = item.note || "";
    savingsContainer?.appendChild(row);
  });

  updateDashboard();
}

// ─────────────────────────────────────────────────────────────────────────────

window.setType = setType;
window.updateDashboard = updateDashboard;
window.addPersonalExpense = addPersonalExpense;
window.addLoan = addLoan;
window.addInsurance = addInsurance;
window.addSavings = addSavings;
window.exportToExcel = exportToExcel;
window.setEmploymentMode = setEmploymentMode;
window.setPensionerMode = setPensionerMode;
window.saveClientProfile = saveClientProfile;
window.loadClientProfile = loadClientProfile;
window.onProfileFileSelected = onProfileFileSelected;

document.addEventListener("DOMContentLoaded", initApp);
