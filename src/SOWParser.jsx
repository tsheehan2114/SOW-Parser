import { useState, useRef } from "react";
import * as XLSX from "xlsx";

const SKUS = [
  { name: "Strategic Market Research", code: "MKT-LCH-MSRCA", feeType: "One-Time", revRec: "Milestone", owner: "David C", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Unified Discovery", code: "MKT-LCH-UD", feeType: "One-Time", revRec: "Milestone", owner: "Melissa", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Go-To-Market Strategy & Positioning", code: "MKT-LCH-GTM", feeType: "One-Time", revRec: "Milestone", owner: "Melissa", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Creative Concepting", code: "MKT-LCH-CC", feeType: "One-Time", revRec: "Milestone", owner: "Melissa", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Website Strategy and Development", code: "MKT-LCH-WEBDEV", feeType: "One-Time", revRec: "Milestone", owner: "Melissa", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Recruitment Marketing Setup", code: "MKT-LCH-RMSET", feeType: "One-Time", revRec: "Milestone", owner: "Melissa", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Retention Marketing Setup", code: "MKT-ANN-RETEN-IMP", feeType: "One-Time", revRec: "Milestone", owner: "Melissa", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Photo & Video Shoot (One-Day)", code: "MKT-LCH-PHOTO", feeType: "One-Time", revRec: "Milestone", owner: "Melissa", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Marketing & Enrollment Analysis & Planning", code: "MKT-ANN-MAP", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Recruitment Marketing (Annual)", code: "MKT-ANN-RMANN", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Program Maturity Assets", code: "MKT-ANN-PMA", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Retention Marketing", code: "MKT-ANN-RETEN-ANN", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Public Relations", code: "MKT-ANN-PUB-REL", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "Social Media Engagement", code: "MKT-ANN-SOCIAL", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "Media Campaigns & Placement", code: "MKT-ANN-MEDIA", feeType: "Annual", revRec: "Usage - Spend", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Media Management Fee", code: "MKT-ANN-MEDIA%", feeType: "Annual", revRec: "Usage - Spend", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Marketing & Enrollment Travel", code: "MKT-LCH-TRVL", feeType: "One-Time", revRec: "Usage - Spend", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Transcript Services", code: "ENR-ANN-TRAN", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Enrollment Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Enrollment Advisors", code: "ENR-ANN-ADVISE", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Enrollment Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Program Design & Discovery", code: "LD-LCH-PRODES", feeType: "One-Time", revRec: "Milestone", owner: "Saskia", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Recurring" },
  { name: "Essentials Course Builds", code: "LD-LCH-ESS", feeType: "One-Time", revRec: "Milestone", owner: "Saskia", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "Course Maintenance", code: "LD-LCH-MAINT", feeType: "One-Time", revRec: "Milestone", owner: "Saskia", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "Orientation Course", code: "LD-LCH-ORIEN", feeType: "One-Time", revRec: "Milestone", owner: "Saskia", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "On-Site Video", code: "LD-LCH-VIDEO", feeType: "One-Time", revRec: "Milestone", owner: "Saskia", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "Media Credits", code: "LD-LCH-MECRED", feeType: "One-Time", revRec: "Milestone", owner: "Saskia", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Term Prep", code: "LD-ANN-TRMPRP", feeType: "One-Time", revRec: "Usage - Spend", owner: "Sushyla", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Faculty Support", code: "LD-ANN-FCLTSP", feeType: "Annual", revRec: "Straight Line", owner: "Saskia", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Course Refresh", code: "LD-LCH-EREF", feeType: "One-Time", revRec: "Milestone", owner: "Saskia", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "Noodle:Dialogue Faculty Training", code: "LD-LCH-TRAIN", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Learning Design Travel", code: "LD-LCH-TRVL", feeType: "One-Time", revRec: "Usage - Spend", owner: "", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Learning Design Consulting", code: "LD-CNSL-LD", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Learning", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "Student Support Coaching", code: "SSP-ANN-COACH", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Student Support & Retention", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Student Support Consulting", code: "SSP-CNSL-SSP", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Student Support & Retention", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Student Support Travel", code: "SSP-LCH-TRVL", feeType: "One-Time", revRec: "Usage - Spend", owner: "", primaryGroup: "Student Support & Retention", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Placement Services", code: "PLC-LCH-LVL", feeType: "One-Time", revRec: "Straight Line", owner: "", primaryGroup: "Placement Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Noodle Learning Platform Licensing Fee", code: "TECH-ANN-NLP-LIC", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Noodle Learning Platform Updated Tier Licensing Fee", code: "TECH-ANN-NLP-LIC-UPD", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Noodle Learning Platform Discovery", code: "TECH-LCH-NLP-DISC", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Noodle Learning Platform LMS Integration", code: "TECH-LCH-NLP-LMSINT", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Noodle Moodle LMS", code: "TECH-ANN-NLP-MOODLE", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Noodle Learning Platform CMS Integration", code: "TECH-LCH-NLP-CMSINT", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "NLP Merchant of Record", code: "TECH-ANN-NLP-MOR", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "N:Engage CMS", code: "TECH-ANN-ENG-CMS", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "N:Engage AI Standalone Implementation", code: "TECH-LCH-ENG-AI-IMP", feeType: "One-Time", revRec: "Milestone", owner: "David D", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "N:Engage AI Standalone Licensing Fee", code: "TECH-ANN-ENG-AI-LIC", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "GEMS One-Time Integration", code: "TECH-LCH-GEMS-INT", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "GEMS Per Student Fee", code: "TECH-ANN-GEMS-STU", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "N:Engage CRM", code: "TECH-ANN-ENG-CRM-CENT", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "N:Engage CRM Lifelong Learning Setup", code: "TECH-LCH-ENG-CRM-IMP", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Noodle:Manage (Institutional Analytics)", code: "TECH-ANN-MNG-IA", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Noodle:Dialogue Implementation", code: "TECH-LCH-DLG-IMP", feeType: "One-Time", revRec: "Milestone", owner: "David D", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Noodle:Dialogue Licensing Fee", code: "TECH-ANN-DLG-LIC", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Noodle:Companion", code: "TECH-ANN-CMPN", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Noodle:Manage", code: "TECH-ANN-MNG", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Tuition Calculator", code: "TECH-ANN-TUI-CALC-ANN", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Tuition Calculator Setup", code: "TECH-LCH-TUI-CALC-IMP", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Non-recurring" },
  { name: "Slate Consulting", code: "TECH-LCH-SLATE", feeType: "One-Time", revRec: "Straight Line", owner: "David D", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Tech Consulting", code: "TECH-CNSL-TECH", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "Tech Travel", code: "TECH-LCH-TRVL", feeType: "One-Time", revRec: "Usage - Spend", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Additional LLM Tokens", code: "TECH-LCH-TOKEN", feeType: "One-Time", revRec: "Straight Line", owner: "", primaryGroup: "Technology", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Marketing & Enrollment Consulting", code: "MKT-CNSL-M&E", feeType: "One-Time", revRec: "Milestone", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Non-recurring" },
  { name: "Performance Based Partnership", code: "MKT-ANN-PERF", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Marketing Services", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Program Fees", code: "PRT-ANN-PGRM-FEE", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Partner Success", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "State Authorizations", code: "STR-ANN-STA-AUTH", feeType: "Annual", revRec: "Straight Line", owner: "", primaryGroup: "Strategy", startLogic: "Billing Date", endLogic: "SOW End Date", frequency: "Recurring" },
  { name: "Partner Success Travel", code: "PRT-LCH-TRVL", feeType: "One-Time", revRec: "Usage - Spend", owner: "", primaryGroup: "Partner Success", startLogic: "Billing Date", endLogic: "Billing Date + 3 Months", frequency: "Recurring" },
];

const EXTRACT_PROMPT = `You are a financial data extraction assistant for a company that sells education services.

SKU Catalog (Name | Code | Fee Type | Rev Rec | Start Date Logic | End Date Logic):
${SKUS.map(s => `- ${s.name} | ${s.code} | ${s.feeType} | ${s.revRec} | Start: ${s.startLogic} | End: ${s.endLogic}`).join("\n")}

IMPORTANT DATE RULES:
- "Billing Date" = the first billing/invoice date for that line item as stated in the SOW.
- "Billing Date + 3 Months" = 3 calendar months after that line item's billing date.
- "SOW End Date" = the overall contract end date found in the SOW.
- Apply these rules to infer serviceStartDate and serviceEndDate for each line item if exact dates are not explicitly stated.

EXTRACTION RULES:
- Read through the ENTIRE document. Do NOT stop after the first billing table.
- Capture every single invoice entry across ALL years (Year 1, Year 2, Year 3, Year 4, etc.).
- Pay special attention to later years where only one or two services may remain active — these are easy to miss but must be included.
- If the same SKU appears across multiple years, consolidate into a SINGLE line item with ALL invoices as individual milestones.
- Each invoice = its own milestone with its own date and amount.
- totalServiceAmount = sum of ALL milestones across all years.
- Only include SKUs explicitly mentioned in the SOW. Do NOT infer or hallucinate line items.
- ALWAYS use the Rev Rec method from the SKU Catalog above. NEVER override it based on billing structure or number of payments.

Return ONLY valid JSON, no markdown:
{
  "school": "string",
  "program": "string",
  "primaryGroup": "string",
  "serviceStartDate": "string",
  "serviceEndDate": "string",
  "lineItems": [
    {
      "skuName": "exact SKU name from catalog",
      "skuCode": "matching SKU code",
      "billingFrequency": "One-Time or Annual",
      "revRecMethod": "Straight Line, Milestone, Usage - Spend, etc.",
      "totalServiceAmount": number,
      "serviceStartDate": "string",
      "serviceEndDate": "string",
      "milestones": [
        {
          "milestoneName": "e.g. Invoice #1 2026",
          "billDate": "e.g. March 1, 2026",
          "billAmount": number
        }
      ]
    }
  ]
}`;

const VERIFY_PROMPT = `You are a financial data verification assistant. You have been given:
1. The original SOW PDF
2. An initial extraction of line items and milestones

Your job is to verify the extraction is complete and accurate by:
- Re-reading every billing table in the SOW across ALL years
- Checking that every service line item mentioned in the SOW is present in the extraction
- Checking that every single invoice across all years is captured as a milestone
- Checking that no invoices have been missed, especially in later years (2027, 2028, 2029, etc.) where fewer services appear
- Correcting any missing line items or milestones
- Ensuring milestone billAmounts are numbers, not strings
- ALWAYS use the Rev Rec method from the SKU Catalog. NEVER override it based on billing structure or number of payments.

Return the corrected complete JSON in exactly the same format. Return ONLY valid JSON, no markdown.`;

const MAIN_COLS = ["School","Primary Group","SKU Name","SKU Code","Frequency","Billing Frequency","Rev Rec Method","Total Service Amount","Service Period Start Date","Service Period End Date","Owner","Total Billing Periods"];
const MILESTONE_COLS = ["School","SKU Name","SKU Code","Owner","Billing Period #","Billing Period Name","Bill Date","Bill Amount","Service Period Start","Service Period End"];
const AUDIT_COLS = ["School","SKU Name","SKU Code","Field","AI Extracted Value","Manually Edited Value","Edited At"];

export default function SOWParser() {
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState("idle");
  const [statusMsg, setStatusMsg] = useState("");
  const [meta, setMeta] = useState({ school: "", program: "", primaryGroup: "", serviceStartDate: "", serviceEndDate: "" });
  const [rows, setRows] = useState([]);
  const [error, setError] = useState("");
  const [exported, setExported] = useState(false);
  const [activeTab, setActiveTab] = useState("main");
  const [auditLog, setAuditLog] = useState([]);
  const [integrityWarnings, setIntegrityWarnings] = useState([]);
  const [skuWarnings, setSkuWarnings] = useState([]);
  const inputRef = useRef();
  const aiOriginals = useRef({});

  function levenshtein(a, b) {
    const m = a.length, n = b.length;
    const dp = Array.from({ length: m + 1 }, (_, i) => Array.from({ length: n + 1 }, (_, j) => i === 0 ? j : j === 0 ? i : 0));
    for (let i = 1; i <= m; i++)
      for (let j = 1; j <= n; j++)
        dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1] : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    return dp[m][n];
  }

  function fuzzyFindSku(skuCode, skuName = "") {
    const exact = SKUS.find(s => s.code === skuCode);
    if (exact) return { sku: exact, matchType: "exact" };
    const caseInsensitive = SKUS.find(s => s.code.toLowerCase() === skuCode.toLowerCase());
    if (caseInsensitive) return { sku: caseInsensitive, matchType: "case-insensitive" };
    const candidates = SKUS.map(s => {
      const codeDist = levenshtein(skuCode.toUpperCase(), s.code.toUpperCase());
      const nameDist = skuName ? levenshtein(skuName.toLowerCase(), s.name.toLowerCase()) : 999;
      return { sku: s, codeDist, nameDist };
    });
    candidates.sort((a, b) => a.codeDist - b.codeDist || a.nameDist - b.nameDist);
    const best = candidates[0];
    const threshold = Math.max(3, Math.ceil(skuCode.length * 0.35));
    if (best.codeDist <= threshold) return { sku: best.sku, matchType: "fuzzy", distance: best.codeDist };
    return { sku: null, matchType: "unmatched" };
  }

  function parseBillDate(str) {
    if (!str) return null;
    const d = new Date(str);
    if (!isNaN(d)) return d;
    const m = str.match(/^(\w+)\s+(\d+),?\s+(\d{4})$/);
    if (m) return new Date(`${m[1]} ${m[2]}, ${m[3]}`);
    return null;
  }

  function dayBefore(date) {
    const d = new Date(date);
    d.setDate(d.getDate() - 1);
    return d;
  }

  function formatDate(date) {
    if (!date || isNaN(date)) return "";
    return date.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
  }

  function computeServicePeriod(milestones, index, sowEndDate, skuCode) {
    const current = parseBillDate(milestones[index]?.billDate);
    if (!current) return { start: "", end: "" };
    const start = formatDate(current);

    // For non-last billing periods: end = day before next invoice
    const next = milestones[index + 1];
    if (next) {
      const nextDate = parseBillDate(next.billDate);
      if (nextDate) return { start, end: formatDate(dayBefore(nextDate)) };
    }

    // For the last billing period: use SKU endLogic from catalog
    const { sku } = fuzzyFindSku(skuCode || "");
    const endLogic = sku?.endLogic || "SOW End Date";

    if (endLogic === "Billing Date + 3 Months") {
      const threeMonths = new Date(current);
      threeMonths.setMonth(threeMonths.getMonth() + 3);
      threeMonths.setDate(threeMonths.getDate() - 1); // end of day before 3-month mark
      return { start, end: formatDate(threeMonths) };
    }

    // Default: SOW End Date
    if (sowEndDate) {
      const sowEnd = parseBillDate(sowEndDate);
      if (sowEnd && !isNaN(sowEnd)) return { start, end: formatDate(sowEnd) };
    }
    return { start, end: "" };
  }

  const toNumber = (val) => {
    if (val === null || val === undefined || val === "") return 0;
    if (typeof val === "number") return val;
    const n = parseFloat(String(val).replace(/[$,]/g, ""));
    return isNaN(n) ? 0 : n;
  };

  const runIntegrityCheck = (rowsToCheck) => {
    const warnings = [];
    const TOLERANCE = 0.02;
    rowsToCheck.forEach((r, i) => {
      const msTotal = (r.milestones || []).reduce((sum, m) => sum + toNumber(m.billAmount), 0);
      const declared = toNumber(r.totalServiceAmount);
      if (declared > 0 && msTotal > 0) {
        const diff = Math.abs(msTotal - declared);
        if (diff / declared > TOLERANCE) {
          warnings.push({ rowIndex: i, skuName: r.skuName, skuCode: r.skuCode, declared, milestonesTotal: msTotal, diff });
        }
      }
      if ((r.milestones || []).length > 0 && msTotal === 0) {
        warnings.push({ rowIndex: i, skuName: r.skuName, skuCode: r.skuCode, declared, milestonesTotal: 0, diff: declared, zeroSum: true });
      }
    });
    setIntegrityWarnings(warnings);
  };

  const runSkuCheck = (rowsToCheck) => {
    const warnings = [];
    rowsToCheck.forEach((r, i) => {
      if (!r.skuCode) return;
      const { matchType, sku, distance } = fuzzyFindSku(r.skuCode, r.skuName);
      if (matchType === "unmatched") {
        warnings.push({ rowIndex: i, skuCode: r.skuCode, skuName: r.skuName, type: "unmatched", suggestion: null });
      } else if (matchType === "fuzzy" || matchType === "case-insensitive") {
        warnings.push({ rowIndex: i, skuCode: r.skuCode, skuName: r.skuName, type: matchType, suggestion: sku.code, distance });
      }
    });
    setSkuWarnings(warnings);
  };

  const getOwner = (skuCode) => { const { sku } = fuzzyFindSku(skuCode); return sku?.owner || ""; };
  const getPrimaryGroup = (skuCode) => { const { sku } = fuzzyFindSku(skuCode); return sku?.primaryGroup || ""; };
  const getFrequency = (skuCode) => { const { sku } = fuzzyFindSku(skuCode); return sku?.frequency || ""; };

  const handleFile = e => {
    const f = e.target.files[0];
    if (f && f.type === "application/pdf") { setFile(f); setStatus("idle"); setExported(false); setAuditLog([]); setIntegrityWarnings([]); setSkuWarnings([]); }
    else alert("Please upload a PDF file.");
  };

  const toBase64 = f => new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = rej;
    r.readAsDataURL(f);
  });

  const parseJSON = (raw) => {
    const cleaned = raw.replace(/^```json\s*/i, "").replace(/```\s*$/i, "").trim();
    try { return JSON.parse(cleaned); } catch (_) {}
    if (!cleaned.endsWith("}")) {
      const lastGood = cleaned.lastIndexOf('},');
      const repaired = lastGood > 0 ? cleaned.slice(0, lastGood + 1) + ']}' : cleaned;
      const parsed = JSON.parse(repaired);
      if (!parsed.lineItems || !Array.isArray(parsed.lineItems)) throw new Error("Repaired JSON missing lineItems array.");
      return parsed;
    }
    throw new Error("Could not parse JSON response.");
  };

  const callAPI = async (b64, systemPrompt, userMessage) => {
    const resp = await fetch("/api/claude", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 16000,
        system: systemPrompt,
        messages: [{ role: "user", content: [
          { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } },
          { type: "text", text: userMessage }
        ]}]
      })
    });
    const json = await resp.json();
    if (json.error) throw new Error(`API error: ${json.error.type} — ${json.error.message}`);
    const raw = json.content?.find(b => b.type === "text")?.text || "";
    if (!raw) throw new Error("The API returned an empty response.");
    return raw;
  };

  const parse = async () => {
    if (!file) return;
    setStatus("loading"); setError(""); setAuditLog([]); setIntegrityWarnings([]); setSkuWarnings([]);
    aiOriginals.current = {};

    const MB = file.size / (1024 * 1024);
    if (MB > 32) {
      setError(`This PDF is ${MB.toFixed(1)}MB — too large to process.\n• Adobe Acrobat: File → Save As → Reduced Size PDF\n• Preview (Mac): File → Export as PDF → Quartz Filter → Reduce File Size\n• Online: ilovepdf.com or smallpdf.com`);
      setStatus("error"); return;
    }
    if (MB > 10) setStatusMsg(`Note: PDF is ${MB.toFixed(1)}MB. If extraction fails, try compressing it first.`);

    try {
      const b64 = await toBase64(file);
      setStatusMsg("Step 1 of 2: Extracting line items and milestones…");
      const raw1 = await callAPI(b64, EXTRACT_PROMPT, "Extract all financial line items and milestones from this SOW across all years.");
      let parsed1;
      try { parsed1 = parseJSON(raw1); }
      catch (e) { throw new Error(`Extraction parse error: ${e.message}. Preview: "${raw1.slice(0, 200)}"`); }

      setStatusMsg("Step 2 of 2: Verifying completeness and accuracy…");
      const raw2 = await callAPI(b64, VERIFY_PROMPT, `Here is the initial extraction. Please verify it is complete and correct against the SOW:\n\n${JSON.stringify(parsed1, null, 2)}`);
      let parsed2;
      try { parsed2 = parseJSON(raw2); }
      catch (e) { console.warn("Verification pass failed, using initial extraction:", e.message); parsed2 = parsed1; }

      setMeta({ school: parsed2.school || "", program: parsed2.program || "", primaryGroup: parsed2.primaryGroup || "", serviceStartDate: parsed2.serviceStartDate || "", serviceEndDate: parsed2.serviceEndDate || "" });

      const processedItems = (parsed2.lineItems || []).map((r, i) => {
        const { sku, matchType } = fuzzyFindSku(r.skuCode, r.skuName);
        const correctedCode = (matchType === "fuzzy" || matchType === "case-insensitive") && sku ? sku.code : r.skuCode;

        // Deduplicate milestones — the two-pass system can produce identical
        // billDate + billAmount pairs. Keep only the first occurrence of each.
        const seen = new Set();
        const dedupedMilestones = (r.milestones || []).filter(m => {
          const key = String(m.billDate).trim() + "|" + String(m.billAmount).trim();
          if (seen.has(key)) return false;
          seen.add(key);
          return true;
        });

        return { ...r, _id: i, skuCode: correctedCode, milestones: dedupedMilestones };
      });

      processedItems.forEach(r => {
        aiOriginals.current[r._id] = {
          skuName: r.skuName, skuCode: r.skuCode, billingFrequency: r.billingFrequency,
          revRecMethod: r.revRecMethod, serviceStartDate: r.serviceStartDate, serviceEndDate: r.serviceEndDate,
          milestones: JSON.parse(JSON.stringify(r.milestones)),
        };
      });

      setRows(processedItems);
      runIntegrityCheck(processedItems);
      runSkuCheck(processedItems);
      setStatus("review");
      setStatusMsg("");
    } catch (e) {
      setError(e.message);
      setStatus("error");
      setStatusMsg("");
    }
  };

  const logEdit = (rowIndex, field, oldVal, newVal, context = {}) => {
    if (String(oldVal) === String(newVal)) return;
    const row = rows[rowIndex];
    setAuditLog(log => [...log, {
      school: meta.school, skuName: context.skuName || row?.skuName || "",
      skuCode: context.skuCode || row?.skuCode || "", field,
      aiValue: String(oldVal ?? ""), editedValue: String(newVal ?? ""),
      editedAt: new Date().toISOString(),
    }]);
  };

  const updateRow = (i, key, val) => {
    const original = aiOriginals.current[rows[i]?._id]?.[key];
    logEdit(i, key, original ?? rows[i]?.[key], val);
    setRows(rs => {
      const updated = rs.map((r, idx) => idx === i ? { ...r, [key]: val } : r);
      if (key === "totalServiceAmount") runIntegrityCheck(updated);
      if (key === "skuCode") runSkuCheck(updated);
      return updated;
    });
  };

  const updateMilestone = (ri, mi, key, val) => {
    const origVal = aiOriginals.current[rows[ri]?._id]?.milestones?.[mi]?.[key];
    const row = rows[ri];
    logEdit(ri, `milestone[${mi}].${key}`, origVal, val, { skuName: row?.skuName, skuCode: row?.skuCode });
    setRows(rs => {
      const updated = rs.map((r, idx) => idx === ri ? { ...r, milestones: r.milestones.map((m, mx) => mx === mi ? { ...m, [key]: val } : m) } : r);
      if (key === "billAmount") runIntegrityCheck(updated);
      return updated;
    });
  };

  const addMilestone = ri => setRows(rs => { const u = rs.map((r, idx) => idx === ri ? { ...r, milestones: [...r.milestones, { milestoneName: "", billDate: "", billAmount: "" }] } : r); runIntegrityCheck(u); return u; });
  const removeMilestone = (ri, mi) => setRows(rs => { const u = rs.map((r, idx) => idx === ri ? { ...r, milestones: r.milestones.filter((_, mx) => mx !== mi) } : r); runIntegrityCheck(u); return u; });
  const addRow = () => setRows(rs => [...rs, { _id: Date.now(), skuName: "", skuCode: "", billingFrequency: "", revRecMethod: "", totalServiceAmount: "", serviceStartDate: meta.serviceStartDate, serviceEndDate: meta.serviceEndDate, milestones: [] }]);
  const removeRow = i => setRows(rs => { const u = rs.filter((_, idx) => idx !== i); runIntegrityCheck(u); runSkuCheck(u); return u; });
  const applySkuSuggestion = (rowIndex, suggestedCode) => updateRow(rowIndex, "skuCode", suggestedCode);

  const exportXLSX = () => {
    const wb = XLSX.utils.book_new();
    const mainRows = rows.map(r => {
      const msTotal = (r.milestones || []).reduce((sum, m) => sum + toNumber(m.billAmount), 0);
      return { "School": meta.school, "Primary Group": getPrimaryGroup(r.skuCode), "SKU Name": r.skuName, "SKU Code": r.skuCode, "Frequency": getFrequency(r.skuCode), "Billing Frequency": r.billingFrequency, "Rev Rec Method": r.revRecMethod, "Total Service Amount": msTotal, "Service Period Start Date": r.serviceStartDate, "Service Period End Date": r.serviceEndDate, "Owner": getOwner(r.skuCode), "Total Billing Periods": (r.milestones || []).length };
    });
    const ws1 = XLSX.utils.json_to_sheet(mainRows, { header: MAIN_COLS });
    ws1["!cols"] = MAIN_COLS.map(c => ({ wch: Math.max(c.length + 2, 14) }));
    XLSX.utils.book_append_sheet(wb, ws1, "SOW Extract");

    const msRows = [];
    rows.forEach(r => {
      (r.milestones || []).forEach((m, mi) => {
        const sp = computeServicePeriod(r.milestones, mi, meta.serviceEndDate, r.skuCode);
        msRows.push({ "School": meta.school, "SKU Name": r.skuName, "SKU Code": r.skuCode, "Owner": getOwner(r.skuCode), "Billing Period #": mi + 1, "Billing Period Name": m.milestoneName, "Bill Date": m.billDate, "Bill Amount": toNumber(m.billAmount), "Service Period Start": sp.start, "Service Period End": sp.end });
      });
    });
    const ws2 = XLSX.utils.json_to_sheet(msRows.length ? msRows : [{}], { header: MILESTONE_COLS });
    ws2["!cols"] = MILESTONE_COLS.map(c => ({ wch: Math.max(c.length + 2, 16) }));
    XLSX.utils.book_append_sheet(wb, ws2, "Billing Periods");

    const ws3 = XLSX.utils.json_to_sheet(SKUS.map(s => ({ "SKU Name": s.name, "SKU Code": s.code, "Fee Type": s.feeType, "Rev Rec Method": s.revRec, "Owner": s.owner, "Start Date Logic": s.startLogic, "End Date Logic": s.endLogic })));
    ws3["!cols"] = [{ wch: 48 }, { wch: 24 }, { wch: 14 }, { wch: 18 }, { wch: 12 }, { wch: 22 }, { wch: 28 }];
    XLSX.utils.book_append_sheet(wb, ws3, "SKU Reference");

    const auditRows = auditLog.map(a => ({ "School": a.school, "SKU Name": a.skuName, "SKU Code": a.skuCode, "Field": a.field, "AI Extracted Value": a.aiValue, "Manually Edited Value": a.editedValue, "Edited At": a.editedAt }));
    const ws4 = XLSX.utils.json_to_sheet(auditRows.length ? auditRows : [{ "Field": "No manual edits made" }], { header: AUDIT_COLS });
    ws4["!cols"] = AUDIT_COLS.map(c => ({ wch: Math.max(c.length + 2, 20) }));
    XLSX.utils.book_append_sheet(wb, ws4, "Audit Trail");

    XLSX.writeFile(wb, `${file?.name?.replace(/\.pdf$/i, "").trim() || "export"}.xlsx`);
    setExported(true);
  };

  const reset = () => { setFile(null); setStatus("idle"); setRows([]); setExported(false); setError(""); setStatusMsg(""); setAuditLog([]); setIntegrityWarnings([]); setSkuWarnings([]); aiOriginals.current = {}; inputRef.current.value = ""; };

  const cell = (val, onChange) => (
    <input value={val || ""} onChange={e => onChange(e.target.value)} style={{ width: "100%", border: "none", outline: "none", fontSize: 12, fontFamily: "inherit", background: "transparent" }} />
  );

  const tabBtn = (id, label, badgeCount = 0, badgeColor = "#ef4444") => (
    <button onClick={() => setActiveTab(id)} style={{ padding: "7px 16px", fontSize: 13, fontWeight: activeTab === id ? 700 : 400, background: "none", border: "none", borderBottom: activeTab === id ? "2px solid #111" : "2px solid transparent", cursor: "pointer", color: activeTab === id ? "#111" : "#6b7280", display: "flex", alignItems: "center", gap: 6 }}>
      {label}
      {badgeCount > 0 && <span style={{ background: badgeColor, color: "#fff", borderRadius: 10, padding: "1px 7px", fontSize: 10, fontWeight: 700 }}>{badgeCount}</span>}
    </button>
  );

  const totalWarnings = integrityWarnings.length + skuWarnings.length;

  return (
    <div style={{ fontFamily: "Inter, sans-serif", maxWidth: 1100, margin: "0 auto", padding: "28px 20px", color: "#111" }}>
      <div style={{ marginBottom: 20 }}>
        <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0 }}>SOW Parser → NetSuite</h1>
        <p style={{ color: "#555", marginTop: 5, fontSize: 13 }}>Upload a SOW PDF to extract service line items, milestones, and SKU mappings.</p>
      </div>

      <div onClick={() => inputRef.current.click()} style={{ border: "2px dashed #d1d5db", borderRadius: 10, padding: "20px", textAlign: "center", cursor: "pointer", background: file ? "#f0fdf4" : "#fafafa", borderColor: file ? "#22c55e" : "#d1d5db", marginBottom: 16 }}>
        <input ref={inputRef} type="file" accept="application/pdf" onChange={handleFile} style={{ display: "none" }} />
        <div style={{ fontSize: 22, marginBottom: 4 }}>{file ? "✅" : "📄"}</div>
        <div style={{ fontSize: 13, fontWeight: 600, color: file ? "#16a34a" : "#374151" }}>
          {file ? `${file.name} (${(file.size / 1024 / 1024).toFixed(1)}MB)` : "Click to upload a PDF SOW"}
        </div>
        {!file && <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 4, lineHeight: 1.5 }}>PDF files only · Best results under 10MB · Max 32MB<br /><span style={{ color: "#6b7280" }}>Large file? Compress at ilovepdf.com or export as "Reduced Size PDF" from Acrobat</span></div>}
        {file && file.size > 10 * 1024 * 1024 && <div style={{ fontSize: 11, color: "#d97706", marginTop: 4 }}>⚠ Large file — extraction may be slower. Compress if it fails.</div>}
      </div>

      <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
        <button onClick={parse} disabled={!file || status === "loading"} style={{ background: !file || status === "loading" ? "#e5e7eb" : "#111", color: !file || status === "loading" ? "#9ca3af" : "#fff", border: "none", borderRadius: 7, padding: "9px 20px", fontSize: 13, fontWeight: 600, cursor: !file || status === "loading" ? "not-allowed" : "pointer" }}>
          {status === "loading" ? "Extracting…" : "Extract Line Items"}
        </button>
        {status !== "idle" && <button onClick={reset} style={{ background: "none", border: "1px solid #d1d5db", borderRadius: 7, padding: "9px 16px", fontSize: 13, cursor: "pointer", color: "#555" }}>Reset</button>}
      </div>

      {status === "loading" && (
        <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 10, padding: 20, textAlign: "center", color: "#555", fontSize: 13 }}>
          <div style={{ marginBottom: 8 }}>🔍 {statusMsg || "Processing…"}</div>
          <div style={{ fontSize: 11, color: "#9ca3af" }}>This may take 20–30 seconds with the verification pass.</div>
        </div>
      )}
      {status === "error" && <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 10, padding: 14, color: "#dc2626", fontSize: 13, whiteSpace: "pre-line" }}>{error}</div>}

      {status === "review" && (
        <div>
          {totalWarnings > 0 && (
            <div style={{ background: "#fffbeb", border: "1px solid #fcd34d", borderRadius: 10, padding: 14, marginBottom: 16 }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "#92400e", marginBottom: 8 }}>⚠ {totalWarnings} warning{totalWarnings !== 1 ? "s" : ""} found — review before exporting</div>
              {integrityWarnings.map((w, i) => (
                <div key={`int-${i}`} style={{ fontSize: 12, color: "#78350f", marginBottom: 4, paddingLeft: 8, borderLeft: "3px solid #f59e0b" }}>
                  {w.zeroSum ? <>💰 <strong>{w.skuName}</strong> — billing periods sum to $0</> : <>💰 <strong>{w.skuName}</strong> — declared total <strong>${w.declared.toLocaleString()}</strong> vs billing periods sum <strong>${w.milestonesTotal.toLocaleString()}</strong> (Δ ${w.diff.toLocaleString()})</>}
                </div>
              ))}
              {skuWarnings.map((w, i) => (
                <div key={`sku-${i}`} style={{ fontSize: 12, color: "#78350f", marginBottom: 4, paddingLeft: 8, borderLeft: "3px solid #f59e0b", display: "flex", alignItems: "center", gap: 8 }}>
                  {w.type === "unmatched"
                    ? <>🔍 <strong>{w.skuName}</strong> — code <code style={{ background: "#fef3c7", padding: "1px 4px", borderRadius: 3 }}>{w.skuCode}</code> not in SKU catalog</>
                    : <>🔍 <strong>{w.skuName}</strong> — <code style={{ background: "#fef3c7", padding: "1px 4px", borderRadius: 3 }}>{w.skuCode}</code> → <code style={{ background: "#fef3c7", padding: "1px 4px", borderRadius: 3 }}>{w.suggestion}</code>
                        <button onClick={() => applySkuSuggestion(w.rowIndex, w.suggestion)} style={{ fontSize: 10, background: "#f59e0b", color: "#fff", border: "none", borderRadius: 4, padding: "2px 8px", cursor: "pointer" }}>Apply fix</button>
                      </>
                  }
                </div>
              ))}
            </div>
          )}

          <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 10, padding: 14, marginBottom: 16, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            {[["School","school"],["Primary Group","primaryGroup"]].map(([label, key]) => (
              <div key={key}>
                <div style={{ fontSize: 10, fontWeight: 700, color: "#6b7280", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 3 }}>{label}</div>
                <input value={meta[key]} onChange={e => setMeta(m => ({ ...m, [key]: e.target.value }))} placeholder={`Enter ${label}`} style={{ width: "100%", border: "1px solid #e5e7eb", borderRadius: 6, padding: "6px 10px", fontSize: 13, fontFamily: "inherit", boxSizing: "border-box" }} />
              </div>
            ))}
          </div>

          <div style={{ borderBottom: "1px solid #e5e7eb", marginBottom: 16, display: "flex", gap: 4 }}>
            {tabBtn("main", `Line Items (${rows.length})`, integrityWarnings.length + skuWarnings.length)}
            {tabBtn("milestones", `Billing Periods (${rows.reduce((a, r) => a + (r.milestones?.length || 0), 0)})`)}
            {tabBtn("audit", `Audit Trail (${auditLog.length})`, 0, "#6b7280")}
          </div>

          {activeTab === "main" && (
            <div style={{ overflowX: "auto", marginBottom: 14 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#111", color: "#fff" }}>
                    {["SKU Name","SKU Code","Billing Freq","Rev Rec","Total Amount","Start Date","End Date","Owner","Billing Periods",""].map((h, i) => (
                      <th key={i} style={{ padding: "8px 10px", textAlign: "left", fontWeight: 600, fontSize: 11, whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, i) => {
                    const msTotal = (r.milestones || []).reduce((sum, m) => sum + toNumber(m.billAmount), 0);
                    const hasIntegrityIssue = integrityWarnings.some(w => w.rowIndex === i);
                    const hasSkuIssue = skuWarnings.some(w => w.rowIndex === i);
                    return (
                      <tr key={r._id} style={{ background: hasIntegrityIssue || hasSkuIssue ? "#fffbeb" : i % 2 === 0 ? "#fff" : "#f9fafb", borderBottom: "1px solid #f0f0f0", borderLeft: hasIntegrityIssue || hasSkuIssue ? "3px solid #f59e0b" : "3px solid transparent" }}>
                        {[["skuName",160],["skuCode",110],["billingFrequency",90],["revRecMethod",100]].map(([key, w]) => (
                          <td key={key} style={{ padding: "6px 10px", minWidth: w, borderRight: "1px solid #f0f0f0" }}>{cell(r[key], v => updateRow(i, key, v))}</td>
                        ))}
                        <td style={{ padding: "6px 10px", minWidth: 90, borderRight: "1px solid #f0f0f0", fontSize: 12, color: hasIntegrityIssue ? "#b45309" : "inherit", fontWeight: hasIntegrityIssue ? 600 : 400 }}>
                          ${msTotal.toLocaleString()}{hasIntegrityIssue && <span title="Sum mismatch"> ⚠</span>}
                        </td>
                        {[["serviceStartDate",100],["serviceEndDate",100]].map(([key, w]) => (
                          <td key={key} style={{ padding: "6px 10px", minWidth: w, borderRight: "1px solid #f0f0f0" }}>{cell(r[key], v => updateRow(i, key, v))}</td>
                        ))}
                        <td style={{ padding: "6px 10px", borderRight: "1px solid #f0f0f0", fontSize: 11, color: "#6b7280", whiteSpace: "nowrap" }}>{getOwner(r.skuCode) || "—"}</td>
                        <td style={{ padding: "6px 10px", borderRight: "1px solid #f0f0f0" }}>
                          <span style={{ background: (r.milestones?.length || 0) > 3 ? "#fef3c7" : "#f0fdf4", color: (r.milestones?.length || 0) > 3 ? "#92400e" : "#166534", borderRadius: 5, padding: "2px 8px", fontSize: 11, fontWeight: 600 }}>{r.milestones?.length || 0}</span>
                        </td>
                        <td style={{ padding: "6px 8px" }}><button onClick={() => removeRow(i)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ef4444", fontSize: 14, padding: 0 }}>✕</button></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              <button onClick={addRow} style={{ background: "none", border: "1px dashed #d1d5db", borderRadius: 7, padding: "6px 14px", fontSize: 12, cursor: "pointer", color: "#555", marginTop: 10 }}>+ Add Row</button>
            </div>
          )}

          {activeTab === "milestones" && (
            <div>
              {rows.map((r, ri) => {
                const msTotal = (r.milestones || []).reduce((sum, m) => sum + toNumber(m.billAmount), 0);
                const hasIssue = integrityWarnings.some(w => w.rowIndex === ri);
                return (
                  <div key={r._id} style={{ marginBottom: 20, border: `1px solid ${hasIssue ? "#fcd34d" : "#e5e7eb"}`, borderRadius: 10, overflow: "hidden" }}>
                    <div style={{ background: hasIssue ? "#fffbeb" : "#f8fafc", padding: "10px 14px", borderBottom: "1px solid #e5e7eb", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div>
                        <span style={{ fontWeight: 700, fontSize: 13 }}>{r.skuName || "Unnamed SKU"}</span>
                        <span style={{ fontSize: 11, color: "#6b7280", marginLeft: 8 }}>{r.skuCode}</span>
                        {getOwner(r.skuCode) && <span style={{ fontSize: 11, color: "#6b7280", marginLeft: 8 }}>· {getOwner(r.skuCode)}</span>}
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        {hasIssue && <span style={{ fontSize: 11, color: "#b45309" }}>⚠ Sum mismatch</span>}
                        <span style={{ fontSize: 12, color: hasIssue ? "#b45309" : "#6b7280", fontWeight: hasIssue ? 600 : 400 }}>${msTotal.toLocaleString()}</span>
                      </div>
                    </div>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr style={{ background: "#374151", color: "#fff" }}>
                          {["#","Billing Period Name","Bill Date","Bill Amount","Service Period Start","Service Period End",""].map((h, i) => (
                            <th key={i} style={{ padding: "7px 12px", textAlign: "left", fontWeight: 600, fontSize: 11 }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {(r.milestones || []).map((m, mi) => {
                          const sp = computeServicePeriod(r.milestones, mi, meta.serviceEndDate, r.skuCode);
                          return (
                            <tr key={mi} style={{ background: mi % 2 === 0 ? "#fff" : "#f9fafb", borderBottom: "1px solid #f0f0f0" }}>
                              <td style={{ padding: "6px 12px", color: "#9ca3af", width: 30 }}>{mi + 1}</td>
                              <td style={{ padding: "6px 12px", minWidth: 180 }}>{cell(m.milestoneName, v => updateMilestone(ri, mi, "milestoneName", v))}</td>
                              <td style={{ padding: "6px 12px", minWidth: 130 }}>{cell(m.billDate, v => updateMilestone(ri, mi, "billDate", v))}</td>
                              <td style={{ padding: "6px 12px", minWidth: 100 }}>{cell(m.billAmount, v => updateMilestone(ri, mi, "billAmount", v))}</td>
                              <td style={{ padding: "6px 12px", minWidth: 140, color: "#6b7280", fontSize: 11 }}>{sp.start}</td>
                              <td style={{ padding: "6px 12px", minWidth: 140, color: "#6b7280", fontSize: 11 }}>{sp.end || <span style={{ color: "#d1d5db" }}>—</span>}</td>
                              <td style={{ padding: "6px 8px" }}><button onClick={() => removeMilestone(ri, mi)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ef4444", fontSize: 13, padding: 0 }}>✕</button></td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                    <div style={{ padding: "8px 12px", borderTop: "1px solid #f0f0f0" }}>
                      <button onClick={() => addMilestone(ri)} style={{ background: "none", border: "1px dashed #d1d5db", borderRadius: 6, padding: "4px 12px", fontSize: 11, cursor: "pointer", color: "#555" }}>+ Add Billing Period</button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {activeTab === "audit" && (
            <div>
              {auditLog.length === 0 ? (
                <div style={{ textAlign: "center", padding: "32px 0", color: "#9ca3af", fontSize: 13 }}>No manual edits made yet.</div>
              ) : (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: "#111", color: "#fff" }}>
                        {["SKU Name","SKU Code","Field","AI Extracted","Manually Edited","Edited At"].map((h, i) => (
                          <th key={i} style={{ padding: "8px 10px", textAlign: "left", fontWeight: 600, fontSize: 11, whiteSpace: "nowrap" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {auditLog.map((a, i) => (
                        <tr key={i} style={{ background: i % 2 === 0 ? "#fff" : "#f9fafb", borderBottom: "1px solid #f0f0f0" }}>
                          <td style={{ padding: "6px 10px", borderRight: "1px solid #f0f0f0" }}>{a.skuName}</td>
                          <td style={{ padding: "6px 10px", borderRight: "1px solid #f0f0f0", color: "#6b7280" }}>{a.skuCode}</td>
                          <td style={{ padding: "6px 10px", borderRight: "1px solid #f0f0f0" }}><code style={{ background: "#f3f4f6", padding: "1px 5px", borderRadius: 3, fontSize: 11 }}>{a.field}</code></td>
                          <td style={{ padding: "6px 10px", borderRight: "1px solid #f0f0f0", color: "#dc2626", textDecoration: "line-through", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.aiValue || "—"}</td>
                          <td style={{ padding: "6px 10px", borderRight: "1px solid #f0f0f0", color: "#16a34a", fontWeight: 500, maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.editedValue}</td>
                          <td style={{ padding: "6px 10px", color: "#9ca3af", fontSize: 11, whiteSpace: "nowrap" }}>{new Date(a.editedAt).toLocaleString()}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          <div style={{ display: "flex", gap: 12, alignItems: "center", marginTop: 16 }}>
            <button onClick={exportXLSX} style={{ background: "#1d6f42", color: "#fff", border: "none", borderRadius: 7, padding: "10px 22px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>⬇ Export as Excel (.xlsx)</button>
            {exported && <span style={{ fontSize: 13, color: "#16a34a", fontWeight: 500 }}>✓ File downloaded</span>}
            {totalWarnings > 0 && !exported && <span style={{ fontSize: 12, color: "#d97706" }}>⚠ {totalWarnings} unresolved warning{totalWarnings !== 1 ? "s" : ""} — review before exporting</span>}
          </div>
          <p style={{ fontSize: 11, color: "#9ca3af", marginTop: 8 }}>Exports 4 tabs: "SOW Extract", "Billing Periods", "SKU Reference", "Audit Trail".</p>
        </div>
      )}
    </div>
  );
}
