import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";

// ---------------------------------------------------------------------------
// mock market data — swap this for real plan data plum has access to
// used by the "compare insurers" and "compare policies" tabs
// ---------------------------------------------------------------------------
const PLANS = [
  {
    id: "hdfc-optima-secure",
    insurer: "HDFC ERGO",
    plan: "Optima Secure",
    basePremium: 14200,
    claimSettlement: 98.4,
    networkHospitals: 13000,
    roomRent: "no capping",
    scores: { maternity: 62, diabetes: 88, criticalIllness: 91, general: 85 },
    notes: {
      maternity: "2-year waiting period, cover up to 1L",
      diabetes: "covers pre-existing diabetes after 2 yrs, no sub-limits",
      criticalIllness: "covers 37 illnesses, 100% payout on diagnosis",
    },
  },
  {
    id: "carehealth-supreme",
    insurer: "Care Health",
    plan: "Care Supreme",
    basePremium: 11800,
    claimSettlement: 94.2,
    networkHospitals: 18000,
    roomRent: "no capping",
    scores: { maternity: 90, diabetes: 70, criticalIllness: 80, general: 82 },
    notes: {
      maternity: "9-month waiting period, cover up to 2L, newborn day-1 cover",
      diabetes: "2-yr waiting period, moderate sub-limits on insulin",
      criticalIllness: "covers 20 illnesses, lump sum payout",
    },
  },
  {
    id: "niva-aspire",
    insurer: "Niva Bupa",
    plan: "Aspire Titanium",
    basePremium: 16500,
    claimSettlement: 91.7,
    networkHospitals: 10000,
    roomRent: "single private AC room",
    scores: { maternity: 85, diabetes: 92, criticalIllness: 75, general: 80 },
    notes: {
      maternity: "waiting period 2 yrs, cover up to 1.5L incl. complications",
      diabetes: "day-1 cover for diabetes management, no waiting period",
      criticalIllness: "covers 15 illnesses only",
    },
  },
  {
    id: "star-comprehensive",
    insurer: "Star Health",
    plan: "Comprehensive",
    basePremium: 9600,
    claimSettlement: 89.5,
    networkHospitals: 14000,
    roomRent: "1% of sum insured/day",
    scores: { maternity: 55, diabetes: 65, criticalIllness: 70, general: 76 },
    notes: {
      maternity: "4-year waiting period, low cover cap",
      diabetes: "standard 3-yr PED waiting period",
      criticalIllness: "covers 30 illnesses, tiered payout",
    },
  },
  {
    id: "icici-elevate",
    insurer: "ICICI Lombard",
    plan: "Elevate",
    basePremium: 13100,
    claimSettlement: 96.8,
    networkHospitals: 11500,
    roomRent: "no capping",
    scores: { maternity: 70, diabetes: 80, criticalIllness: 95, general: 83 },
    notes: {
      maternity: "2-yr waiting period, cover up to 1L",
      diabetes: "2-yr PED waiting, decent sub-limits",
      criticalIllness: "covers 40+ illnesses, highest CI payout ratio",
    },
  },
  {
    id: "bajaj-health-guard",
    insurer: "Bajaj Allianz",
    plan: "Health Guard",
    basePremium: 8900,
    claimSettlement: 92.1,
    networkHospitals: 9000,
    roomRent: "2% of sum insured/day",
    scores: { maternity: 48, diabetes: 60, criticalIllness: 66, general: 72 },
    notes: {
      maternity: "not recommended — long waiting period, thin cover",
      diabetes: "basic PED cover, standard terms",
      criticalIllness: "covers 12 illnesses only",
    },
  },
];

// ---------------------------------------------------------------------------
// "find me a plan" — rebuilt on tiered underwriting logic (condition → tier
// → eligible policy, with family floater structuring). uses a DIFFERENT,
// narrower policy set (the 8 named plans below) than the market-wide PLANS
// list above — that's intentional given the source rules, but worth
// flagging: the two tabs speak different "languages" right now. see note at
// the bottom of this file.
// ---------------------------------------------------------------------------

const POLICIES = {
  OS_PLUS: {
    code: "OS_PLUS",
    name: "HDFC OS+",
    insurer: "HDFC ERGO",
    tier: 1,
    maxAdults: 4,
    maxChildren: 6,
  },
  CARE_ULTIMATE: {
    code: "CARE_ULTIMATE",
    name: "Care Ultimate",
    insurer: "Care Health",
    tier: 1,
    maxAdults: 2,
    maxChildren: 2,
  },
  REASSURE_3: {
    code: "REASSURE_3",
    name: "Niva Bupa Reassure 3.0",
    insurer: "Niva Bupa",
    tier: 1,
    maxAdults: 2,
    maxChildren: 3,
  },
  ASPIRE: {
    code: "ASPIRE",
    name: "Niva Bupa Aspire",
    insurer: "Niva Bupa",
    tier: 2,
    maxAdults: 2,
    maxChildren: 2,
  },
  ELEVATE: {
    code: "ELEVATE",
    name: "ICICI Elevate",
    insurer: "ICICI Lombard",
    tier: 2,
    maxAdults: 2,
    maxChildren: 2,
  },
  SUPREME_SHINE: {
    code: "SUPREME_SHINE",
    name: "Care Supreme Shine",
    insurer: "Care Health",
    tier: 3,
    maxAdults: 2,
    maxChildren: 0,
  },
  FREEDOM: {
    code: "FREEDOM",
    name: "Care Freedom",
    insurer: "Care Health",
    tier: 3,
    maxAdults: 2,
    maxChildren: 0,
  },
  HEART: {
    code: "HEART",
    name: "Care Heart",
    insurer: "Care Health",
    tier: 3,
    maxAdults: 2,
    maxChildren: 0,
  },
};

// placeholder condition → tier/policy mapping. only the tier list itself was
// given — which specific conditions route to which policy within a tier is
// a mock for demo purposes and needs to be replaced with the real
// underwriting table before this goes anywhere near production.
const CONDITIONS = [
  { id: "cataract", label: "cataract", tier: 1, allowed: ["OS_PLUS", "CARE_ULTIMATE", "REASSURE_3"] },
  { id: "hypertension", label: "hypertension (controlled)", tier: 1, allowed: ["OS_PLUS", "CARE_ULTIMATE"] },
  { id: "thyroid", label: "thyroid disorder", tier: 1, allowed: ["OS_PLUS", "CARE_ULTIMATE", "REASSURE_3"] },
  { id: "diabetes", label: "diabetes (controlled)", tier: 2, allowed: ["ASPIRE", "ELEVATE"] },
  { id: "cardiac", label: "cardiac condition", tier: 3, allowed: ["SUPREME_SHINE", "HEART"] },
  { id: "cancer", label: "cancer (cured 2+ yrs)", tier: 3, allowed: ["FREEDOM", "HEART"] },
];

const MEMBER_RELATIONS = [
  "self",
  "spouse",
  "father",
  "mother",
  "father-in-law",
  "mother-in-law",
  "son",
  "daughter",
];

const MARRIED_PAIRS = [
  ["self", "spouse"],
  ["father", "mother"],
  ["father-in-law", "mother-in-law"],
];

function isChild(member) {
  return member.relation === "son" || member.relation === "daughter";
}

function isMarriedPair(a, b) {
  return MARRIED_PAIRS.some(
    ([r1, r2]) => (a.relation === r1 && b.relation === r2) || (a.relation === r2 && b.relation === r1)
  );
}

// core tier logic: highest tier among selected conditions wins; policies
// eligible at that tier are the intersection across those conditions, or
// the union if the intersection is empty — matches the pattern in the
// source rules (overlap narrows to the shared policy; no overlap means
// every tier-matching policy stays in play)
function eligibilityForMember(member) {
  if (member.conditions.length === 0) {
    return { tier: 0, policies: Object.keys(POLICIES), drivenBy: [] };
  }

  const conds = member.conditions.map((id) => CONDITIONS.find((c) => c.id === id));
  const maxTier = Math.max(...conds.map((c) => c.tier));
  const condsAtMaxTier = conds.filter((c) => c.tier === maxTier);

  let intersection = null;
  condsAtMaxTier.forEach((c) => {
    intersection = intersection === null ? c.allowed : intersection.filter((code) => c.allowed.includes(code));
  });

  const allowed = intersection.length > 0
    ? intersection
    : [...new Set(condsAtMaxTier.flatMap((c) => c.allowed))];

  return { tier: maxTier, policies: allowed, drivenBy: condsAtMaxTier.map((c) => c.label) };
}

function familyEligibility(members) {
  const perMember = members.map((m) => ({ member: m, ...eligibilityForMember(m) }));
  const allSets = perMember.map((p) => p.policies);
  const shared = allSets.reduce((a, b) => a.filter((code) => b.includes(code)));
  return { perMember, shared };
}

// simplified floater grouping — follows the core rules (married pairs stay
// together, children go with self/spouse, unmarried adults combine up to
// each policy's capacity) but doesn't cover every edge case in the full
// ruleset (e.g. father-in-law/mother-in-law overlaps)
function buildFloaterGroups(policyCode, members) {
  const policy = POLICIES[policyCode];
  const adults = members.filter((m) => !isChild(m));
  const children = members.filter(isChild);

  const used = new Set();
  const pairs = [];
  for (let i = 0; i < adults.length; i++) {
    for (let j = i + 1; j < adults.length; j++) {
      if (!used.has(adults[i].id) && !used.has(adults[j].id) && isMarriedPair(adults[i], adults[j])) {
        pairs.push([adults[i], adults[j]]);
        used.add(adults[i].id);
        used.add(adults[j].id);
      }
    }
  }
  const singles = adults.filter((a) => !used.has(a.id));

  let groups = pairs.map((pair) => ({ adults: [...pair], children: [], warning: null }));

  singles.forEach((s) => {
    let target =
      groups.find((g) => g.adults.length < policy.maxAdults && g.adults.some((a) => a.relation === "self")) ||
      groups.find((g) => g.adults.length < policy.maxAdults);
    if (!target) {
      target = { adults: [], children: [], warning: null };
      groups.push(target);
    }
    target.adults.push(s);
  });

  if (groups.length === 0) groups.push({ adults: [], children: [], warning: null });

  const childHomeGroup =
    groups.find((g) => g.adults.some((a) => a.relation === "self" || a.relation === "spouse")) || groups[0];

  children.forEach((c) => {
    let target = childHomeGroup.children.length < policy.maxChildren ? childHomeGroup : null;
    if (!target) target = groups.find((g) => g !== childHomeGroup && g.children.length < policy.maxChildren);
    if (!target) {
      target = { adults: [], children: [], warning: "additional floater needed — no capacity left for this child" };
      groups.push(target);
    }
    target.children.push(c);
  });

  return { groups, maxAdults: policy.maxAdults, maxChildren: policy.maxChildren };
}

// ---------------------------------------------------------------------------
// tokens
// ---------------------------------------------------------------------------
const ink = "#152420";
const inkSoft = "#4C5C56";
const bg = "#F2F4EF";
const surface = "#FFFFFF";
const line = "#DCE1D8";
const teal = "#1F6F5C";
const tealDeep = "#0F4A3C";
const amber = "#C77A2E";
const amberSoft = "#F4E2CB";
const tealSoft = "#DCEDE6";
const red = "#A3392E";
const redSoft = "#F6DEDB";

const displayFont = "'Fraunces', Georgia, serif";
const bodyFont = "'IBM Plex Sans', -apple-system, sans-serif";
const monoFont = "'IBM Plex Mono', monospace";

let nextMemberId = 2;

export default function PlumPersonalDemo() {
  const isAdminRoute = new URLSearchParams(window.location.search).has("admin");
  if (isAdminRoute) {
    return <AdminView />;
  }
  return <MainApp />;
}

function MainApp() {
  const [tab, setTab] = useState("insurers"); // "insurers" | "policies" | "recommend"
  const [sortKey, setSortKey] = useState("overall");
  const [sortDir, setSortDir] = useState("desc");
  const [insurerSortKey, setInsurerSortKey] = useState("overall");
  const [insurerSortDir, setInsurerSortDir] = useState("desc");

  const ratedPlans = useMemo(() => {
    return PLANS.map((p) => {
      const network = Math.min(100, (p.networkHospitals / 18000) * 100);
      const overall = Math.round(p.claimSettlement * 0.3 + network * 0.2 + p.scores.general * 0.5);
      return { ...p, network: Math.round(network), overall };
    });
  }, []);

  const sortedPlans = useMemo(() => {
    const arr = [...ratedPlans];
    const key = sortKey === "premium" ? "basePremium" : sortKey;
    arr.sort((a, b) => {
      const va = key === "insurer" ? a.insurer : a[key] ?? a.scores[key];
      const vb = key === "insurer" ? b.insurer : b[key] ?? b.scores[key];
      if (typeof va === "string") return sortDir === "asc" ? va.localeCompare(vb) : vb.localeCompare(va);
      return sortDir === "asc" ? va - vb : vb - va;
    });
    return arr;
  }, [ratedPlans, sortKey, sortDir]);

  const ratedInsurers = useMemo(() => {
    const groups = {};
    ratedPlans.forEach((p) => {
      if (!groups[p.insurer]) groups[p.insurer] = [];
      groups[p.insurer].push(p);
    });
    return Object.entries(groups).map(([insurer, plans]) => {
      const avg = (fn) => plans.reduce((sum, p) => sum + fn(p), 0) / plans.length;
      return {
        insurer,
        planCount: plans.length,
        overall: Math.round(avg((p) => p.overall)),
        claimSettlement: Math.round(avg((p) => p.claimSettlement) * 10) / 10,
        network: Math.round(avg((p) => p.network)),
        basePremium: Math.round(avg((p) => p.basePremium)),
        maternity: Math.round(avg((p) => p.scores.maternity)),
        diabetes: Math.round(avg((p) => p.scores.diabetes)),
        criticalIllness: Math.round(avg((p) => p.scores.criticalIllness)),
      };
    });
  }, [ratedPlans]);

  const sortedInsurers = useMemo(() => {
    const arr = [...ratedInsurers];
    const key = insurerSortKey === "premium" ? "basePremium" : insurerSortKey;
    arr.sort((a, b) => {
      const va = a[key];
      const vb = b[key];
      if (typeof va === "string") return insurerSortDir === "asc" ? va.localeCompare(vb) : vb.localeCompare(va);
      return insurerSortDir === "asc" ? va - vb : vb - va;
    });
    return arr;
  }, [ratedInsurers, insurerSortKey, insurerSortDir]);

  function toggleSort(key) {
    if (sortKey === key) setSortDir(sortDir === "asc" ? "desc" : "asc");
    else {
      setSortKey(key);
      setSortDir("desc");
    }
  }

  function toggleInsurerSort(key) {
    if (insurerSortKey === key) setInsurerSortDir(insurerSortDir === "asc" ? "desc" : "asc");
    else {
      setInsurerSortKey(key);
      setInsurerSortDir("desc");
    }
  }

  return (
    <div
      style={{
        fontFamily: bodyFont,
        background: bg,
        color: ink,
        minHeight: "100vh",
        padding: "48px 24px 80px",
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        input[type="range"] { accent-color: ${teal}; }
        button.chip { transition: all .15s ease; }
        button.chip:hover { border-color: ${teal} !important; }
        select.plum-select { font-family: ${bodyFont}; }
      `}</style>

      <div style={{ maxWidth: 880, margin: "0 auto" }}>
        {/* header */}
        <div style={{ marginBottom: 40 }}>
          <div
            style={{
              fontFamily: monoFont,
              fontSize: 12,
              letterSpacing: 2,
              color: teal,
              textTransform: "uppercase",
              marginBottom: 10,
            }}
          >
            plum personal — proof of concept
          </div>
          <h1
            style={{
              fontFamily: displayFont,
              fontWeight: 500,
              fontSize: 40,
              lineHeight: 1.15,
              margin: 0,
              maxWidth: 620,
            }}
          >
            find the health plan that actually fits your life.
          </h1>
          <p style={{ color: inkSoft, fontSize: 15, marginTop: 12, maxWidth: 560, lineHeight: 1.6 }}>
            an independent rating engine that scores every insurer and plan in the market —
            then layers on a recommendation tuned to what matters to you: maternity, diabetes,
            critical illness, or just plain value.
          </p>
        </div>

        {/* tab switcher */}
        <div
          style={{
            display: "inline-flex",
            border: `1px solid ${line}`,
            borderRadius: 10,
            padding: 4,
            marginBottom: 32,
            background: surface,
          }}
        >
          <TabButton active={tab === "insurers"} onClick={() => setTab("insurers")}>
            compare insurers
          </TabButton>
          <TabButton active={tab === "policies"} onClick={() => setTab("policies")}>
            compare policies
          </TabButton>
          <TabButton active={tab === "recommend"} onClick={() => setTab("recommend")}>
            find me a plan
          </TabButton>
        </div>

        {tab === "insurers" && (
          <InsurersView
            insurers={sortedInsurers}
            sortKey={insurerSortKey}
            sortDir={insurerSortDir}
            onSort={toggleInsurerSort}
          />
        )}
        {tab === "policies" && (
          <RatingsView plans={sortedPlans} sortKey={sortKey} sortDir={sortDir} onSort={toggleSort} />
        )}
        {tab === "recommend" && <FindMyPlanView />}
      </div>
    </div>
  );
}

const labelStyle = {
  fontSize: 12,
  fontWeight: 500,
  color: inkSoft,
  textTransform: "uppercase",
  letterSpacing: 0.6,
};

function TabButton({ active, onClick, children }) {
  return (
    <button
      onClick={onClick}
      style={{
        border: "none",
        background: active ? tealDeep : "transparent",
        color: active ? "#fff" : inkSoft,
        padding: "9px 18px",
        borderRadius: 7,
        fontFamily: bodyFont,
        fontSize: 13,
        fontWeight: 500,
        cursor: "pointer",
      }}
    >
      {children}
    </button>
  );
}

function grade(score) {
  if (score >= 90) return "A+";
  if (score >= 80) return "A";
  if (score >= 70) return "B+";
  if (score >= 60) return "B";
  return "C";
}

function InsurersView({ insurers, sortKey, sortDir, onSort }) {
  const columns = [
    { key: "insurer", label: "insurer" },
    { key: "overall", label: "overall rating" },
    { key: "claimSettlement", label: "claim settlement" },
    { key: "network", label: "network reach" },
    { key: "premium", label: "avg. annual premium" },
    { key: "maternity", label: "maternity" },
    { key: "diabetes", label: "diabetes" },
    { key: "criticalIllness", label: "critical illness" },
  ];

  return (
    <div>
      <div
        style={{
          fontFamily: monoFont,
          fontSize: 12,
          letterSpacing: 1,
          color: inkSoft,
          textTransform: "uppercase",
          marginBottom: 16,
        }}
      >
        {insurers.length} insurers rated — click a column to sort
      </div>

      <div style={{ background: surface, border: `1px solid ${line}`, borderRadius: 14, overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: bg }}>
              {columns.map((c) => (
                <th
                  key={c.key}
                  onClick={() => onSort(c.key)}
                  style={{
                    textAlign: c.key === "insurer" ? "left" : "center",
                    padding: "12px 14px",
                    fontWeight: 500,
                    color: inkSoft,
                    fontSize: 11,
                    textTransform: "uppercase",
                    letterSpacing: 0.4,
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                  }}
                >
                  {c.label}
                  {sortKey === (c.key === "premium" ? "premium" : c.key) && (
                    <span style={{ marginLeft: 4, color: teal }}>{sortDir === "asc" ? "↑" : "↓"}</span>
                  )}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {insurers.map((ins, i) => (
              <tr key={ins.insurer} style={{ borderTop: `1px solid ${line}` }}>
                <td style={{ padding: "12px 14px" }}>
                  <div style={{ fontWeight: 500 }}>{ins.insurer}</div>
                  <div style={{ color: inkSoft, fontSize: 12 }}>
                    {ins.planCount} plan{ins.planCount > 1 ? "s" : ""} rated
                  </div>
                </td>
                <td style={{ textAlign: "center" }}>
                  <span
                    style={{
                      fontFamily: monoFont,
                      fontWeight: 500,
                      background: i === 0 ? amberSoft : tealSoft,
                      color: i === 0 ? "#8A4F13" : tealDeep,
                      padding: "3px 10px",
                      borderRadius: 6,
                    }}
                  >
                    {grade(ins.overall)} · {ins.overall}
                  </span>
                </td>
                <td style={{ textAlign: "center", fontFamily: monoFont }}>{ins.claimSettlement}%</td>
                <td style={{ textAlign: "center", fontFamily: monoFont }}>{ins.network}</td>
                <td style={{ textAlign: "center", fontFamily: monoFont }}>
                  ₹{ins.basePremium.toLocaleString("en-IN")}
                </td>
                <td style={{ textAlign: "center", fontFamily: monoFont, color: inkSoft }}>{ins.maternity}</td>
                <td style={{ textAlign: "center", fontFamily: monoFont, color: inkSoft }}>{ins.diabetes}</td>
                <td style={{ textAlign: "center", fontFamily: monoFont, color: inkSoft }}>
                  {ins.criticalIllness}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <p style={{ fontSize: 12, color: inkSoft, marginTop: 14, lineHeight: 1.6 }}>
        each row is an average across that insurer's rated plans. with only one plan per
        insurer in this mock dataset the numbers match the plan-level view — this becomes
        meaningful once multiple plans per insurer are added.
      </p>
    </div>
  );
}

function RatingsView({ plans, sortKey, sortDir, onSort }) {
  const columns = [
    { key: "insurer", label: "insurer / plan" },
    { key: "overall", label: "overall rating" },
    { key: "claimSettlement", label: "claim settlement" },
    { key: "network", label: "network reach" },
    { key: "premium", label: "annual premium" },
    { key: "maternity", label: "maternity" },
    { key: "diabetes", label: "diabetes" },
    { key: "criticalIllness", label: "critical illness" },
  ];

  return (
    <div>
      <div
        style={{
          fontFamily: monoFont,
          fontSize: 12,
          letterSpacing: 1,
          color: inkSoft,
          textTransform: "uppercase",
          marginBottom: 16,
        }}
      >
        {plans.length} plans rated across {new Set(plans.map((p) => p.insurer)).size} insurers — click a
        column to sort
      </div>

      <div style={{ background: surface, border: `1px solid ${line}`, borderRadius: 14, overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: bg }}>
              {columns.map((c) => (
                <th
                  key={c.key}
                  onClick={() => onSort(c.key)}
                  style={{
                    textAlign: c.key === "insurer" ? "left" : "center",
                    padding: "12px 14px",
                    fontWeight: 500,
                    color: inkSoft,
                    fontSize: 11,
                    textTransform: "uppercase",
                    letterSpacing: 0.4,
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                  }}
                >
                  {c.label}
                  {sortKey === (c.key === "premium" ? "premium" : c.key) && (
                    <span style={{ marginLeft: 4, color: teal }}>{sortDir === "asc" ? "↑" : "↓"}</span>
                  )}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {plans.map((p, i) => (
              <tr key={p.id} style={{ borderTop: `1px solid ${line}` }}>
                <td style={{ padding: "12px 14px" }}>
                  <div style={{ fontWeight: 500 }}>{p.insurer}</div>
                  <div style={{ color: inkSoft, fontSize: 12 }}>{p.plan}</div>
                </td>
                <td style={{ textAlign: "center" }}>
                  <span
                    style={{
                      fontFamily: monoFont,
                      fontWeight: 500,
                      background: i === 0 ? amberSoft : tealSoft,
                      color: i === 0 ? "#8A4F13" : tealDeep,
                      padding: "3px 10px",
                      borderRadius: 6,
                    }}
                  >
                    {grade(p.overall)} · {p.overall}
                  </span>
                </td>
                <td style={{ textAlign: "center", fontFamily: monoFont }}>{p.claimSettlement}%</td>
                <td style={{ textAlign: "center", fontFamily: monoFont }}>
                  {p.networkHospitals.toLocaleString("en-IN")}
                </td>
                <td style={{ textAlign: "center", fontFamily: monoFont }}>
                  ₹{p.basePremium.toLocaleString("en-IN")}
                </td>
                <td style={{ textAlign: "center", fontFamily: monoFont, color: inkSoft }}>
                  {p.scores.maternity}
                </td>
                <td style={{ textAlign: "center", fontFamily: monoFont, color: inkSoft }}>
                  {p.scores.diabetes}
                </td>
                <td style={{ textAlign: "center", fontFamily: monoFont, color: inkSoft }}>
                  {p.scores.criticalIllness}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <p style={{ fontSize: 12, color: inkSoft, marginTop: 14, lineHeight: 1.6 }}>
        overall rating = 30% claim settlement ratio + 20% network reach + 50% general plan
        quality. use-case columns (maternity / diabetes / critical illness) are scored
        independently — a plan can rate lower overall but top the list for a specific need.
      </p>
    </div>
  );
}

function Chip({ active, onClick, children, tone }) {
  const activeBg = tone === "danger" ? redSoft : tealSoft;
  const activeText = tone === "danger" ? red : tealDeep;
  const activeBorder = tone === "danger" ? red : teal;
  return (
    <button
      className="chip"
      onClick={onClick}
      style={{
        padding: "8px 14px",
        borderRadius: 20,
        border: `1px solid ${active ? activeBorder : line}`,
        background: active ? activeBg : "#fff",
        color: active ? activeText : ink,
        fontSize: 13,
        fontFamily: bodyFont,
        cursor: "pointer",
      }}
    >
      {children}
    </button>
  );
}

// ---------------------------------------------------------------------------
// lead storage — Supabase (real Postgres table + file storage bucket).
// the anon key below is meant to be public in frontend code — that's normal
// for Supabase; access control is enforced by Row Level Security policies
// on the database side, not by hiding this key. see setup notes at the
// bottom of this file for the SQL to run in your Supabase project.
// ---------------------------------------------------------------------------
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://zvqldnykeihoyavrjrdv.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp2cWxkbnlrZWlob3lhdnJqcmR2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODcxMjcwODUsImV4cCI6MjEwMjcwMzA4NX0.2LFJ1TzJKruG7ul9KnrC9N73k6YtRj1Ufx3WKcrIg-0";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

async function saveLeadToServer(lead) {
  const { error } = await supabase.from("leads").insert({
    lead_name: lead.leadName,
    lead_pincode: lead.leadPincode,
    members: lead.members,
    shared_policies: lead.sharedPolicies,
  });
  if (error) throw new Error(error.message);
}

async function fetchAllLeads() {
  const { data, error } = await supabase.from("leads").select("*").order("saved_at", { ascending: false });
  if (error) throw new Error(error.message);
  return data.map((row) => ({
    leadName: row.lead_name,
    leadPincode: row.lead_pincode,
    members: row.members,
    sharedPolicies: row.shared_policies,
    savedAt: row.saved_at,
  }));
}

async function uploadCheckupFile(file) {
  const path = `checkups/${Date.now()}-${file.name}`;
  const { error } = await supabase.storage.from("checkup-documents").upload(path, file);
  if (error) throw new Error(error.message);
  return path;
}

async function getCheckupFileUrl(path) {
  const { data, error } = await supabase.storage.from("checkup-documents").createSignedUrl(path, 3600);
  if (error) throw new Error(error.message);
  return data.signedUrl;
}

function exportLeadsToExcel(leads) {
  const rows = [];
  leads.forEach((lead) => {
    lead.members.forEach((m) => {
      rows.push({
        "saved at": new Date(lead.savedAt).toLocaleString("en-IN"),
        "lead name": lead.leadName || "",
        pincode: lead.leadPincode || "",
        relation: m.relation,
        age: m.age,
        "medical conditions": m.conditions.join(", "),
        "checkup date": m.checkupDate || "",
        "checkup file": m.checkupFile?.fileName || "",
        "recommended policies": lead.sharedPolicies.map((c) => POLICIES[c]?.name).join(", ") || "none shared",
      });
    });
  });

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Customer Data");
  XLSX.writeFile(workbook, `plum-personal-leads-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// ---------------------------------------------------------------------------
// find me a plan — tiered underwriting recommendation + floater structuring
// ---------------------------------------------------------------------------
function FindMyPlanView() {
  const [leadName, setLeadName] = useState("");
  const [leadPincode, setLeadPincode] = useState("");
  const [members, setMembers] = useState([
    { id: 1, relation: "self", age: 32, conditions: [], checkupDate: "", checkupFile: null, uploading: false, uploadError: null },
  ]);
  const [submitted, setSubmitted] = useState(false);
  const [floaterPolicy, setFloaterPolicy] = useState(null);
  const [saveConfirmation, setSaveConfirmation] = useState(false);
  const [saveError, setSaveError] = useState(null);

  function addMember() {
    const id = nextMemberId++;
    setMembers([
      ...members,
      { id, relation: "spouse", age: 30, conditions: [], checkupDate: "", checkupFile: null, uploading: false, uploadError: null },
    ]);
  }

  function removeMember(id) {
    setMembers(members.filter((m) => m.id !== id));
  }

  function updateMember(id, patch) {
    setMembers(members.map((m) => (m.id === id ? { ...m, ...patch } : m)));
  }

  function toggleCondition(id, conditionId) {
    setMembers(
      members.map((m) => {
        if (m.id !== id) return m;
        const has = m.conditions.includes(conditionId);
        return { ...m, conditions: has ? m.conditions.filter((c) => c !== conditionId) : [...m.conditions, conditionId] };
      })
    );
  }

  const MAX_FILE_BYTES = 4 * 1024 * 1024; // 4MB — keeps things reasonable for a demo

  async function handleCheckupUpload(memberId, file) {
    if (!file) return;
    if (file.size > MAX_FILE_BYTES) {
      updateMember(memberId, { uploadError: "file too large — please use one under 4MB" });
      return;
    }
    updateMember(memberId, { uploading: true, uploadError: null });
    try {
      const path = await uploadCheckupFile(file);
      updateMember(memberId, {
        // the source document is stored in Supabase Storage — this just
        // keeps a reference to it. conditions and date are entered
        // manually below, not auto-extracted.
        checkupFile: { path, fileName: file.name, mediaType: file.type },
        uploading: false,
        uploadError: null,
      });
    } catch (err) {
      updateMember(memberId, { uploading: false, uploadError: err.message });
    }
  }

  async function saveLead() {
    const snapshot = {
      leadName,
      leadPincode,
      members: members.map(({ id, relation, age, conditions, checkupDate, checkupFile }) => ({
        id,
        relation,
        age,
        conditions,
        checkupDate,
        checkupFile, // { base64, mediaType, fileName } or null — the source document itself
      })),
      sharedPolicies: shared,
      savedAt: Date.now(),
    };
    setSaveError(null);
    try {
      await saveLeadToServer(snapshot);
      setSaveConfirmation(true);
      setTimeout(() => setSaveConfirmation(false), 2500);
    } catch (err) {
      setSaveError(err.message);
    }
  }

  const { perMember, shared } = useMemo(() => familyEligibility(members), [members]);

  const activeFloaterPolicy = floaterPolicy || shared[0] || null;
  const floater = useMemo(
    () => (activeFloaterPolicy ? buildFloaterGroups(activeFloaterPolicy, members) : null),
    [activeFloaterPolicy, members]
  );

  return (
    <div>
      {/* lead details */}
      <div style={{ background: surface, border: `1px solid ${line}`, borderRadius: 14, padding: "28px 32px", marginBottom: 24 }}>
        <label style={labelStyle}>lead details</label>
        <div style={{ display: "flex", gap: 12, marginTop: 10, flexWrap: "wrap" }}>
          <input
            placeholder="lead name"
            value={leadName}
            onChange={(e) => setLeadName(e.target.value)}
            style={{ padding: "9px 12px", borderRadius: 6, border: `1px solid ${line}`, fontSize: 13, fontFamily: bodyFont, flex: "1 1 200px" }}
          />
          <input
            placeholder="pincode"
            value={leadPincode}
            onChange={(e) => setLeadPincode(e.target.value)}
            style={{ padding: "9px 12px", borderRadius: 6, border: `1px solid ${line}`, fontSize: 13, fontFamily: bodyFont, width: 140 }}
          />
        </div>
      </div>

      {/* member builder */}
      <div style={{ background: surface, border: `1px solid ${line}`, borderRadius: 14, padding: "28px 32px", marginBottom: 24 }}>
        <label style={labelStyle}>who's covered</label>
        <div style={{ display: "flex", flexDirection: "column", gap: 14, marginTop: 12 }}>
          {members.map((m) => (
            <div key={m.id} style={{ border: `1px solid ${line}`, borderRadius: 10, padding: "14px 16px" }}>
              <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                <select
                  className="plum-select"
                  value={m.relation}
                  onChange={(e) => updateMember(m.id, { relation: e.target.value })}
                  style={{ padding: "7px 10px", borderRadius: 6, border: `1px solid ${line}`, fontSize: 13 }}
                >
                  {MEMBER_RELATIONS.map((r) => (
                    <option key={r} value={r}>
                      {r}
                    </option>
                  ))}
                </select>
                <input
                  type="number"
                  min={0}
                  max={100}
                  value={m.age}
                  onChange={(e) => updateMember(m.id, { age: +e.target.value })}
                  style={{ width: 64, padding: "7px 10px", borderRadius: 6, border: `1px solid ${line}`, fontSize: 13 }}
                />
                <span style={{ fontSize: 12, color: inkSoft }}>yrs</span>
                {members.length > 1 && (
                  <button
                    onClick={() => removeMember(m.id)}
                    style={{ marginLeft: "auto", border: "none", background: "none", color: red, fontSize: 12, cursor: "pointer" }}
                  >
                    remove
                  </button>
                )}
              </div>

              <div style={{ marginTop: 12 }}>
                <div style={{ fontSize: 11, color: inkSoft, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 6 }}>
                  upload health checkup — image or pdf (optional, stored with this lead)
                </div>
                <input
                  type="file"
                  accept="image/*,application/pdf"
                  onChange={(e) => handleCheckupUpload(m.id, e.target.files[0])}
                  style={{ fontSize: 12, fontFamily: bodyFont }}
                  disabled={m.uploading}
                />
                {m.uploading && <div style={{ fontSize: 12, color: teal, marginTop: 6 }}>uploading…</div>}
                {m.uploadError && <div style={{ fontSize: 12, color: red, marginTop: 6 }}>couldn't upload: {m.uploadError}</div>}
                {m.checkupFile && !m.uploading && (
                  <div style={{ fontSize: 12, color: inkSoft, marginTop: 6 }}>
                    ✓ attached: <span style={{ fontFamily: monoFont }}>{m.checkupFile.fileName}</span>
                  </div>
                )}
              </div>

              <div style={{ marginTop: 10 }}>
                <div style={{ fontSize: 11, color: inkSoft, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 6 }}>
                  checkup date (enter manually)
                </div>
                <input
                  type="date"
                  value={m.checkupDate || ""}
                  onChange={(e) => updateMember(m.id, { checkupDate: e.target.value })}
                  style={{ padding: "7px 10px", borderRadius: 6, border: `1px solid ${line}`, fontSize: 13, fontFamily: bodyFont }}
                />
              </div>

              <div style={{ marginTop: 10 }}>
                <div style={{ fontSize: 11, color: inkSoft, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 6 }}>
                  medical conditions
                </div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {CONDITIONS.map((c) => (
                    <Chip key={c.id} active={m.conditions.includes(c.id)} onClick={() => toggleCondition(m.id, c.id)}>
                      {c.label}
                    </Chip>
                  ))}
                </div>
              </div>
            </div>
          ))}
        </div>

        <button
          onClick={addMember}
          style={{
            marginTop: 14,
            background: "#fff",
            color: tealDeep,
            border: `1px solid ${teal}`,
            borderRadius: 8,
            padding: "8px 16px",
            fontFamily: bodyFont,
            fontWeight: 500,
            fontSize: 13,
            cursor: "pointer",
          }}
        >
          + add family member
        </button>

        <div style={{ display: "flex", gap: 12, marginTop: 24, flexWrap: "wrap" }}>
          <button
            onClick={() => {
              setSubmitted(true);
              setFloaterPolicy(null);
            }}
            style={{
              background: tealDeep,
              color: "#fff",
              border: "none",
              borderRadius: 8,
              padding: "13px 24px",
              fontFamily: bodyFont,
              fontWeight: 500,
              fontSize: 14,
              cursor: "pointer",
            }}
          >
            get recommendation →
          </button>

          <button
            onClick={saveLead}
            style={{
              background: "#fff",
              color: tealDeep,
              border: `1px solid ${teal}`,
              borderRadius: 8,
              padding: "13px 24px",
              fontFamily: bodyFont,
              fontWeight: 500,
              fontSize: 14,
              cursor: "pointer",
            }}
          >
            {saveConfirmation ? "saved ✓" : "save lead"}
          </button>
          {saveError && <span style={{ color: red, fontSize: 12, alignSelf: "center" }}>couldn't save: {saveError}</span>}
        </div>
      </div>

      {submitted && (
        <div>
          <div style={{ fontFamily: monoFont, fontSize: 12, letterSpacing: 1, color: inkSoft, textTransform: "uppercase", marginBottom: 12 }}>
            recommended policies
          </div>

          {shared.length > 0 ? (
            <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 24 }}>
              {shared.map((code) => (
                <PolicyCard
                  key={code}
                  code={code}
                  selected={activeFloaterPolicy === code}
                  onClick={() => setFloaterPolicy(code)}
                  perMember={perMember}
                />
              ))}
            </div>
          ) : (
            <div style={{ background: redSoft, border: `1px solid ${red}`, borderRadius: 10, padding: "16px 20px", marginBottom: 24 }}>
              <div style={{ fontWeight: 500, color: red, marginBottom: 6 }}>no single policy covers everyone</div>
              <p style={{ fontSize: 13, color: ink, margin: 0, lineHeight: 1.6 }}>
                the family's conditions don't share an eligible policy. here's what fits each member individually:
              </p>
              <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 6 }}>
                {perMember.map(({ member, tier, policies }) => (
                  <div key={member.id} style={{ fontSize: 13 }}>
                    <span style={{ fontWeight: 500, textTransform: "capitalize" }}>{member.relation}</span>
                    <span style={{ color: inkSoft }}> — tier {tier}, eligible: </span>
                    <span style={{ fontFamily: monoFont }}>{policies.join(", ") || "none"}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {floater && (
            <div>
              <div style={{ fontFamily: monoFont, fontSize: 12, letterSpacing: 1, color: inkSoft, textTransform: "uppercase", marginBottom: 12 }}>
                suggested floater structure — {POLICIES[activeFloaterPolicy].name}
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {floater.groups.map((g, i) => (
                  <div key={i} style={{ background: surface, border: `1px solid ${line}`, borderRadius: 10, padding: "14px 18px" }}>
                    <div style={{ fontSize: 12, color: inkSoft, marginBottom: 6 }}>
                      plan {i + 1} · {g.adults.length}/{floater.maxAdults} adults · {g.children.length}/{floater.maxChildren} children
                    </div>
                    <div style={{ fontSize: 13, display: "flex", flexWrap: "wrap", gap: "4px 10px" }}>
                      {[...g.adults, ...g.children].map((m) => (
                        <span key={m.id} style={{ textTransform: "capitalize" }}>
                          {m.relation} ({m.age})
                        </span>
                      ))}
                    </div>
                    {g.warning && <div style={{ fontSize: 12, color: red, marginTop: 6 }}>{g.warning}</div>}
                  </div>
                ))}
              </div>
              <p style={{ fontSize: 12, color: inkSoft, marginTop: 12, lineHeight: 1.6 }}>
                structuring follows the core grouping rules (married pairs together, children with
                self/spouse, unmarried-adult handling per policy) — always verify in-law and edge
                cases against the full underwriting ruleset before using this with a customer.
              </p>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function PolicyCard({ code, selected, onClick, perMember }) {
  const policy = POLICIES[code];
  const drivers = [...new Set(perMember.flatMap((p) => p.drivenBy))];
  return (
    <button
      onClick={onClick}
      style={{
        textAlign: "left",
        background: selected ? tealSoft : surface,
        border: selected ? `2px solid ${teal}` : `1px solid ${line}`,
        borderRadius: 12,
        padding: "16px 20px",
        cursor: "pointer",
        fontFamily: bodyFont,
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <span
          style={{
            fontFamily: monoFont,
            fontSize: 11,
            background: amberSoft,
            color: "#8A4F13",
            padding: "2px 8px",
            borderRadius: 4,
          }}
        >
          tier {policy.tier}
        </span>
        <span style={{ fontFamily: displayFont, fontSize: 17, fontWeight: 500, color: ink }}>{policy.name}</span>
        <span style={{ fontSize: 12, color: inkSoft }}>{policy.insurer}</span>
      </div>
      <div style={{ fontSize: 12, color: inkSoft, marginTop: 6 }}>
        default structure: {policy.maxAdults} adults + {policy.maxChildren} children
        {drivers.length > 0 && <> · driven by: {drivers.join(", ")}</>}
      </div>
    </button>
  );
}

// ---------------------------------------------------------------------------
// admin-only view — reachable at yoursite.netlify.app/?admin=<ADMIN_KEY>
// nobody using the normal app ever sees this route or a link to it.
// ---------------------------------------------------------------------------
function AdminView() {
  const [session, setSession] = useState(undefined); // undefined = checking, null = logged out
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState(null);
  const [loggingIn, setLoggingIn] = useState(false);

  const [status, setStatus] = useState("idle"); // "idle" | "loading" | "ok" | "error"
  const [leads, setLeads] = useState([]);
  const [error, setError] = useState(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => setSession(newSession));
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) return;
    setStatus("loading");
    fetchAllLeads()
      .then((data) => {
        setLeads(data);
        setStatus("ok");
      })
      .catch((err) => {
        setError(err.message);
        setStatus("error");
      });
  }, [session]);

  async function handleLogin(e) {
    e.preventDefault();
    setLoginError(null);
    setLoggingIn(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    setLoggingIn(false);
    if (error) setLoginError(error.message);
  }

  async function handleLogout() {
    await supabase.auth.signOut();
  }

  const totalMembers = leads.reduce((sum, l) => sum + l.members.length, 0);

  const shellStyle = { fontFamily: bodyFont, background: bg, color: ink, minHeight: "100vh", padding: "48px 24px 80px" };
  const fontImport = (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
      * { box-sizing: border-box; }
    `}</style>
  );

  // still checking whether a session already exists
  if (session === undefined) {
    return (
      <div style={shellStyle}>
        {fontImport}
        <div style={{ maxWidth: 420, margin: "0 auto" }}>
          <p style={{ color: inkSoft }}>checking session…</p>
        </div>
      </div>
    );
  }

  // not logged in — show a real login form, not a URL secret
  if (!session) {
    return (
      <div style={shellStyle}>
        {fontImport}
        <div style={{ maxWidth: 420, margin: "0 auto" }}>
          <div style={{ fontFamily: monoFont, fontSize: 12, letterSpacing: 2, color: teal, textTransform: "uppercase", marginBottom: 10 }}>
            admin
          </div>
          <h1 style={{ fontFamily: displayFont, fontWeight: 500, fontSize: 28, margin: "0 0 20px" }}>log in</h1>
          <form onSubmit={handleLogin} style={{ background: surface, border: `1px solid ${line}`, borderRadius: 14, padding: "24px 28px" }}>
            <label style={labelStyle}>email</label>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              style={{ width: "100%", padding: "9px 12px", borderRadius: 6, border: `1px solid ${line}`, fontSize: 13, fontFamily: bodyFont, marginTop: 6, marginBottom: 16 }}
              required
            />
            <label style={labelStyle}>password</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              style={{ width: "100%", padding: "9px 12px", borderRadius: 6, border: `1px solid ${line}`, fontSize: 13, fontFamily: bodyFont, marginTop: 6, marginBottom: 16 }}
              required
            />
            {loginError && <p style={{ color: red, fontSize: 12, marginTop: -8, marginBottom: 14 }}>{loginError}</p>}
            <button
              type="submit"
              disabled={loggingIn}
              style={{ background: tealDeep, color: "#fff", border: "none", borderRadius: 8, padding: "11px 20px", fontFamily: bodyFont, fontWeight: 500, fontSize: 14, cursor: "pointer", width: "100%" }}
            >
              {loggingIn ? "logging in…" : "log in"}
            </button>
          </form>
          <p style={{ fontSize: 12, color: inkSoft, marginTop: 14, lineHeight: 1.6 }}>
            this account is created directly in your Supabase project's dashboard (Authentication → Users) — there's no public sign-up form here.
          </p>
        </div>
      </div>
    );
  }

  // logged in — show the data
  return (
    <div style={shellStyle}>
      {fontImport}
      <div style={{ maxWidth: 720, margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 10 }}>
          <div style={{ fontFamily: monoFont, fontSize: 12, letterSpacing: 2, color: teal, textTransform: "uppercase" }}>admin</div>
          <button onClick={handleLogout} style={{ border: "none", background: "none", color: inkSoft, fontSize: 12, cursor: "pointer" }}>
            log out
          </button>
        </div>
        <h1 style={{ fontFamily: displayFont, fontWeight: 500, fontSize: 32, margin: "0 0 24px" }}>saved leads</h1>

        {status === "loading" && <p style={{ color: inkSoft }}>loading…</p>}

        {status === "error" && (
          <div style={{ background: redSoft, border: `1px solid ${red}`, borderRadius: 10, padding: "16px 20px" }}>
            <div style={{ fontWeight: 500, color: red }}>couldn't load leads</div>
            <p style={{ fontSize: 13, margin: "6px 0 0" }}>{error}</p>
          </div>
        )}

        {status === "ok" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
              <div style={{ fontSize: 13, color: inkSoft }}>
                {leads.length} lead{leads.length !== 1 ? "s" : ""} saved · {totalMembers} member{totalMembers !== 1 ? "s" : ""} total
              </div>
              {leads.length > 0 && (
                <button
                  onClick={() => exportLeadsToExcel(leads)}
                  style={{ background: tealDeep, color: "#fff", border: "none", borderRadius: 8, padding: "10px 18px", fontFamily: bodyFont, fontWeight: 500, fontSize: 13, cursor: "pointer" }}
                >
                  ↓ export all to excel
                </button>
              )}
            </div>

            {leads.length === 0 ? (
              <p style={{ fontSize: 13, color: inkSoft }}>no leads saved yet.</p>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {leads.map((lead, i) => (
                  <div key={i} style={{ background: surface, border: `1px solid ${line}`, borderRadius: 10, padding: "14px 18px", fontSize: 13 }}>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontWeight: 500 }}>{lead.leadName || "unnamed lead"}</span>
                      <span style={{ color: inkSoft, fontSize: 12 }}>{new Date(lead.savedAt).toLocaleString("en-IN")}</span>
                    </div>
                    <div style={{ color: inkSoft, fontSize: 12, marginTop: 4 }}>
                      pincode {lead.leadPincode || "—"} · {lead.members.length} member{lead.members.length > 1 ? "s" : ""}
                    </div>
                    <div style={{ marginTop: 8, display: "flex", flexDirection: "column", gap: 4 }}>
                      {lead.members.map((m) => (
                        <div key={m.id} style={{ fontSize: 12 }}>
                          <span style={{ textTransform: "capitalize" }}>{m.relation} ({m.age})</span>
                          {m.conditions.length > 0 && <span style={{ color: inkSoft }}> — {m.conditions.join(", ")}</span>}
                          {m.checkupDate && <span style={{ color: inkSoft }}> · checkup {m.checkupDate}</span>}
                          {m.checkupFile && (
                            <>
                              {" "}·{" "}
                              <button
                                onClick={async () => {
                                  try {
                                    const url = await getCheckupFileUrl(m.checkupFile.path);
                                    window.open(url, "_blank");
                                  } catch {
                                    alert("couldn't open file — it may have been removed from storage");
                                  }
                                }}
                                style={{ border: "none", background: "none", color: teal, cursor: "pointer", fontSize: 12, padding: 0, textDecoration: "underline" }}
                              >
                                view {m.checkupFile.fileName}
                              </button>
                            </>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
