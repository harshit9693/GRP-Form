import React, { useState, useMemo, useRef } from "react";
import {
  Plus, Pencil, Trash2, X, Search, ChevronRight, AlertTriangle,
  Check, CheckCircle2, XCircle, RotateCcw, Lock, Upload,
  ChevronLeft, ChevronsLeft, ChevronsRight, ArrowUp, ArrowDown, ArrowUpDown
} from "lucide-react";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------------
   SEED DATA
   Transcribed from the four source forms (State / District /
   GRP Station / GRP Station–RPF Link). Child tables reference
   their parent by CODE (not name) so renaming a parent anywhere
   propagates everywhere it's used — this is what keeps the four
   tables genuinely interlinked instead of four separate lists.
   Swap the useState initializers for API calls in production.

   IMPORTANT: every "code" field (state code, district code,
   station code, etc.) is stored and compared as a STRING
   throughout this file — never coerced to Number — because codes
   can arrive from manual entry or Excel import with leading zeros
   or mixed formats, and coercion would silently corrupt them.
----------------------------------------------------------------*/

const STATE_TUPLES = [];
const RAW_STATES = STATE_TUPLES.map(([code, name, strCode, address]) => ({
  code: String(code), name, strCode, address: address || name,
  status: name === "Test State" ? "pending" : "correct",
}));
const stateCodeByName = Object.fromEntries(RAW_STATES.map((s) => [s.name, s.code]));

const DISTRICT_TUPLES = [];
const RAW_DISTRICTS = DISTRICT_TUPLES.map(([stateName, code, name, strCode, address]) => ({
  stateCode: stateCodeByName[stateName], code: String(code), name, strCode: strCode || "", address: address || name,
  shiftType: "", status: stateName === "Test State" ? "pending" : "correct",
}));
const districtKey = (stateCode, name) => `${stateCode}::${name}`;
const districtCodeByKey = Object.fromEntries(RAW_DISTRICTS.map((d) => [districtKey(d.stateCode, d.name), d.code]));

const STATION_TUPLES = [];
const RAW_STATIONS = STATION_TUPLES.map(([stateName, districtName, code, name, strCode, address]) => {
  const stateCode = stateCodeByName[stateName];
  const districtCode = districtCodeByKey[districtKey(stateCode, districtName)];
  return {
    stateCode, districtCode, code: String(code), name, strCode, address: address || name,
    status: stateName === "Test State" ? "pending" : "correct",
  };
});
const stationKey = (stateCode, districtCode, name) => `${stateCode}::${districtCode}::${name}`;
const stationCodeByKey = Object.fromEntries(RAW_STATIONS.map((s) => [stationKey(s.stateCode, s.districtCode, s.name), s.code]));

// RPF reference hierarchy — read-only here, managed by its own master elsewhere.
const ZONES = [{ name: "TRAINING ZONE" }, { name: "EASTERN RAILWAY" }];
const DIVISIONS = [
  { zone: "TRAINING ZONE", name: "TRAINING DIVISION1" },
  { zone: "EASTERN RAILWAY", name: "Asansol" },
];
const POSTS = [
  { division: "TRAINING DIVISION1", name: "TRG-DIV1-POST1" },
  { division: "Asansol", name: "Andal(Line)" }, { division: "Asansol", name: "ANDAL YARD" },
  { division: "Asansol", name: "Ukhara" }, { division: "Asansol", name: "SIURI" },
  { division: "Asansol", name: "Asansol(East)" }, { division: "Asansol", name: "Asansol(West)" },
  { division: "Asansol", name: "Sitarampur" }, { division: "Asansol", name: "Barakar" },
  { division: "Asansol", name: "BARACHAK" }, { division: "Asansol", name: "Panagarh" },
];

const LINK_TUPLES = [];
const RAW_LINKS = LINK_TUPLES.map(([stateName, districtName, stationName, zone, division, post], i) => {
  const stateCode = stateCodeByName[stateName];
  const districtCode = districtCodeByKey[districtKey(stateCode, districtName)];
  const stationCode = stationCodeByKey[stationKey(stateCode, districtCode, stationName)];
  return {
    id: i + 1, stateCode, districtCode, stationCode, zone, division, post,
    status: stateName === "Test State" ? "pending" : "correct",
  };
});

/* ------------------------------- THEME -------------------------------- */
const FONT_LINK =
  "https://fonts.googleapis.com/css2?family=Barlow+Semi+Condensed:wght@600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500&display=swap";

const C = {
  navy: "#1B3358", navyDeep: "#132540", amber: "#D98E04", amberDeep: "#B67600",
  bg: "#F6F4EF", card: "#FFFFFF", ink: "#20242B", sub: "#6B6558",
  line: "#DFDACC", danger: "#B23A34", success: "#2F7D5C",
};

const TABS = [
  { id: "state", label: "State", plural: "States" },
  { id: "district", label: "District", plural: "Districts" },
  { id: "station", label: "GRP Station", plural: "GRP Stations" },
  { id: "link", label: "GRP–RPF Link", plural: "Links" },
];

// Expected column layout per tab (row 1 = report title, row 2 = header, data from row 3).
const EXCEL_COLUMNS = {
  state: ["State Code", "State Name", "State Str Code", "Location Address"],
  district: ["State Name", "District Code", "District Name", "District Str Code", "Location Address"],
  station: ["State Name", "District Name", "Station Code", "Station Name", "Station Str Code", "Location Address"],
  link: ["Sno", "GRP Unit (State/District/Station)", "RPF Unit (Zone/Division/Post)"],
};

const ITEMS_PER_PAGE = 20;

/* ------------------------------ FILE PARSING -----------------------------*/
// Reads an .xlsx / .xls workbook (via SheetJS) and returns a plain
// array-of-string-arrays — the sheet exactly as it appears, every cell
// coerced to a string so codes with leading zeros etc. survive intact.
async function readExcelFile(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
  return rows
    .map((r) => r.map((c) => String(c ?? "").trim()))
    .filter((r) => r.some((c) => c !== ""));
}

const normalizeName = (s) => String(s || "").trim().toLowerCase();

/* ------------------------------ SMALL UI -------------------------------*/
function Field({ label, children, required }) {
  return (
    <label className="block mb-4">
      <span className="block text-sm font-medium mb-1" style={{ color: C.ink }}>
        {label}{required && <span style={{ color: C.danger }}> *</span>}
      </span>
      {children}
    </label>
  );
}

const inputBase = { borderColor: C.line, color: C.ink, fontFamily: "Inter, sans-serif" };

function TextInput({ value, onChange, placeholder, disabled }) {
  return (
    <input
      value={value}
      disabled={disabled}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      className="w-full rounded-md border px-3 py-2 text-sm focus:outline-none focus:ring-2"
      style={{ ...inputBase, background: disabled ? "#EFEBE0" : "#fff", color: disabled ? C.sub : C.ink }}
      onFocus={(e) => !disabled && (e.target.style.borderColor = C.amber)}
      onBlur={(e) => (e.target.style.borderColor = C.line)}
    />
  );
}

function Select({ value, onChange, options, placeholder = "-Select one-", disabled }) {
  const norm = options.map((o) => (typeof o === "string" ? { value: o, label: o } : o));
  return (
    <select
      value={value === undefined || value === null || value === "" ? "" : String(value)}
      disabled={disabled}
      onChange={(e) => onChange(e.target.value)}
      className="w-full rounded-md border px-3 py-2 text-sm bg-white focus:outline-none"
      style={{ ...inputBase, opacity: disabled ? 0.55 : 1, background: disabled ? "#EFEBE0" : "#fff" }}
    >
      <option value="">{placeholder}</option>
      {norm.map((o) => (
        <option key={o.value} value={String(o.value)}>{o.label}</option>
      ))}
    </select>
  );
}

function Crumb({ parts }) {
  return (
    <div className="flex items-center flex-wrap gap-1 text-xs mb-4 px-3 py-2 rounded-md" style={{ background: C.bg, border: `1px dashed ${C.line}` }}>
      {parts.map((p, i) => (
        <React.Fragment key={i}>
          {i > 0 && <ChevronRight size={12} style={{ color: C.sub }} />}
          <span style={{ color: p ? C.navy : C.sub, fontFamily: "IBM Plex Mono, monospace", fontWeight: 500 }}>
            {p || "—"}
          </span>
        </React.Fragment>
      ))}
    </div>
  );
}

/* -------------------------------- FORMS --------------------------------*/
function StateForm({ data, set, mode }) {
  return (
    <>
      <Field label="State Code" required>
        <TextInput value={data.code} disabled={mode === "edit"} onChange={(v) => set("code", v.replace(/\D/g, ""))} placeholder="Enter State Code" />
      </Field>
      <Field label="State Name" required>
        <TextInput value={data.name} onChange={(v) => set("name", v.toUpperCase())} placeholder="Enter State Name" />
      </Field>
      <Field label="State Str Code" required>
        <TextInput value={data.strCode} onChange={(v) => set("strCode", v.toUpperCase())} placeholder="Enter State Str Code" />
      </Field>
      <Field label="Location Address" required>
        <TextInput value={data.address} onChange={(v) => set("address", v)} placeholder="Enter Location Address" />
      </Field>
    </>
  );
}

/* ----------------------------- PAGINATION CONTROLS ------------------------ */
function TablePagination({ currentPage, totalItems, onPageChange }) {
  const totalPages = Math.ceil(totalItems / ITEMS_PER_PAGE) || 1;
  if (totalPages <= 1) return null;

  return (
    <div className="flex items-center justify-between px-4 py-3" style={{ borderTop: `1px solid ${C.line}`, background: C.bg }}>
      <div className="text-xs" style={{ color: C.sub }}>
        Showing <span className="font-semibold">{Math.min((currentPage - 1) * ITEMS_PER_PAGE + 1, totalItems)}</span> to{" "}
        <span className="font-semibold">{Math.min(currentPage * ITEMS_PER_PAGE, totalItems)}</span> of{" "}
        <span className="font-semibold">{totalItems}</span> entries
      </div>
      <div className="flex items-center gap-1">
        <button
          onClick={() => onPageChange(1)}
          disabled={currentPage === 1}
          className="p-1 rounded hover:bg-black/[0.05] disabled:opacity-40"
          style={{ color: C.navy }}
        >
          <ChevronsLeft size={16} />
        </button>
        <button
          onClick={() => onPageChange(currentPage - 1)}
          disabled={currentPage === 1}
          className="p-1 rounded hover:bg-black/[0.05] disabled:opacity-40 mr-1"
          style={{ color: C.navy }}
        >
          <ChevronLeft size={16} />
        </button>
        <span className="text-xs px-2 py-1 rounded" style={{ background: C.card, color: C.ink, border: `1px solid ${C.line}` }}>
          Page <b>{currentPage}</b> of <b>{totalPages}</b>
        </span>
        <button
          onClick={() => onPageChange(currentPage + 1)}
          disabled={currentPage === totalPages}
          className="p-1 rounded hover:bg-black/[0.05] disabled:opacity-40 ml-1"
          style={{ color: C.navy }}
        >
          <ChevronRight size={16} />
        </button>
        <button
          onClick={() => onPageChange(totalPages)}
          disabled={currentPage === totalPages}
          className="p-1 rounded hover:bg-black/[0.05] disabled:opacity-40"
          style={{ color: C.navy }}
        >
          <ChevronsRight size={16} />
        </button>
      </div>
    </div>
  );
}

function DistrictForm({ data, set, mode, states }) {
  const stateOptions = states.filter((s) => s.status === "correct").map((s) => ({ value: s.code, label: s.name }));
  const stateName = states.find((s) => s.code === data.stateCode)?.name || "";
  return (
    <>
      <Field label="State Name" required>
        <Select value={data.stateCode} disabled={mode === "edit"} onChange={(v) => set("stateCode", v)} options={stateOptions} />
        {mode === "add" && stateOptions.length === 0 && (
          <span className="block text-xs mt-1" style={{ color: C.danger }}>No verified states yet — mark a state Correct on the State tab first.</span>
        )}
      </Field>
      <Field label="District Code" required>
        <TextInput value={data.code} disabled placeholder="Auto-generated once a state is picked" />
      </Field>
      <Field label="District Name" required>
        <TextInput value={data.name} onChange={(v) => set("name", v)} placeholder="Enter District Name" />
      </Field>
      <Field label="District Str Code" required>
        <TextInput value={data.strCode} onChange={(v) => set("strCode", v)} placeholder="Enter District Str Code" />
      </Field>
      <Field label="Location Address" required>
        <TextInput value={data.address} onChange={(v) => set("address", v)} placeholder="Enter Location Address" />
      </Field>
      <Field label="Shift Type" required>
        <Select value={data.shiftType || ""} onChange={(v) => set("shiftType", v)} options={["Shift A", "Shift B"]} />
      </Field>
      <Crumb parts={[stateName, data.name || "District"]} />
    </>
  );
}

function StationForm({ data, set, mode, states, districts }) {
  const stateOptions = states.filter((s) => s.status === "correct").map((s) => ({ value: s.code, label: s.name }));
  const districtOptions = districts
    .filter((d) => d.stateCode === data.stateCode && d.status === "correct")
    .map((d) => ({ value: d.code, label: d.name }));
  const stateName = states.find((s) => s.code === data.stateCode)?.name || "";
  const districtName = districts.find((d) => d.stateCode === data.stateCode && d.code === data.districtCode)?.name || "";
  return (
    <>
      <Field label="State Name" required>
        <Select value={data.stateCode} disabled={mode === "edit"} onChange={(v) => set("stateCode", v)} options={stateOptions} />
      </Field>
      <Field label="District" required>
        <Select value={data.districtCode} disabled={mode === "edit" || !data.stateCode} onChange={(v) => set("districtCode", v)} options={districtOptions} />
        {mode === "add" && data.stateCode && districtOptions.length === 0 && (
          <span className="block text-xs mt-1" style={{ color: C.danger }}>No verified districts for this state yet — verify one on the District tab first.</span>
        )}
      </Field>
      <Field label="Station Code" required>
        <TextInput value={data.code} disabled placeholder="Auto-generated once a district is picked" />
      </Field>
      <Field label="Station Name" required>
        <TextInput value={data.name} onChange={(v) => set("name", v)} placeholder="Enter Station Name" />
      </Field>
      <Field label="Station Str Code" required>
        <TextInput value={data.strCode} onChange={(v) => set("strCode", v)} placeholder="Enter Station Str Code" />
      </Field>
      <Field label="Location Address" required>
        <TextInput value={data.address} onChange={(v) => set("address", v)} placeholder="Enter Location Address" />
      </Field>
      <Crumb parts={[stateName, districtName, data.name || "Station"]} />
    </>
  );
}

function LinkForm({ data, set, states, districts, stations }) {
  const stateOptions = states.filter((s) => s.status === "correct").map((s) => ({ value: s.code, label: s.name }));
  const districtOptions = districts
    .filter((d) => d.stateCode === data.stateCode && d.status === "correct")
    .map((d) => ({ value: d.code, label: d.name }));
  const stationOptions = stations
    .filter((s) => s.stateCode === data.stateCode && s.districtCode === data.districtCode && s.status === "correct")
    .map((s) => ({ value: s.code, label: s.name }));
  const divisionOptions = DIVISIONS.filter((d) => d.zone === data.zone).map((d) => d.name);
  const postOptions = POSTS.filter((p) => p.division === data.division).map((p) => p.name);

  const stateName = states.find((s) => s.code === data.stateCode)?.name || "";
  const districtName = districts.find((d) => d.stateCode === data.stateCode && d.code === data.districtCode)?.name || "";
  const stationName = stations.find((s) => s.stateCode === data.stateCode && s.districtCode === data.districtCode && s.code === data.stationCode)?.name || "";

  return (
    <>
      <div className="text-xs font-semibold uppercase tracking-wide mb-2" style={{ color: C.amberDeep, fontFamily: "Barlow Semi Condensed, sans-serif" }}>GRP Station</div>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="State" required>
          <Select value={data.stateCode} onChange={(v) => { set("stateCode", v); set("districtCode", ""); set("stationCode", ""); }} options={stateOptions} />
        </Field>
        <Field label="District" required>
          <Select value={data.districtCode} onChange={(v) => { set("districtCode", v); set("stationCode", ""); }} options={districtOptions} disabled={!data.stateCode} />
        </Field>
      </div>
      <Field label="Station" required>
        <Select value={data.stationCode} onChange={(v) => set("stationCode", v)} options={stationOptions} disabled={!data.districtCode} />
        {data.districtCode && stationOptions.length === 0 && (
          <span className="block text-xs mt-1" style={{ color: C.danger }}>No verified GRP stations for this district yet — verify one on the GRP Station tab first.</span>
        )}
      </Field>
      <Crumb parts={[stateName, districtName, stationName]} />

      <div className="text-xs font-semibold uppercase tracking-wide mb-2 mt-5" style={{ color: C.amberDeep, fontFamily: "Barlow Semi Condensed, sans-serif" }}>RPF Post</div>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Zone" required>
          <Select value={data.zone} onChange={(v) => { set("zone", v); set("division", ""); set("post", ""); }} options={ZONES.map((z) => z.name)} />
        </Field>
        <Field label="Division" required>
          <Select value={data.division} onChange={(v) => { set("division", v); set("post", ""); }} options={divisionOptions} disabled={!data.zone} />
        </Field>
      </div>
      <Field label="Post" required>
        <Select value={data.post} onChange={(v) => set("post", v)} options={postOptions} disabled={!data.division} />
      </Field>
      <Crumb parts={[data.zone, data.division, data.post]} />
    </>
  );
}

/* -------------------------------- MODAL ---------------------------------*/
function Modal({ tab, mode, data, set, onClose, onSubmit, onReset, states, districts, stations, error }) {
  const titleMap = { state: "State Details", district: "District Details", station: "GRP Station Details", link: "Link GRP Station With RPF Post" };
  return (
    <div className="fixed inset-0 z-50 flex items-start md:items-center justify-center p-3 overflow-y-auto" style={{ background: "rgba(19,37,64,0.55)" }}>
      <div className="w-full rounded-lg overflow-hidden shadow-2xl" style={{ maxWidth: 480, background: C.card, marginTop: 24, marginBottom: 24 }}>
        <div className="flex items-center justify-between px-5 py-4" style={{ background: C.navy }}>
          <h2 className="text-white text-base font-semibold" style={{ fontFamily: "Barlow Semi Condensed, sans-serif", letterSpacing: 0.3 }}>
            {mode === "edit" ? "Edit" : "Add New"} {titleMap[tab]}
          </h2>
          <button onClick={onClose} aria-label="Close" className="text-white/80 hover:text-white">
            <X size={20} />
          </button>
        </div>
        <div className="px-5 pt-5 pb-2 max-h-[70vh] overflow-y-auto">
          {error && (
            <div className="mb-4 px-3 py-2 rounded-md text-sm flex items-center gap-2" style={{ background: "#FBEAE9", color: C.danger }}>
              <AlertTriangle size={16} /> {error}
            </div>
          )}
          {tab === "state" && <StateForm data={data} set={set} mode={mode} />}
          {tab === "district" && <DistrictForm data={data} set={set} mode={mode} states={states} />}
          {tab === "station" && <StationForm data={data} set={set} mode={mode} states={states} districts={districts} />}
          {tab === "link" && <LinkForm data={data} set={set} states={states} districts={districts} stations={stations} />}
        </div>
        <div className="flex gap-2 px-5 py-4" style={{ borderTop: `1px solid ${C.line}` }}>
          <button onClick={onSubmit} className="px-5 py-2 rounded-md text-white text-sm font-semibold" style={{ background: C.navy }}>Submit</button>
          <button onClick={onReset} className="px-5 py-2 rounded-md text-sm font-semibold" style={{ background: C.bg, color: C.ink, border: `1px solid ${C.line}` }}>Reset</button>
          <button onClick={onClose} className="ml-auto px-4 py-2 rounded-md text-sm" style={{ color: C.sub }}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

function ConfirmDialog({ title, message, confirmLabel, onCancel, onConfirm, tone = "danger" }) {
  const toneColor = tone === "danger" ? C.danger : C.navy;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3" style={{ background: "rgba(19,37,64,0.55)" }}>
      <div className="w-full rounded-lg shadow-2xl p-5" style={{ maxWidth: 380, background: C.card }}>
        <div className="flex items-center gap-2 mb-2" style={{ color: toneColor }}>
          <AlertTriangle size={20} />
          <h3 className="font-semibold text-base" style={{ fontFamily: "Barlow Semi Condensed, sans-serif" }}>{title}</h3>
        </div>
        <p className="text-sm mb-5" style={{ color: C.sub }}>{message}</p>
        <div className="flex gap-2 justify-end">
          <button onClick={onCancel} className="px-4 py-2 rounded-md text-sm" style={{ color: C.sub }}>Cancel</button>
          <button onClick={onConfirm} className="px-4 py-2 rounded-md text-sm text-white font-semibold" style={{ background: toneColor }}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

/* ----------------------------- ROW ACTIONS -------------------------------*/
function RowActions({ status, view, onEdit, onDelete, onMarkCorrect, onMarkIncorrect, onRestore, onDeletePermanent }) {
  if (view === "incorrect") {
    return (
      <div className="flex items-center gap-2 justify-end">
        <button onClick={onRestore} title="Restore to pending" className="flex items-center gap-1 px-2 py-1 rounded text-xs font-medium" style={{ color: C.navy, background: "#EAF0F8" }}>
          <RotateCcw size={13} /> Restore
        </button>
        <button onClick={onDeletePermanent} title="Delete permanently" className="p-1.5 rounded" style={{ color: C.danger }}>
          <Trash2 size={15} />
        </button>
      </div>
    );
  }
  if (status === "correct") {
    return (
      <div className="flex items-center gap-1 justify-end text-xs font-medium" style={{ color: C.success }}>
        <Lock size={13} /> Verified
      </div>
    );
  }
  return (
    <div className="flex items-center gap-1 justify-end">
      <button onClick={onEdit} title="Edit" className="p-1.5 rounded" style={{ color: C.navy }}><Pencil size={15} /></button>
      <button onClick={onDelete} title="Delete" className="p-1.5 rounded" style={{ color: C.danger }}><Trash2 size={15} /></button>
      <span className="w-px h-4 mx-1" style={{ background: C.line }} />
      <button onClick={onMarkCorrect} title="Mark correct" className="p-1.5 rounded" style={{ color: C.success }}><CheckCircle2 size={16} /></button>
      <button onClick={onMarkIncorrect} title="Mark incorrect" className="p-1.5 rounded" style={{ color: C.danger }}><XCircle size={16} /></button>
    </div>
  );
}

/* --------------------------------- ROOT ---------------------------------*/
export default function GrpMasterConsole() {
  const [states, setStates] = useState(RAW_STATES);
  const [districts, setDistricts] = useState(RAW_DISTRICTS);
  const [stations, setStations] = useState(RAW_STATIONS);
  const [links, setLinks] = useState(RAW_LINKS);

  const [activeTab, setActiveTab] = useState("state");
  const [view, setView] = useState("active"); // "active" | "incorrect"

  // Separate individual queries per table tab for clean encapsulated searching
  const [queries, setQueries] = useState({
    state: "",
    district: "",
    station: "",
    link: ""
  });

  // Unique pagination states running across tables
  const [pages, setPages] = useState({
    state: 1,
    district: 1,
    station: 1,
    link: 1
  });

  // Sort direction for the GRP–RPF Link table, keyed by number of linked RPF posts.
  // "none" = insertion order, "desc" = most posts first, "asc" = fewest posts first.
  const [linkPostSort, setLinkPostSort] = useState("none");

  const [modal, setModal] = useState(null); // { mode, data }
  const [error, setError] = useState("");
  const [confirmAction, setConfirmAction] = useState(null);
  const [toast, setToast] = useState(null);
  const [expandedLinks, setExpandedLinks] = useState(() => new Set());

  const excelInputRef = useRef(null);

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(null), 2600); };

  // ---- lookups so a rename anywhere propagates instantly ----
  const stateName = (code) => states.find((s) => s.code === code)?.name || "";
  const districtName = (stateCode, code) => districts.find((d) => d.stateCode === stateCode && d.code === code)?.name || "";
  const stationName = (stateCode, districtCode, code) => stations.find((s) => s.stateCode === stateCode && s.districtCode === districtCode && s.code === code)?.name || "";

  const nextDistrictCode = (stateCode) =>
    String(Math.max(9, ...districts.filter((d) => d.stateCode === stateCode).map((d) => Number(d.code) || 0)) + 1);
  const nextStationCode = (stateCode, districtCode) =>
    String(Math.max(9, ...stations.filter((s) => s.stateCode === stateCode && s.districtCode === districtCode).map((s) => Number(s.code) || 0)) + 1);

  const emptyData = {
    state: { code: "", name: "", strCode: "", address: "" },
    district: { stateCode: "", code: "", name: "", strCode: "", address: "", shiftType: "" },
    station: { stateCode: "", districtCode: "", code: "", name: "", strCode: "", address: "" },
    link: { stateCode: "", districtCode: "", stationCode: "", zone: "", division: "", post: "" },
  };

  const openAdd = () => { setError(""); setView("active"); setModal({ mode: "add", data: { ...emptyData[activeTab] } }); };
  const openEdit = (row) => { setError(""); setModal({ mode: "edit", data: { ...row } }); };
  const closeModal = () => setModal(null);
  const resetForm = () => setModal((m) => ({ ...m, data: { ...emptyData[activeTab] } }));

  function setField(field, value) {
    setModal((m) => {
      let data = { ...m.data, [field]: value };
      if (activeTab === "station" && field === "stateCode") data.districtCode = "";
      if (m.mode === "add") {
        if (activeTab === "district" && data.stateCode) data.code = nextDistrictCode(data.stateCode);
        if (activeTab === "station" && data.stateCode && data.districtCode) data.code = nextStationCode(data.stateCode, data.districtCode);
      }
      return { ...m, data };
    });
  }

  function handleQueryChange(val) {
    setQueries(prev => ({ ...prev, [activeTab]: val }));
    setPages(prev => ({ ...prev, [activeTab]: 1 })); // reset page to 1 on filter
  }

  function handlePageChange(newPage) {
    setPages(prev => ({ ...prev, [activeTab]: newPage }));
  }

  // Cycles the "Linked RPF Posts" sort: none -> most posts first -> fewest posts first -> none.
  function cycleLinkPostSort() {
    setLinkPostSort((prev) => (prev === "none" ? "desc" : prev === "desc" ? "asc" : "none"));
    setPages((prev) => ({ ...prev, link: 1 }));
  }

  function validate(tab, d) {
    if (tab === "state") {
      if (!d.code || !d.name || !d.strCode || !d.address) return "Please fill in all fields.";
      const codeStr = String(d.code).trim();
      const dupe = states.some((s) => String(s.code) === codeStr && !(modal.mode === "edit" && String(modal.data.code) === codeStr));
      if (dupe) return `State code ${d.code} already exists.`;
    }
    if (tab === "district") {
      if (!d.stateCode || !d.code || !d.name || !d.strCode || !d.address || !d.shiftType) return "Please fill in all required fields, including Shift Type.";
    }
    if (tab === "station") {
      if (!d.stateCode || !d.districtCode || !d.code || !d.name || !d.strCode || !d.address) return "Please fill in all required fields.";
    }
    if (tab === "link") {
      if (!d.stateCode || !d.districtCode || !d.stationCode || !d.zone || !d.division || !d.post) return "Please complete both the GRP Station and RPF Post selections.";
      const dupe = links.some((x) => x.stateCode === d.stateCode && x.districtCode === d.districtCode && x.stationCode === d.stationCode && x.zone === d.zone && x.division === d.division && x.post === d.post && x.id !== d.id);
      if (dupe) return "This GRP Station is already linked to this RPF Post.";
    }
    return "";
  }

  function handleSubmit() {
    const { mode, data } = modal;
    const err = validate(activeTab, data);
    if (err) { setError(err); return; }
    const payload = { ...data, code: data.code !== "" && data.code !== undefined ? String(data.code) : data.code };

    if (activeTab === "state") {
      setStates((prev) => mode === "add" ? [{ ...payload, status: "pending" }, ...prev] : prev.map((s) => (s.code === payload.code ? { ...s, ...payload } : s)));
    } else if (activeTab === "district") {
      setDistricts((prev) => mode === "add" ? [{ ...payload, status: "pending" }, ...prev] : prev.map((x) => (x.stateCode === payload.stateCode && x.code === payload.code ? { ...x, ...payload } : x)));
    } else if (activeTab === "station") {
      setStations((prev) => mode === "add" ? [{ ...payload, status: "pending" }, ...prev] : prev.map((x) => (x.stateCode === payload.stateCode && x.districtCode === payload.districtCode && x.code === payload.code ? { ...x, ...payload } : x)));
    } else if (activeTab === "link") {
      setLinks((prev) => mode === "add" ? [{ ...payload, id: Math.max(0, ...prev.map((l) => l.id)) + 1, status: "pending" }, ...prev] : prev.map((x) => (x.id === payload.id ? { ...x, ...payload } : x)));
    }
    showToast(mode === "add" ? "Entry added — pending verification." : "Entry updated.");
    closeModal();
  }

  // ---- verification workflow ----
  function markStatus(tab, matchRow, status) {
    const upd = (list, matcher) => list.map((r) => (matcher(r) ? { ...r, status } : r));
    if (tab === "state") setStates((p) => upd(p, (r) => r.code === matchRow.code));
    if (tab === "district") setDistricts((p) => upd(p, (r) => r.stateCode === matchRow.stateCode && r.code === matchRow.code));
    if (tab === "station") setStations((p) => upd(p, (r) => r.stateCode === matchRow.stateCode && r.districtCode === matchRow.districtCode && r.code === matchRow.code));
    if (tab === "link") setLinks((p) => upd(p, (r) => r.id === matchRow.id));
    showToast(status === "correct" ? "Marked correct — record is now locked." : status === "incorrect" ? "Moved to Incorrect." : "Restored to pending.");
  }

  function dependents(tab, row) {
    if (tab === "state") return districts.some((d) => d.stateCode === row.code) ? "This state has districts under it. Remove those first." : "";
    if (tab === "district") return stations.some((s) => s.stateCode === row.stateCode && s.districtCode === row.code) ? "This district has GRP stations under it. Remove those first." : "";
    if (tab === "station") return links.some((l) => l.stateCode === row.stateCode && l.districtCode === row.districtCode && l.stationCode === row.code) ? "This GRP station is linked to an RPF post. Remove the link first." : "";
    return "";
  }

  function performDelete(tab, row) {
    if (tab === "state") setStates((p) => p.filter((s) => s.code !== row.code));
    if (tab === "district") setDistricts((p) => p.filter((x) => !(x.stateCode === row.stateCode && x.code === row.code)));
    if (tab === "station") setStations((p) => p.filter((x) => !(x.stateCode === row.stateCode && x.districtCode === row.districtCode && x.code === row.code)));
    if (tab === "link") setLinks((p) => p.filter((x) => x.id !== row.id));
    showToast("Entry deleted.");
    setConfirmAction(null);
  }

  function askDelete(tab, row, label) {
    const blocked = dependents(tab, row);
    if (blocked) { setConfirmAction({ title: "Can't delete yet", message: blocked, confirmLabel: "Got it", onConfirm: () => setConfirmAction(null), tone: "navy" }); return; }
    setConfirmAction({ title: "Delete entry?", message: `This removes ${label} and cannot be undone.`, confirmLabel: "Delete", tone: "danger", onConfirm: () => performDelete(tab, row) });
  }

  // ---- Excel import ----
  function triggerExcelImport() {
    if (excelInputRef.current) excelInputRef.current.click();
  }

  async function handleExcelFile(e) {
    const file = e.target.files?.[0];
    const tab = activeTab;
    e.target.value = ""; // allow re-selecting the same file later
    if (!file) return;
    try {
      const rows = await readExcelFile(file);
      // Row 1: "Railway Security Response System" title, Row 2: column headers, data from Row 3.
      const dataRows = rows.slice(2);
      importForTab(tab, dataRows);
    } catch (err) {
      showToast("Couldn't read that file — please upload a valid Excel (.xlsx / .xls) export.");
    }
  }

  function importForTab(tab, dataRows) {
    let added = 0, skipped = 0;

    if (tab === "state") {
      const newOnes = [];
      dataRows.forEach((cols) => {
        const code = String(cols[0] ?? "").trim();
        const name = String(cols[1] ?? "").trim();
        const strCode = String(cols[2] ?? "").trim();
        const address = String(cols[3] ?? "").trim();
        if (!code || !name) { skipped++; return; }
        const dup = states.some((s) => s.code === code) || newOnes.some((s) => s.code === code);
        if (dup) { skipped++; return; }
        newOnes.push({ code, name, strCode, address: address || name, status: "pending" });
        added++;
      });
      if (newOnes.length) setStates((prev) => [...newOnes, ...prev]);
    }

    else if (tab === "district") {
      const newOnes = [];
      dataRows.forEach((cols) => {
        const sName = String(cols[0] ?? "").trim();
        const code = String(cols[1] ?? "").trim();
        const name = String(cols[2] ?? "").trim();
        const strCode = String(cols[3] ?? "").trim();
        const address = String(cols[4] ?? "").trim();
        const st = states.find((s) => normalizeName(s.name) === normalizeName(sName));
        if (!st || !code || !name) { skipped++; return; }
        const dup = districts.some((d) => d.stateCode === st.code && d.code === code) || newOnes.some((d) => d.stateCode === st.code && d.code === code);
        if (dup) { skipped++; return; }
        newOnes.push({ stateCode: st.code, code, name, strCode, address: address || name, shiftType: "", status: "pending" });
        added++;
      });
      if (newOnes.length) setDistricts((prev) => [...newOnes, ...prev]);
    }

    else if (tab === "station") {
      const newOnes = [];
      dataRows.forEach((cols) => {
        const sName = String(cols[0] ?? "").trim();
        const dName = String(cols[1] ?? "").trim();
        const code = String(cols[2] ?? "").trim();
        const name = String(cols[3] ?? "").trim();
        const strCode = String(cols[4] ?? "").trim();
        const address = String(cols[5] ?? "").trim();
        const st = states.find((s) => normalizeName(s.name) === normalizeName(sName));
        const dist = st && districts.find((d) => d.stateCode === st.code && normalizeName(d.name) === normalizeName(dName));
        if (!st || !dist || !code || !name) { skipped++; return; }
        const dup = stations.some((s) => s.stateCode === st.code && s.districtCode === dist.code && s.code === code)
          || newOnes.some((s) => s.stateCode === st.code && s.districtCode === dist.code && s.code === code);
        if (dup) { skipped++; return; }
        newOnes.push({ stateCode: st.code, districtCode: dist.code, code, name, strCode, address: address || name, status: "pending" });
        added++;
      });
      if (newOnes.length) setStations((prev) => [...newOnes, ...prev]);
    }

    else if (tab === "link") {
      const newOnes = [];
      let nextId = Math.max(0, ...links.map((l) => l.id), ...newOnes.map((l) => l.id || 0));
      dataRows.forEach((cols) => {
        const grpUnit = String(cols[1] ?? "").trim();
        const rpfUnit = String(cols[2] ?? "").trim();
        const grpParts = grpUnit.split("/").map((p) => p.trim());
        const rpfParts = rpfUnit.split("/").map((p) => p.trim());
        if (grpParts.length < 3 || rpfParts.length < 3) { skipped++; return; }
        const [sName, dName, stName] = grpParts;
        const [zone, division, post] = rpfParts;
        const st = states.find((s) => normalizeName(s.name) === normalizeName(sName));
        const dist = st && districts.find((d) => d.stateCode === st.code && normalizeName(d.name) === normalizeName(dName));
        const stn = dist && stations.find((s) => s.stateCode === st.code && s.districtCode === dist.code && normalizeName(s.name) === normalizeName(stName));
        if (!st || !dist || !stn || !zone || !division || !post) { skipped++; return; }
        const dup = links.some((l) => l.stateCode === st.code && l.districtCode === dist.code && l.stationCode === stn.code && l.zone === zone && l.division === division && l.post === post)
          || newOnes.some((l) => l.stateCode === st.code && l.districtCode === dist.code && l.stationCode === stn.code && l.zone === zone && l.division === division && l.post === post);
        if (dup) { skipped++; return; }
        nextId += 1;
        newOnes.push({ id: nextId, stateCode: st.code, districtCode: dist.code, stationCode: stn.code, zone, division, post, status: "pending" });
        added++;
      });
      if (newOnes.length) setLinks((prev) => [...newOnes, ...prev]);
    }

    showToast(`Imported ${added} ${added === 1 ? "entry" : "entries"}${skipped ? `, skipped ${skipped} (duplicate or unmatched)` : ""}.`);
  }

  function toggleExpandLink(key) {
    setExpandedLinks((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  }

  // Quick-add: prefill the GRP Station side from an existing group so the
  // person only has to pick Zone / Division / Post for the new RPF link.
  function openAddRpfPostForGroup(first) {
    setError("");
    setExpandedLinks((prev) => new Set(prev).add(`${first.stateCode}::${first.districtCode}::${first.stationCode}`));
    setModal({
      mode: "add",
      data: { stateCode: first.stateCode, districtCode: first.districtCode, stationCode: first.stationCode, zone: "", division: "", post: "" },
    });
  }

  const rowsForTab = { state: states, district: districts, station: stations, link: links }[activeTab];
  const activeCount = rowsForTab.filter((r) => r.status !== "incorrect").length;
  const incorrectCount = rowsForTab.filter((r) => r.status === "incorrect").length;

  const filtered = useMemo(() => {
    const q = queries[activeTab].trim().toLowerCase();
    const byView = rowsForTab.filter((r) => (view === "incorrect" ? r.status === "incorrect" : r.status !== "incorrect"));
    if (!q) return byView;
    return byView.filter((r) => Object.values(r).some((v) => String(v).toLowerCase().includes(q)));
  }, [rowsForTab, view, queries, activeTab]);

  // Paginated visible standard rows (state, district, station)
  const paginatedRows = useMemo(() => {
    const startIdx = (pages[activeTab] - 1) * ITEMS_PER_PAGE;
    return filtered.slice(startIdx, startIdx + ITEMS_PER_PAGE);
  }, [filtered, pages, activeTab]);

  // Links grouped by their GRP Unit (state/district/station), so a station that's
  // mapped to several RPF posts shows as one expandable row instead of N flat rows.
  // Optionally sorted by number of linked RPF posts via linkPostSort.
  const groupedLinks = useMemo(() => {
    const q = queries.link.trim().toLowerCase();
    const matches = links.filter((r) => {
      const inView = view === "incorrect" ? r.status === "incorrect" : r.status !== "incorrect";
      if (!inView) return false;
      if (!q) return true;
      return Object.values(r).some((v) => String(v).toLowerCase().includes(q));
    });
    const map = new Map();
    matches.forEach((r) => {
      const key = `${r.stateCode}::${r.districtCode}::${r.stationCode}`;
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(r);
    });
    let groups = Array.from(map.entries()).map(([key, entries]) => ({ key, entries }));
    if (linkPostSort === "desc") groups = groups.sort((a, b) => b.entries.length - a.entries.length);
    else if (linkPostSort === "asc") groups = groups.sort((a, b) => a.entries.length - b.entries.length);
    return groups;
  }, [links, view, queries.link, linkPostSort]);

  // Paginated grouped links array
  const paginatedGroupedLinks = useMemo(() => {
    if (activeTab !== "link") return [];
    const startIdx = (pages.link - 1) * ITEMS_PER_PAGE;
    return groupedLinks.slice(startIdx, startIdx + ITEMS_PER_PAGE);
  }, [groupedLinks, pages.link, activeTab]);

  const counts = { state: states.length, district: districts.length, station: stations.length, link: links.length };
  const isEmpty = activeTab === "link" ? groupedLinks.length === 0 : filtered.length === 0;

  const LinkSortIcon = linkPostSort === "desc" ? ArrowDown : linkPostSort === "asc" ? ArrowUp : ArrowUpDown;

  return (
    <div className="min-h-screen w-full" style={{ background: C.bg, fontFamily: "Inter, sans-serif", color: C.ink }}>
      <link rel="stylesheet" href={FONT_LINK} />

      {/* Header */}
      <div style={{ background: `linear-gradient(180deg, ${C.navy}, ${C.navyDeep})` }} className="px-5 md:px-8 pt-6 pb-5">
        <div className="max-w-6xl mx-auto">
          <div className="flex items-center gap-2 text-xs mb-1" style={{ color: "#9FB3CE", fontFamily: "IBM Plex Mono, monospace" }}>
            <span>GRP MASTER DATA</span><ChevronRight size={12} /><span>CONSOLE</span>
          </div>
          <h1 className="text-white text-2xl md:text-3xl font-bold" style={{ fontFamily: "Barlow Semi Condensed, sans-serif", letterSpacing: 0.4 }}>
            State &rarr; District &rarr; GRP Station &rarr; RPF Link
          </h1>

          <div className="flex items-center mt-6 overflow-x-auto pb-1">
            {TABS.map((t, i) => (
              <React.Fragment key={t.id}>
                {i > 0 && <div className="h-px w-8 md:w-14 shrink-0" style={{ background: "#3C557C" }} />}
                <button onClick={() => { setActiveTab(t.id); setView("active"); }} className="flex items-center gap-2 shrink-0 px-1 py-1">
                  <span className="w-3.5 h-3.5 rounded-full shrink-0" style={{ background: activeTab === t.id ? C.amber : "transparent", border: `2px solid ${activeTab === t.id ? C.amber : "#5B739A"}` }} />
                  <span className="text-sm md:text-base whitespace-nowrap" style={{ fontFamily: "Barlow Semi Condensed, sans-serif", fontProject: activeTab === t.id ? 700 : 600, color: activeTab === t.id ? "#FFFFFF" : "#93A7C4" }}>{t.label}</span>
                  <span className="text-xs px-1.5 rounded" style={{ background: "rgba(255,255,255,0.1)", color: "#C7D3E3", fontFamily: "IBM Plex Mono, monospace" }}>{counts[t.id]}</span>
                </button>
              </React.Fragment>
            ))}
          </div>
        </div>
      </div>

      {/* Toolbar */}
      <div className="max-w-6xl mx-auto px-5 md:px-8 mt-5 flex flex-col gap-3">
        <div className="flex flex-col sm:flex-row gap-3 sm:items-center sm:justify-between">
          <div className="relative w-full sm:w-72">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: C.sub }} />
            <input value={queries[activeTab]} onChange={(e) => handleQueryChange(e.target.value)} placeholder={`Search ${TABS.find((t) => t.id === activeTab).plural.toLowerCase()}…`} className="w-full rounded-md border pl-9 pr-3 py-2 text-sm focus:outline-none" style={{ borderColor: C.line, background: C.card }} />
          </div>
          <div className="flex gap-2 shrink-0">
            <input
              ref={excelInputRef}
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              className="hidden"
              onChange={handleExcelFile}
            />
            <button
              onClick={triggerExcelImport}
              title={`Expected columns: ${EXCEL_COLUMNS[activeTab].join(", ")} (row 1 = report title, row 2 = headers, data from row 3). Accepts .xlsx or .xls.`}
              className="flex items-center justify-center gap-1.5 px-4 py-2 rounded-md text-sm font-semibold"
              style={{ background: C.card, color: C.navy, border: `1px solid ${C.line}` }}
            >
              <Upload size={16} /> Import Excel
            </button>
            <button onClick={openAdd} className="flex items-center justify-center gap-1.5 px-4 py-2 rounded-md text-white text-sm font-semibold" style={{ background: C.amberDeep }}>
              <Plus size={16} /> Add New {TABS.find((t) => t.id === activeTab).label}
            </button>
          </div>
        </div>

        <div className="flex items-center justify-between flex-wrap gap-2">
          <div className="flex gap-2">
            <button onClick={() => { setView("active"); setPages(p => ({ ...p, [activeTab]: 1 })); }} className="px-3 py-1.5 rounded-full text-xs font-semibold" style={{ background: view === "active" ? C.navy : C.card, color: view === "active" ? "#fff" : C.sub, border: `1px solid ${view === "active" ? C.navy : C.line}` }}>
              Active ({activeCount})
            </button>
            <button onClick={() => { setView("incorrect"); setPages(p => ({ ...p, [activeTab]: 1 })); }} className="px-3 py-1.5 rounded-full text-xs font-semibold flex items-center gap-1" style={{ background: view === "incorrect" ? C.danger : C.card, color: view === "incorrect" ? "#fff" : C.sub, border: `1px solid ${view === "incorrect" ? C.danger : C.line}` }}>
              <XCircle size={13} /> Incorrect ({incorrectCount})
            </button>
          </div>
          <span className="text-xs" style={{ color: C.sub, fontFamily: "IBM Plex Mono, monospace" }}>
            Excel columns: {EXCEL_COLUMNS[activeTab].join(" · ")}
          </span>
        </div>
      </div>

      {/* Table Content */}
      <div className="max-w-6xl mx-auto px-5 md:px-8 py-5">
        <div className="rounded-lg overflow-hidden" style={{ background: C.card, border: `1px solid ${C.line}` }}>
          {activeTab === "state" && (
            <>
              <Table cols={["State Code", "State Name", "Str Code", "Location Address", "Verification"]} rows={paginatedRows} renderRow={(r) => [
                <Mono>{r.code}</Mono>, r.name, <Mono>{r.strCode}</Mono>, r.address,
                <RowActions status={r.status} view={view}
                  onEdit={() => openEdit(r)} onDelete={() => askDelete("state", r, r.name)}
                  onMarkCorrect={() => markStatus("state", r, "correct")} onMarkIncorrect={() => markStatus("state", r, "incorrect")}
                  onRestore={() => markStatus("state", r, "pending")} onDeletePermanent={() => askDelete("state", r, r.name)} />,
              ]} />
              <TablePagination currentPage={pages.state} totalItems={filtered.length} onPageChange={(p) => handlePageChange(p)} />
            </>
          )}
          {activeTab === "district" && (
            <>
              <Table cols={["State Name", "District Code", "District Name", "Str Code", "Shift", "Location Address", "Verification"]} rows={paginatedRows} renderRow={(r) => [
                stateName(r.stateCode), <Mono>{r.code}</Mono>, r.name, <Mono>{r.strCode}</Mono>, r.shiftType || "—", r.address,
                <RowActions status={r.status} view={view}
                  onEdit={() => openEdit(r)} onDelete={() => askDelete("district", r, r.name)}
                  onMarkCorrect={() => markStatus("district", r, "correct")} onMarkIncorrect={() => markStatus("district", r, "incorrect")}
                  onRestore={() => markStatus("district", r, "pending")} onDeletePermanent={() => askDelete("district", r, r.name)} />,
              ]} />
              <TablePagination currentPage={pages.district} totalItems={filtered.length} onPageChange={(p) => handlePageChange(p)} />
            </>
          )}
          {activeTab === "station" && (
            <>
              <Table cols={["State Name", "District Name", "Station Code", "Station Name", "Str Code", "Location Address", "Verification"]} rows={paginatedRows} renderRow={(r) => [
                stateName(r.stateCode), districtName(r.stateCode, r.districtCode), <Mono>{r.code}</Mono>, r.name, <Mono>{r.strCode}</Mono>, r.address,
                <RowActions status={r.status} view={view}
                  onEdit={() => openEdit(r)} onDelete={() => askDelete("station", r, r.name)}
                  onMarkCorrect={() => markStatus("station", r, "correct")} onMarkIncorrect={() => markStatus("station", r, "incorrect")}
                  onRestore={() => markStatus("station", r, "pending")} onDeletePermanent={() => askDelete("station", r, r.name)} />,
              ]} />
              <TablePagination currentPage={pages.station} totalItems={filtered.length} onPageChange={(p) => handlePageChange(p)} />
            </>
          )}
          {activeTab === "link" && (
            <>
              <div className="overflow-x-auto">
                <table className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: C.bg }}>
                      <th className="text-left px-4 py-3 font-semibold whitespace-nowrap" style={{ color: C.sub, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, borderBottom: `1px solid ${C.line}` }}>GRP Unit</th>
                      <th className="text-right px-4 py-3 font-semibold whitespace-nowrap" style={{ color: C.sub, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, borderBottom: `1px solid ${C.line}` }}>
                        <button
                          onClick={cycleLinkPostSort}
                          title={linkPostSort === "none" ? "Sort by number of posts" : linkPostSort === "desc" ? "Sorted: most posts first" : "Sorted: fewest posts first"}
                          className="inline-flex items-center gap-1 uppercase"
                          style={{ color: linkPostSort === "none" ? C.sub : C.navy, fontSize: 11, letterSpacing: 0.5, fontWeight: 600 }}
                        >
                          Linked RPF Posts <LinkSortIcon size={12} />
                        </button>
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {paginatedGroupedLinks.map(({ key, entries }) => {
                      const first = entries[0];
                      const expanded = expandedLinks.has(key);
                      const label = `${stateName(first.stateCode)} / ${districtName(first.stateCode, first.districtCode)} / ${stationName(first.stateCode, first.districtCode, first.stationCode)}`;
                      const correctCt = entries.filter((e) => e.status === "correct").length;
                      return (
                        <React.Fragment key={key}>
                          <tr
                            onClick={() => toggleExpandLink(key)}
                            className="cursor-pointer"
                            style={{ borderBottom: expanded ? "none" : `1px solid ${C.line}`, background: expanded ? C.bg : "transparent" }}
                          >
                            <td className="px-4 py-3 align-top">
                              <div className="flex items-center gap-2">
                                <ChevronRight size={14} style={{ transform: expanded ? "rotate(90deg)" : "none", transition: "transform 0.15s ease", color: C.sub, flexShrink: 0 }} />
                                <span style={{ fontWeight: 600 }}>{label}</span>
                              </div>
                            </td>
                            <td className="px-4 py-3 align-top text-right">
                              <span className="text-xs px-2.5 py-1 rounded-full" style={{ background: "#EAF0F8", color: C.navy, fontFamily: "IBM Plex Mono, monospace" }}>
                                {entries.length} post{entries.length !== 1 ? "s" : ""}{view === "active" ? ` · ${correctCt} verified` : ""}
                              </span>
                            </td>
                          </tr>
                          {expanded && (
                            <tr style={{ borderBottom: `1px solid ${C.line}` }}>
                              <td colSpan={2} className="px-0 py-0">
                                <table className="w-full text-sm" style={{ borderCollapse: "collapse", background: "#FAF9F5" }}>
                                  <tbody>
                                    {entries.map((r) => (
                                      <tr key={r.id} style={{ borderTop: `1px solid ${C.line}` }}>
                                        <td className="py-2.5 align-top" style={{ paddingLeft: 40, paddingRight: 16, width: "60%" }}>
                                          <span style={{ color: C.sub }}>{r.zone} / {r.division} / </span><b>{r.post}</b>
                                        </td>
                                        <td className="px-4 py-2.5 align-top text-right">
                                          <RowActions status={r.status} view={view}
                                            onEdit={() => openEdit(r)}
                                            onDelete={() => askDelete("link", r, `${stationName(r.stateCode, r.districtCode, r.stationCode)} ↔ ${r.post}`)}
                                            onMarkCorrect={() => markStatus("link", r, "correct")}
                                            onMarkIncorrect={() => markStatus("link", r, "incorrect")}
                                            onRestore={() => markStatus("link", r, "pending")}
                                            onDeletePermanent={() => askDelete("link", r, `${stationName(r.stateCode, r.districtCode, r.stationCode)} ↔ ${r.post}`)} />
                                        </td>
                                      </tr>
                                    ))}
                                    <tr style={{ borderTop: `1px solid ${C.line}` }}>
                                      <td colSpan={2} className="py-2 px-4" style={{ paddingLeft: 40 }}>
                                        <button
                                          onClick={() => openAddRpfPostForGroup(first)}
                                          className="flex items-center gap-1 text-xs font-semibold"
                                          style={{ color: C.amberDeep }}
                                        >
                                          <Plus size={13} /> Link another RPF Post here
                                        </button>
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <TablePagination currentPage={pages.link} totalItems={groupedLinks.length} onPageChange={(p) => handlePageChange(p)} />
            </>
          )}
          {isEmpty && (
            <div className="py-14 text-center">
              <p className="text-sm mb-3" style={{ color: C.sub }}>
                {queries[activeTab] ? "No entries match your search." : view === "incorrect" ? "Nothing flagged incorrect here." : `No ${TABS.find((t) => t.id === activeTab).plural.toLowerCase()} yet.`}
              </p>
              {!queries[activeTab] && view === "active" && (
                <button onClick={openAdd} className="text-sm font-semibold underline" style={{ color: C.amberDeep }}>Add the first one</button>
              )}
            </div>
          )}
        </div>
      </div>

      {modal && (
        <Modal tab={activeTab} mode={modal.mode} data={modal.data} set={setField} onClose={closeModal} onSubmit={handleSubmit} onReset={resetForm} states={states} districts={districts} stations={stations} error={error} />
      )}
      {confirmAction && (
        <ConfirmDialog title={confirmAction.title} message={confirmAction.message} confirmLabel={confirmAction.confirmLabel} tone={confirmAction.tone} onCancel={() => setConfirmAction(null)} onConfirm={confirmAction.onConfirm} />
      )}
      {toast && (
        <div className="fixed bottom-5 left-1/2 -translate-x-1/2 z-50 flex items-center gap-2 px-4 py-2.5 rounded-md text-sm text-white shadow-lg" style={{ background: C.navy }}>
          <Check size={16} style={{ color: C.amber }} /> {toast}
        </div>
      )}
    </div>
  );
}

function Mono({ children }) {
  return <span style={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 13, color: C.navy }}>{children}</span>;
}

function Table({ cols, rows, renderRow }) {
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
        <thead>
          <tr style={{ background: C.bg }}>
            {cols.map((c, i) => (
              <th key={i} className={`text-left px-4 py-3 font-semibold whitespace-nowrap ${i === cols.length - 1 ? "text-right" : ""}`} style={{ color: C.sub, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, borderBottom: `1px solid ${C.line}` }}>{c}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => {
            const cells = renderRow(r, i);
            return (
              <tr key={i} style={{ borderBottom: `1px solid ${C.line}` }} className="hover:bg-black/[0.02]">
                {cells.map((cell, j) => (
                  <td key={j} className={`px-4 py-3 align-top ${j === cells.length - 1 ? "text-right" : ""}`} style={{ color: C.ink }}>{cell}</td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
