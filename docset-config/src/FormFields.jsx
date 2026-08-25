import { useEffect, useRef, useState } from "react";
import { fetchChoices, searchLookup } from "./xrm";

/** Renders one field from a schema descriptor. */
export function Field({ table, field, value, onChange, disabled }) {
  const { kind, label, required, derived } = field;
  const readOnly = disabled || derived;

  return (
    <label className={"ff" + (kind === "memo" ? " ff-wide" : "")}>
      <span className="ff-label">
        {label}
        {required && <i className="ff-req">*</i>}
        {derived && <em className="ff-hint">filled in on save</em>}
      </span>
      <Control
        table={table}
        field={field}
        value={value}
        onChange={onChange}
        readOnly={readOnly}
      />
    </label>
  );
}

function Control({ table, field, value, onChange, readOnly }) {
  switch (field.kind) {
    case "memo":
      return (
        <textarea className="ff-input" rows={3} value={value ?? ""}
          disabled={readOnly} onChange={(e) => onChange(e.target.value)} />
      );

    case "bool":
      return (
        <select className="ff-input"
          value={value === true ? "true" : value === false ? "false" : ""}
          disabled={readOnly}
          onChange={(e) => onChange(e.target.value === "" ? null : e.target.value === "true")}>
          <option value="">—</option>
          <option value="false">No</option>
          <option value="true">Yes</option>
        </select>
      );

    case "choice":
      return <ChoiceControl table={table} column={field.name} value={value}
        onChange={onChange} readOnly={readOnly} />;

    case "lookup":
      return <LookupControl target={field.lookupTo} value={value}
        onChange={onChange} readOnly={readOnly} />;

    default:
      return (
        <input className="ff-input" type="text" value={value ?? ""}
          disabled={readOnly} onChange={(e) => onChange(e.target.value)} />
      );
  }
}

/**
 * Options come from column metadata at runtime.
 *
 * Hard-coding them would break the moment a client's option values differ from
 * ours — and Dataverse assigns 100000000-style values by default, so they very
 * often do.
 */
function ChoiceControl({ table, column, value, onChange, readOnly }) {
  const [options, setOptions] = useState([]);

  useEffect(() => {
    let alive = true;
    fetchChoices(table, column).then((o) => { if (alive) setOptions(o); });
    return () => { alive = false; };
  }, [table, column]);

  return (
    <select className="ff-input" value={value ?? ""} disabled={readOnly}
      onChange={(e) => onChange(e.target.value === "" ? null : Number(e.target.value))}>
      <option value="">—</option>
      {options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
    </select>
  );
}

/** Type-ahead lookup. Value is `{ id, name }` or null. */
function LookupControl({ target, value, onChange, readOnly }) {
  const [term, setTerm] = useState("");
  const [results, setResults] = useState([]);
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const boxRef = useRef(null);

  useEffect(() => {
    const onDocClick = (e) => {
      if (boxRef.current && !boxRef.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  useEffect(() => {
    if (!open) return;
    let alive = true;
    setBusy(true);
    // Debounced, so typing does not fire a query per keystroke.
    const t = setTimeout(async () => {
      try {
        const rows = await searchLookup(target, term);
        if (alive) setResults(rows);
      } catch (e) {
        if (alive) setResults([]);
        console.warn("Lookup search failed:", e);
      } finally {
        if (alive) setBusy(false);
      }
    }, 250);
    return () => { alive = false; clearTimeout(t); };
  }, [term, open, target]);

  if (readOnly) {
    return <input className="ff-input" type="text" value={value?.name ?? ""} disabled />;
  }

  if (value) {
    return (
      <div className="ff-lookup-set">
        <span className="ff-token">{value.name}</span>
        <button type="button" className="ff-clear"
          onClick={() => { onChange(null); setTerm(""); }} aria-label="Clear">✕</button>
      </div>
    );
  }

  return (
    <div className="ff-lookup" ref={boxRef}>
      <input className="ff-input" type="text" value={term} placeholder="Search…"
        onFocus={() => setOpen(true)}
        onChange={(e) => { setTerm(e.target.value); setOpen(true); }} />
      {open && (
        <div className="ff-drop">
          {busy && <div className="ff-drop-empty">Searching…</div>}
          {!busy && results.length === 0 && <div className="ff-drop-empty">No matches</div>}
          {!busy && results.map((r) => (
            <button type="button" key={r.id} className="ff-drop-item"
              onClick={() => { onChange(r); setOpen(false); setTerm(""); }}>
              {r.name}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
