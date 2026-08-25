import { useCallback, useEffect, useState } from "react";
import {
  TABLES, KEYS, LIBRARY_FIELDS, MAPPING_FIELDS, FIELD_MAPPING_FIELDS,
} from "./schema";
import {
  fetchList, retrieveRecord, createRecord, updateRecord, deleteRecord,
  entitySetName, isMock,
  registerProvisioningStep, deactivateProvisioningStep, provisioningStepState,
} from "./xrm";
import { fetchEntities, tableColumns, mappableColumns } from "./metadata";
import { Field } from "./FormFields";
import { toFormState, toPayload } from "./configState";

export default function App() {
  return (
    <div className="page">
      {isMock() && (
        <div className="mock-banner">
          Preview mode — Xrm not detected. Inside the model-driven app this reads and
          writes live Dataverse configuration.
        </div>
      )}

      <div className="header">
        <div>
          <h1>Document Set Configuration</h1>
          <p>
            Which Dataverse tables file their documents into which SharePoint libraries,
            and what metadata each document set carries.
          </p>
        </div>
      </div>

      <ConfigAdmin />
    </div>
  );
}

/**
 * One configuration surface for every table.
 *
 * The flow is deliberately: pick a table, see what already exists, then decide
 * explicitly whether to use it or start a new one. "Create" and "update" look
 * identical until you press save and one of them has replaced a working
 * configuration, so the choice is a stop rather than a default.
 */
function ConfigAdmin() {
  const [entities, setEntities] = useState([]);
  const [entity, setEntity] = useState("");
  const [existing, setExisting] = useState([]);
  const [mode, setMode] = useState(null);            // null | "existing" | "new"
  const [mappingId, setMappingId] = useState(null);

  const [mapping, setMapping] = useState(null);
  const [library, setLibrary] = useState(null);
  const [libraryId, setLibraryId] = useState(null);
  const [fieldMaps, setFieldMaps] = useState([]);
  const [columns, setColumns] = useState([]);        // mappable columns on the chosen table
  const [stepActive, setStepActive] = useState(null); // null = unknown

  const [mappingCols, setMappingCols] = useState(null);
  const [libraryCols, setLibraryCols] = useState(null);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");

  useEffect(() => {
    fetchEntities().then(setEntities).catch((e) => setError(e.message || String(e)));
    Promise.all([tableColumns(TABLES.mapping), tableColumns(TABLES.library)])
      .then(([m, l]) => { setMappingCols(m); setLibraryCols(l); })
      .catch(() => { setMappingCols(null); setLibraryCols(null); });
  }, []);

  const visible = (fields, cols) => (!cols ? fields : fields.filter((f) => cols.has(f.name)));

  const loadExisting = useCallback(async (logicalName) => {
    setLoading(true); setError(""); setNotice("");
    setMode(null); setMappingId(null);
    setMapping(null); setLibrary(null); setLibraryId(null); setFieldMaps([]);
    try {
      const rows = await fetchList(
        TABLES.mapping,
        `?$select=${KEYS[TABLES.mapping].id},bw_name,bw_namingpattern,statecode,` +
        `_bw_docsetlibraryid_value` +
        `&$filter=bw_entitylogicalname eq '${logicalName}'`
      );
      setExisting(rows.map((r) => ({
        id: r[KEYS[TABLES.mapping].id],
        name: r.bw_name || "(unnamed)",
        pattern: r.bw_namingpattern || "",
        active: r.statecode === 0,
        libraryName:
          r["_bw_docsetlibraryid_value@OData.Community.Display.V1.FormattedValue"] || "",
      })));

      setStepActive(await provisioningStepState(logicalName));
      setColumns(await mappableColumns(logicalName));
    } catch (e) {
      setError(e.message || String(e));
      setExisting([]);
    } finally {
      setLoading(false);
    }
  }, []);

  const onEntityChange = (logicalName) => {
    setEntity(logicalName);
    if (logicalName) loadExisting(logicalName);
    else { setExisting([]); setStepActive(null); setColumns([]); }
  };

  const useExisting = async (id) => {
    setLoading(true); setError(""); setNotice("");
    try {
      const m = await retrieveRecord(TABLES.mapping, id,
        MAPPING_FIELDS.map((f) => (f.kind === "lookup" ? `_${f.name}_value` : f.name)));
      setMappingId(id);
      setMapping(toFormState(MAPPING_FIELDS, m));

      const libId = m["_bw_docsetlibraryid_value"];
      if (libId) {
        const lib = await retrieveRecord(TABLES.library, libId,
          LIBRARY_FIELDS.map((f) => f.name));
        setLibraryId(libId);
        setLibrary(toFormState(LIBRARY_FIELDS, lib));
      }

      const fms = await fetchList(
        TABLES.fieldMapping,
        `?$select=${FIELD_MAPPING_FIELDS.map((f) => f.name).join(",")},` +
        `${KEYS[TABLES.fieldMapping].id}` +
        `&$filter=_bw_docsetmappingid_value eq ${id} and statecode eq 0`
      );
      setFieldMaps(fms.map((r) => ({
        id: r[KEYS[TABLES.fieldMapping].id],
        ...toFormState(FIELD_MAPPING_FIELDS, r),
      })));

      setMode("existing");
    } catch (e) {
      setError(e.message || String(e));
    } finally {
      setLoading(false);
    }
  };

  /**
   * Starts a blank configuration.
   *
   * mappingId stays null, which is what makes the save create rather than update.
   * Nothing is read from the existing configuration, so nothing can leak into it.
   */
  const startNew = () => {
    const label = entities.find((e) => e.logicalName === entity)?.label || entity;
    setMappingId(null);
    setMapping({
      bw_name: `${label} documents`,
      bw_entitylogicalname: entity,
      bw_entitydisplayname: label,
      bw_namingpattern: "{primaryname}",
      bw_updatesyncenabled: false,
    });
    setLibrary(null);
    setLibraryId(null);
    setFieldMaps([]);
    setMode("new");
    setNotice("");
  };

  const saveLibrary = async () => {
    setError(""); setNotice("");
    try {
      const payload = await toPayload(visible(LIBRARY_FIELDS, libraryCols), library);
      if (libraryId) {
        await updateRecord(TABLES.library, libraryId, payload);
        const lib = await retrieveRecord(TABLES.library, libraryId,
          LIBRARY_FIELDS.map((f) => f.name));
        setLibrary(toFormState(LIBRARY_FIELDS, lib));
        setNotice("Library saved. Site and library settings resolved from SharePoint.");
      } else {
        const id = await createRecord(TABLES.library, payload);
        setLibraryId(id);
        setMapping({
          ...(mapping || {}),
          bw_docsetlibraryid: { id, name: payload.bw_documentsetlibraryname || "Library" },
        });
        setNotice("Library created and linked to this configuration.");
      }
    } catch (e) {
      setError(e.message || String(e));
    }
  };

  const saveMapping = async () => {
    setError(""); setNotice("");
    try {
      const payload = await toPayload(visible(MAPPING_FIELDS, mappingCols), mapping);
      payload.bw_entitylogicalname = entity;

      if (mode === "existing" && mappingId) {
        await updateRecord(TABLES.mapping, mappingId, payload);
        setNotice("Configuration updated.");
      } else {
        const id = await createRecord(TABLES.mapping, payload);
        setMappingId(id);
        setMode("existing");
        setNotice("New configuration created. Any previous one is unchanged.");
      }
    } catch (e) {
      setError(e.message || String(e));
    }
  };

  /**
   * Makes one configuration the only active one, and registers the plug-in step.
   *
   * Provisioning refuses to run when two are active, because choosing between them
   * would mean filing documents into a library nobody picked. Deactivating rather
   * than deleting keeps the old configuration intact and reversible.
   */
  const makeActive = async (id) => {
    setError(""); setNotice("");
    try {
      for (const c of existing) {
        if (c.id === id && !c.active) {
          await updateRecord(TABLES.mapping, c.id, { statecode: 0, statuscode: 1 });
        } else if (c.id !== id && c.active) {
          await updateRecord(TABLES.mapping, c.id, { statecode: 1, statuscode: 2 });
        }
      }

      let stepMessage = "";
      try {
        const res = await registerProvisioningStep(entity);
        stepMessage = " " + res.message;
      } catch (e) {
        // Configuration is still valid without the step - documents filed through
        // an upload provision on demand. Only eager provisioning is lost, so this
        // is reported rather than treated as a failed activation.
        stepMessage =
          " The configuration is active, but the provisioning step could not be " +
          `registered: ${e.message} Document sets will still be created when a ` +
          "document is filed, just not in advance.";
      }

      setNotice("That configuration is now the active one." + stepMessage);
      await loadExisting(entity);
    } catch (e) {
      setError(e.message || String(e));
    }
  };

  const deactivate = async (id) => {
    setError(""); setNotice("");
    try {
      await updateRecord(TABLES.mapping, id, { statecode: 1, statuscode: 2 });
      let stepMessage = "";
      try {
        const res = await deactivateProvisioningStep(entity);
        stepMessage = " " + res.message;
      } catch (e) {
        stepMessage = ` The provisioning step could not be stopped: ${e.message}`;
      }
      setNotice("Configuration deactivated. It is kept and can be reactivated." + stepMessage);
      await loadExisting(entity);
    } catch (e) {
      setError(e.message || String(e));
    }
  };

  const saveFieldMap = async (row) => {
    setError(""); setNotice("");
    try {
      if (!mappingId) throw new Error("Save the configuration before adding field mappings.");
      const payload = await toPayload(FIELD_MAPPING_FIELDS, row);
      const set = await entitySetName(TABLES.mapping);
      payload["bw_docsetmappingid@odata.bind"] = `/${set}(${mappingId})`;

      if (row.id) await updateRecord(TABLES.fieldMapping, row.id, payload);
      else await createRecord(TABLES.fieldMapping, payload);

      setNotice("Field mapping saved.");
      await useExisting(mappingId);
    } catch (e) {
      setError(e.message || String(e));
    }
  };

  const removeFieldMap = async (id) => {
    try {
      await deleteRecord(TABLES.fieldMapping, id);
      await useExisting(mappingId);
    } catch (e) { setError(e.message || String(e)); }
  };

  return (
    <div className="panel">
      {error && <div className="cfg-error">{error}</div>}
      {notice && <div className="cfg-ok">{notice}</div>}

      <div className="cfg-section">
        <label className="ff" style={{ maxWidth: 420 }}>
          <span className="ff-label">Table to configure</span>
          <select className="ff-input" value={entity}
            onChange={(e) => onEntityChange(e.target.value)}>
            <option value="">Select a table…</option>
            {entities.map((e) => (
              <option key={e.logicalName} value={e.logicalName}>
                {e.label} ({e.logicalName})
              </option>
            ))}
          </select>
        </label>

        {entity && stepActive === false && (
          <div className="ds-detail" style={{ marginTop: 8 }}>
            No provisioning step is registered for this table yet. Activating a
            configuration registers it.
          </div>
        )}
      </div>

      {loading && <div className="empty">Loading…</div>}

      {entity && !loading && mode === null && (
        <ChooseMode existing={existing} onUse={useExisting} onNew={startNew}
          onActivate={makeActive} onDeactivate={deactivate} />
      )}

      {entity && !loading && mode !== null && (
        <>
          <div className="cfg-mode">
            <span className={"pill " + (mode === "new" ? "amber" : "blue")}>
              {mode === "new" ? "Creating a new configuration" : "Editing an existing configuration"}
            </span>
            <button type="button" className="link-btn" onClick={() => setMode(null)}>
              ← Choose a different configuration
            </button>
          </div>

          <section className="cfg-section">
            <div className="cfg-head">
              <h4>SharePoint library</h4>
              <span className="ff-hint">
                Enter the site URL and library title; the rest resolves on save.
              </span>
            </div>
            <div className="ff-grid">
              {visible(LIBRARY_FIELDS, libraryCols).map((f) => (
                <Field key={f.name} table={TABLES.library} field={f}
                  value={library?.[f.name]}
                  onChange={(v) => setLibrary({ ...(library || {}), [f.name]: v })} />
              ))}
            </div>
            <div className="cfg-actions">
              <button type="button" className="btn-secondary" onClick={saveLibrary}>
                {libraryId ? "Save library" : "Create library"}
              </button>
            </div>
          </section>

          <section className="cfg-section">
            <div className="cfg-head">
              <h4>Configuration</h4>
              <span className="ff-hint">
                Document set name tokens: {"{primaryname}"}, {"{recordid}"},
                {" {attr:logicalname}"}.
              </span>
            </div>
            <div className="ff-grid">
              {visible(MAPPING_FIELDS, mappingCols)
                .filter((f) => f.name !== `${TABLES.mapping.split("_")[0]}_entitylogicalname`)
                .map((f) => (
                  <Field key={f.name} table={TABLES.mapping} field={f}
                    value={mapping?.[f.name]}
                    onChange={(v) => setMapping({ ...(mapping || {}), [f.name]: v })} />
                ))}
            </div>
            <div className="cfg-actions">
              <button type="button" className="btn-primary" onClick={saveMapping}>
                {mode === "new" ? "Create configuration" : "Update configuration"}
              </button>
            </div>
          </section>

          <section className="cfg-section">
            <div className="cfg-head">
              <h4>Field mappings</h4>
              <span className="ff-hint">
                One row per SharePoint column to populate from the record.
              </span>
            </div>

            {!mappingId && (
              <div className="empty">
                Save the configuration first — field mappings attach to it.
              </div>
            )}

            {mappingId && fieldMaps.map((row, i) => (
              <div className="cfg-row" key={row.id || `new-${i}`}>
                <div className="ff-grid ff-grid-tight">
                  {FIELD_MAPPING_FIELDS.map((f) => (
                    <Field key={f.name} table={TABLES.fieldMapping} field={f}
                      value={row?.[f.name]}
                      onChange={(v) => {
                        const copy = [...fieldMaps];
                        copy[i] = { ...row, [f.name]: v };
                        setFieldMaps(copy);
                      }} />
                  ))}
                </div>
                {columns.length > 0 && (
                  <div className="ff-hint" style={{ marginTop: 6 }}>
                    {columns.length} columns available on this table — use the logical
                    name, not the display name.
                  </div>
                )}
                <div className="cfg-actions">
                  <button type="button" className="btn-secondary"
                    onClick={() => saveFieldMap(fieldMaps[i])}>Save</button>
                  {row.id && (
                    <button type="button" className="btn-danger"
                      onClick={() => removeFieldMap(row.id)}>Remove</button>
                  )}
                </div>
              </div>
            ))}

            {mappingId && (
              <div className="cfg-actions">
                <button type="button" className="btn-secondary"
                  onClick={() => setFieldMaps([...fieldMaps, {}])}>
                  + Add field mapping
                </button>
              </div>
            )}
          </section>
        </>
      )}
    </div>
  );
}

/** The deliberate stop between picking a table and editing anything. */
function ChooseMode({ existing, onUse, onNew, onActivate, onDeactivate }) {
  const activeCount = existing.filter((c) => c.active).length;

  return (
    <div className="cfg-section">
      {existing.length === 0 ? (
        <div className="cfg-choose">
          <div>
            <b>No configuration yet for this table</b>
            <div className="ds-detail">
              Documents cannot be filed against it until one exists.
            </div>
          </div>
          <button type="button" className="btn-primary" onClick={onNew}>
            Create configuration
          </button>
        </div>
      ) : (
        <>
          {activeCount > 1 && (
            <div className="ds-bar err">
              <div>
                <b>{activeCount} configurations are active at once</b>
                <div className="ds-detail">
                  Provisioning will not run until exactly one is active — it will not
                  guess which library documents belong in. Choose one below; the rest
                  are kept, just inactive.
                </div>
              </div>
            </div>
          )}
          {activeCount === 0 && (
            <div className="ds-bar warn">
              <div>
                <b>No active configuration</b>
                <div className="ds-detail">
                  Every configuration for this table is inactive, so documents cannot be
                  filed. Activate one below.
                </div>
              </div>
            </div>
          )}

          <div className="cfg-choose">
            <div>
              <b>{existing.length} configuration{existing.length > 1 ? "s" : ""}</b>
              <div className="ds-detail">
                Open one to review or change it, or start a new one. Creating a new
                configuration leaves the existing ones untouched.
              </div>
            </div>
            <button type="button" className="btn-secondary" onClick={onNew}>
              + Create new configuration
            </button>
          </div>

          <div className="cfg-list">
            {existing.map((c) => (
              <div key={c.id} className={"cfg-list-item" + (c.active ? " is-active" : "")}>
                <button type="button" className="cfg-list-main" onClick={() => onUse(c.id)}>
                  <span className="cfg-list-name">
                    {c.name}
                    <span className={"pill " + (c.active ? "green" : "gray")}>
                      {c.active ? "Active" : "Inactive"}
                    </span>
                  </span>
                  <span className="cfg-list-meta">
                    {c.libraryName || "no library linked"}
                    {c.pattern ? `  ·  ${c.pattern}` : ""}
                  </span>
                </button>
                <div className="cfg-list-actions">
                  {c.active ? (
                    <button type="button" className="btn-secondary"
                      onClick={() => onDeactivate(c.id)}>Deactivate</button>
                  ) : (
                    <button type="button" className="btn-secondary"
                      onClick={() => onActivate(c.id)}>Make active</button>
                  )}
                </div>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}
