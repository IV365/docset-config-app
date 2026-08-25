import { TABLES, KEYS, PROVISION_PLUGIN_TYPE } from "./schema";

const API = "v9.2";

export function getXrm() {
  try { if (window.Xrm && window.Xrm.Utility) return window.Xrm; } catch (e) { /* cross-origin */ }
  try {
    if (window.parent && window.parent.Xrm && window.parent.Xrm.Utility) return window.parent.Xrm;
  } catch (e) { /* cross-origin */ }
  return null;
}

export const isMock = () => getXrm() === null;

const FV = "@OData.Community.Display.V1.FormattedValue";
export const fv = (rec, field) => rec[field + FV] ?? "";

function clientUrl() {
  try { return getXrm().Utility.getGlobalContext().getClientUrl(); }
  catch (e) { return ""; }
}

/**
 * Raw Web API call. Used for everything here rather than Xrm.WebApi, because this
 * app reads metadata and writes sdkmessageprocessingstep rows, neither of which
 * Xrm.WebApi covers. Same-origin inside the app, so the session cookie authenticates it.
 */
export async function api(path, { method = "GET", body, headers = {} } = {}) {
  const res = await fetch(`${clientUrl()}/api/data/${API}/${path}`, {
    method,
    credentials: "include",
    headers: {
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      ...(body ? { "Content-Type": "application/json" } : {}),
      ...headers,
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!res.ok) {
    let detail = await res.text();
    try { detail = JSON.parse(detail)?.error?.message || detail; } catch (e) { /* raw */ }
    throw new Error(`${res.status} ${detail}`);
  }
  if (res.status === 204) return null;
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

/* ---------------- entity set names ---------------- */

const setCache = {};

/**
 * Entity set names come from metadata rather than guessing, because pluralisation
 * is irregular — bw_documentsetlibrary becomes bw_documentsetlibraries — and
 * @odata.bind needs the set name, not the logical name.
 */
export async function entitySetName(logicalName) {
  if (setCache[logicalName]) return setCache[logicalName];
  if (isMock()) return logicalName + "s";
  const res = await api(`EntityDefinitions(LogicalName='${logicalName}')?$select=EntitySetName`);
  setCache[logicalName] = res.EntitySetName;
  return res.EntitySetName;
}

/* ---------------- CRUD ---------------- */

export async function fetchList(table, query) {
  if (isMock()) return [];
  const set = await entitySetName(table);
  const res = await api(set + (query || ""));
  return res?.value ?? [];
}

export async function retrieveRecord(table, id, columns) {
  if (isMock()) return {};
  const set = await entitySetName(table);
  const q = columns?.length ? `?$select=${columns.join(",")}` : "";
  return api(`${set}(${id})${q}`);
}

export async function createRecord(table, data) {
  if (isMock()) { console.info("[mock] create", table, data); return "mock-id"; }
  const set = await entitySetName(table);
  const res = await fetch(`${clientUrl()}/api/data/${API}/${set}`, {
    method: "POST",
    credentials: "include",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Prefer: "return=representation",
    },
    body: JSON.stringify(data),
  });
  if (!res.ok) {
    let d = await res.text();
    try { d = JSON.parse(d)?.error?.message || d; } catch (e) { /* raw */ }
    throw new Error(`${res.status} ${d}`);
  }
  const created = await res.json();
  return created[KEYS[table]?.id] ?? created.id;
}

export async function updateRecord(table, id, data) {
  if (isMock()) { console.info("[mock] update", table, id, data); return; }
  const set = await entitySetName(table);
  await api(`${set}(${id})`, { method: "PATCH", body: data });
}

export async function deleteRecord(table, id) {
  if (isMock()) { console.info("[mock] delete", table, id); return; }
  const set = await entitySetName(table);
  await api(`${set}(${id})`, { method: "DELETE" });
}

/* ---------------- choices ---------------- */

const choiceCache = {};

export async function fetchChoices(table, column) {
  const key = `${table}.${column}`;
  if (choiceCache[key]) return choiceCache[key];
  if (isMock()) return [];
  try {
    const res = await api(
      `EntityDefinitions(LogicalName='${table}')/Attributes(LogicalName='${column}')` +
      `/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName` +
      `&$expand=OptionSet($select=Options)`
    );
    const options = (res?.OptionSet?.Options ?? []).map((o) => ({
      value: o.Value,
      label: o.Label?.UserLocalizedLabel?.Label ?? String(o.Value),
    }));
    choiceCache[key] = options;
    return options;
  } catch (e) {
    console.warn(`Choice metadata failed for ${key}:`, e);
    return [];
  }
}

export async function searchLookup(table, term) {
  if (isMock()) return [];
  const keys = KEYS[table];
  if (!keys) throw new Error(`No key definition for '${table}'.`);
  const safe = (term || "").replace(/'/g, "''");
  const filter = safe ? `&$filter=contains(${keys.name},'${safe}')` : "";
  const rows = await fetchList(
    table,
    `?$select=${keys.id},${keys.name}${filter}&$orderby=${keys.name} asc&$top=20`
  );
  return rows.map((r) => ({ id: r[keys.id], name: r[keys.name] ?? "(no name)" }));
}

/* ================================================================== *
 * Plug-in step registration
 * ================================================================== */

/**
 * Registers the provisioning plug-in against one table, or reports that it already is.
 *
 * WHY THIS EXISTS
 *
 * A plug-in step names its table, so steps for a client's tables cannot ship inside a
 * managed solution — the table does not exist until the client has it. Without this, every
 * deployment needs a developer and the Plug-in Registration Tool, which is the difference
 * between a product and a bespoke build.
 *
 * The step is created in the client's own environment, unmanaged. That is correct rather
 * than merely tolerable: it is configuration, and it changes when the configuration does.
 *
 * NEEDS PRIVILEGE: writing sdkmessageprocessingstep requires system administrator or
 * system customizer. Configuring document sets is an administrator task anyway.
 */
export async function registerProvisioningStep(entityLogicalName) {
  if (isMock()) {
    console.info("[mock] register step for", entityLogicalName);
    return { created: false, message: "Preview mode — no step registered." };
  }

  const stepName =
    `Monarch DocSet: provision on create of ${entityLogicalName}`;

  // Already registered? Registering twice would provision twice per record.
  const existing = await api(
    `sdkmessageprocessingsteps?$select=sdkmessageprocessingstepid,statecode` +
    `&$filter=name eq '${stepName.replace(/'/g, "''")}'&$top=1`
  );
  if (existing?.value?.length) {
    const step = existing.value[0];
    if (step.statecode !== 0) {
      await api(`sdkmessageprocessingsteps(${step.sdkmessageprocessingstepid})`, {
        method: "PATCH",
        body: { statecode: 0, statuscode: 1 },
      });
      return { created: false, message: "Existing provisioning step reactivated." };
    }
    return { created: false, message: "Provisioning step was already registered." };
  }

  const types = await api(
    `plugintypes?$select=plugintypeid&$filter=typename eq '${PROVISION_PLUGIN_TYPE}'&$top=1`
  );
  if (!types?.value?.length)
    throw new Error(
      `The plug-in type '${PROVISION_PLUGIN_TYPE}' is not registered in this environment. ` +
      "Import the solution containing the plug-in assembly first."
    );

  // The filter ties the Create message to this specific table; without it the step
  // would have to run on every table and filter in code.
  const filters = await api(
    `sdkmessagefilters?$select=sdkmessagefilterid` +
    `&$filter=primaryobjecttypecode eq '${entityLogicalName}'` +
    ` and sdkmessageid/name eq 'Create'&$top=1`
  );
  if (!filters?.value?.length)
    throw new Error(
      `No Create message filter exists for '${entityLogicalName}'. Check the logical name.`
    );

  await api("sdkmessageprocessingsteps", {
    method: "POST",
    body: {
      name: stepName,
      description: "Creates the SharePoint document set when a record is created.",
      stage: 40,          // post-operation
      mode: 1,            // asynchronous — SharePoint must never block a save
      rank: 1,
      supporteddeployment: 0,
      invocationsource: 0,
      "plugintypeid@odata.bind":
        `/plugintypes(${types.value[0].plugintypeid})`,
      "sdkmessageid@odata.bind":
        `/sdkmessages(${await createMessageId()})`,
      "sdkmessagefilterid@odata.bind":
        `/sdkmessagefilters(${filters.value[0].sdkmessagefilterid})`,
      // Keep the System Job on success while a deployment is being verified.
      asyncautodelete: false,
    },
  });

  return { created: true, message: `Provisioning step registered for ${entityLogicalName}.` };
}

let createMessageCache = null;
async function createMessageId() {
  if (createMessageCache) return createMessageCache;
  const res = await api(
    "sdkmessages?$select=sdkmessageid&$filter=name eq 'Create'&$top=1"
  );
  if (!res?.value?.length) throw new Error("Could not resolve the Create message.");
  createMessageCache = res.value[0].sdkmessageid;
  return createMessageCache;
}

/** Deactivates the step, so a deactivated configuration stops provisioning. */
export async function deactivateProvisioningStep(entityLogicalName) {
  if (isMock()) return { message: "Preview mode — nothing changed." };

  const stepName = `Monarch DocSet: provision on create of ${entityLogicalName}`;
  const existing = await api(
    `sdkmessageprocessingsteps?$select=sdkmessageprocessingstepid` +
    `&$filter=name eq '${stepName.replace(/'/g, "''")}' and statecode eq 0&$top=1`
  );
  if (!existing?.value?.length) return { message: "No active provisioning step to stop." };

  await api(`sdkmessageprocessingsteps(${existing.value[0].sdkmessageprocessingstepid})`, {
    method: "PATCH",
    body: { statecode: 1, statuscode: 2 },
  });
  return { message: "Provisioning step deactivated." };
}

/** Whether a step is currently active for this table, for the status badge. */
export async function provisioningStepState(entityLogicalName) {
  if (isMock()) return false;
  try {
    const stepName = `Monarch DocSet: provision on create of ${entityLogicalName}`;
    const res = await api(
      `sdkmessageprocessingsteps?$select=sdkmessageprocessingstepid` +
      `&$filter=name eq '${stepName.replace(/'/g, "''")}' and statecode eq 0&$top=1`
    );
    return !!res?.value?.length;
  } catch (e) {
    // Reading these needs privilege the user may not have. Not knowing is not a
    // reason to block configuration, so report unknown rather than failing.
    return null;
  }
}
