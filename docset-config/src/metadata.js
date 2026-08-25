import { api, isMock } from "./xrm";
import { TABLES } from "./schema";

const cache = { entities: null };

/** Tables in this environment that a document set configuration can point at. */
export async function fetchEntities() {
  if (cache.entities) return cache.entities;
  if (isMock()) return [];

  const res = await api(
    "EntityDefinitions?$select=LogicalName,EntitySetName,DisplayName,PrimaryNameAttribute" +
    "&$filter=IsCustomEntity eq true and IsIntersect eq false"
  );

  // Everything custom is offered, not just one publisher's tables. The client's
  // own prefix is unknown to us, and filtering on ours would hide exactly the
  // tables an administrator wants to configure.
  const ours = new Set(Object.values(TABLES));

  const list = (res?.value ?? [])
    .filter((e) => !ours.has(e.LogicalName))
    .map((e) => ({
      logicalName: e.LogicalName,
      setName: e.EntitySetName,
      primaryName: e.PrimaryNameAttribute,
      label: e.DisplayName?.UserLocalizedLabel?.Label || e.LogicalName,
    }))
    .sort((a, b) => a.label.localeCompare(b.label));

  cache.entities = list;
  return list;
}

/**
 * Column logical names on a table.
 *
 * The configuration forms render only columns that exist, so an optional column
 * a client has not created is simply not offered — rather than appearing,
 * accepting input, and failing on save with a platform error.
 */
const columnCache = {};

export async function tableColumns(logicalName) {
  if (columnCache[logicalName]) return columnCache[logicalName];
  if (isMock()) return null;
  const res = await api(
    `EntityDefinitions(LogicalName='${logicalName}')/Attributes?$select=LogicalName`
  );
  const set = new Set((res?.value ?? []).map((a) => a.LogicalName));
  columnCache[logicalName] = set;
  return set;
}

/** Columns on a table that a field mapping could read, for the picker hint. */
export async function mappableColumns(logicalName) {
  if (isMock()) return [];
  try {
    const res = await api(
      `EntityDefinitions(LogicalName='${logicalName}')/Attributes` +
      "?$select=LogicalName,AttributeType,DisplayName,IsValidForRead"
    );
    return (res?.value ?? [])
      .filter((a) => a.IsValidForRead)
      .map((a) => ({
        name: a.LogicalName,
        type: a.AttributeType,
        label: a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName,
      }))
      .sort((a, b) => a.label.localeCompare(b.label));
  } catch (e) {
    return [];
  }
}
