import { entitySetName } from "./xrm";

const FV = "@OData.Community.Display.V1.FormattedValue";

/**
 * Dataverse record to form state.
 *
 * Lookups arrive as a `_col_value` id plus a formatted-value name, so they are
 * rebuilt into the `{ id, name }` shape the lookup control expects. Kept apart
 * from the components so the record form and the admin form share one
 * definition of what "form state" means.
 */
export function toFormState(fields, record) {
  const state = {};
  for (const f of fields) {
    if (f.kind === "lookup") {
      const id = record[`_${f.name}_value`];
      state[f.name] = id
        ? { id, name: record[`_${f.name}_value${FV}`] ?? "(linked)" }
        : null;
    } else {
      state[f.name] = record[f.name] ?? null;
    }
  }
  return state;
}

/** Form state to a Web API payload, resolving lookups to @odata.bind entries. */
export async function toPayload(fields, state) {
  const payload = {};
  for (const f of fields) {
    if (f.derived) continue;           // written by the plug-in, never by the form
    const v = state?.[f.name];

    if (f.kind === "lookup") {
      if (v?.id) {
        const set = await entitySetName(f.lookupTo);
        payload[`${f.name}@odata.bind`] = `/${set}(${v.id})`;
      }
      continue;
    }

    if (v === undefined) continue;
    payload[f.name] = v === "" ? null : v;
  }
  return payload;
}
