/**
 * Schema for the document set configuration solution.
 *
 * PREFIX is the one place the publisher prefix appears. Everything else is built
 * from it, so retargeting the solution at a different publisher is a single edit
 * here plus recreating the tables under that publisher in Dataverse.
 */
export const PREFIX = "bw_";

export const TABLES = {
  library: `${PREFIX}documentsetlibrary`,
  mapping: `${PREFIX}documentsetmapping`,
  fieldMapping: `${PREFIX}documentsetfieldmapping`,
  upload: `${PREFIX}documentupload`,
  provisioning: `${PREFIX}docsetprovisioning`,
};

/** Primary key and primary name per table, for lookups and list rendering. */
export const KEYS = {
  [TABLES.library]: { id: `${PREFIX}documentsetlibraryid`, name: `${PREFIX}documentsetlibraryname` },
  [TABLES.mapping]: { id: `${PREFIX}documentsetmappingid`, name: `${PREFIX}name` },
  [TABLES.fieldMapping]: { id: `${PREFIX}documentsetfieldmappingid`, name: `${PREFIX}name` },
};

/**
 * The plug-in class registered per configured table.
 *
 * Activating a configuration registers a step against this type, which is what
 * removes the Plug-in Registration Tool from client deployments.
 */
export const PROVISION_PLUGIN_TYPE =
  "Monarch.DocSet.Plugins.Plugins.ProvisionOnCreatePlugin";

/** Columns on the library. Derived ones are filled by the plug-in on save. */
export const LIBRARY_FIELDS = [
  { name: `${PREFIX}documentsetlibraryname`, label: "Name", kind: "text", required: true },
  { name: `${PREFIX}sharepointsiteurl`, label: "SharePoint site URL", kind: "text", required: true },
  { name: `${PREFIX}documentlibrarytitle`, label: "Document library title", kind: "text", required: true },
  { name: `${PREFIX}siteserverrelativeurl`, label: "Site server relative URL", kind: "text", derived: true },
  { name: `${PREFIX}graphsiteid`, label: "Graph site ID", kind: "text", derived: true },
  { name: `${PREFIX}documentlibraryid`, label: "Document library ID", kind: "text", derived: true },
  { name: `${PREFIX}libraryserverrelativeurl`, label: "Library server relative URL", kind: "text", derived: true },
  { name: `${PREFIX}docsetcontenttypeid`, label: "Doc set content type ID", kind: "text", derived: true },
  { name: `${PREFIX}tenantid`, label: "Tenant ID", kind: "text" },
  { name: `${PREFIX}clientid`, label: "Client ID", kind: "text" },
  { name: `${PREFIX}secretenvvarname`, label: "Secret env var name", kind: "text" },
];

export const MAPPING_FIELDS = [
  { name: `${PREFIX}name`, label: "Name", kind: "text", required: true },
  { name: `${PREFIX}entitylogicalname`, label: "Entity logical name", kind: "text", required: true },
  { name: `${PREFIX}entitydisplayname`, label: "Entity display name", kind: "text" },
  { name: `${PREFIX}docsetlibraryid`, label: "Doc set library", kind: "lookup",
    lookupTo: TABLES.library, required: true },
  { name: `${PREFIX}namingpattern`, label: "Document set name", kind: "text" },
  { name: `${PREFIX}recordidcolumn`, label: "Record ID column (optional)", kind: "text" },
  { name: `${PREFIX}updatesyncenabled`, label: "Update sync enabled", kind: "bool" },
  { name: `${PREFIX}deletebehaviour`, label: "Delete behaviour", kind: "choice" },
];

export const FIELD_MAPPING_FIELDS = [
  { name: `${PREFIX}name`, label: "Name", kind: "text" },
  { name: `${PREFIX}dataverseattribute`, label: "Dataverse column", kind: "text", required: true },
  { name: `${PREFIX}sharepointfield`, label: "SharePoint column", kind: "text", required: true },
  { name: `${PREFIX}sharepointfieldtype`, label: "Field type", kind: "choice" },
  { name: `${PREFIX}transform`, label: "Transform", kind: "choice" },
  { name: `${PREFIX}defaultvalue`, label: "Default value", kind: "text" },
];
