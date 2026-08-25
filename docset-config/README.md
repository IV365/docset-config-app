# Document Set Configuration

Administrator app for the Monarch document set provisioning solution. Configures which
Dataverse tables file documents into which SharePoint libraries, and with what metadata.

Deliberately separate from any client application: this is product configuration, and
embedding it in one client's dashboard made it look like that client's feature.

## Build

```
yarn install
yarn build
```

Produces a single `dist/index.html`. Upload it as a web resource and point a model-driven
app page at it.

## What it configures

| Table | Holds |
|---|---|
| Document Set Library | One row per SharePoint library — site, library and credentials |
| Document Set Mapping | One row per Dataverse table — which library, how to name sets |
| Document Set Field Mapping | One row per SharePoint column to populate |

## Deployment

See `../monarch/docs/07-productisation-plan.md`. In short: the solution ships the tables,
plug-ins, Custom API and this app; the client supplies an app registration, a SharePoint
library, and their table choice through this UI.

Activating a configuration registers the provisioning plug-in step for that table
automatically, so no Plug-in Registration Tool is needed in a client environment.
