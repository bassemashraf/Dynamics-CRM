---
description: Offline JavaScript API guidelines for Dynamics CRM Xrm.WebApi calls
---

# Offline JavaScript API Guidelines (Dynamics CRM)

// turbo-all

## Core Rules

### `$select` — Always use `_fieldname_value`

No switching needed. Use `_fieldname_value` for lookup fields in `$select` for **both online and offline**.

```javascript
// ✅ CORRECT — same for online and offline
var record = await Xrm.WebApi.retrieveRecord("msdyn_workorder", id,
    "?$select=_duc_processextension_value"
);
```

### `$filter` — Switch for offline

Use schema name (e.g. `duc_process`) when offline, `_fieldname_value` when online.

```javascript
// ✅ CORRECT — switch only in $filter
var isOff = isOffline();
var filterField = isOff ? "msdyn_workorder" : "_msdyn_workorder_value";
var result = await Xrm.WebApi.retrieveMultipleRecords("entity",
    `?$select=fieldid&$filter=${filterField} eq ${id}`
);
```

### Response properties — Always `_fieldname_value`

```javascript
var value = record._duc_processextension_value; // always
```

### `@odata.bind` — Same syntax for both modes

```javascript
await Xrm.WebApi.updateRecord("entity", id, {
    "navprop@odata.bind": `/entities(${relatedId})`
});
```

### No `$expand` — Make separate calls instead

## Quick Reference

| Context | Online | Offline |
|---|---|---|
| `$select` lookup fields | `_fieldname_value` | `_fieldname_value` |
| `$filter` lookup fields | `_fieldname_value` | `fieldname` (schema) |
| Response property | `_fieldname_value` | `_fieldname_value` |
| `@odata.bind` | Same | Same |
