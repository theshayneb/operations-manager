# Changelog

## [0.4.90] - 2026-10-16

- `Type: license` is enough to count seats: per-seat licenses no longer need `Agreement_Types: [software seats]`, and new license notes don't get it. Existing notes that have it keep working, and you can delete it from notes that have `Type: license`.

## [0.4.89] - 2026-10-16

- New service notes include `Link`, `Email`, `Account_Number` and `Attachments`.
- `Billing` can be `One-off` for services that only charge when you use them. One-off services have no billing schedule, so they never raise a charge alert. "Edit billing schedule" offers One-off too.

## [0.4.88] - 2026-10-16

- New notes use a capitalized `Type` property, matching the other properties, and the settings guide shows `Type`. Notes with a lowercase `type` keep working.

## [0.4.87] - 2026-10-15

### Less clutter in your notes
- The plugin no longer adds `Next_Renewal` to notes by itself. Renewal dates estimated from `Ordered` + `Renewal` are worked out when shown.
- "Update" on equipment saves only the fields you change, under the property names the note already uses, so it no longer adds `Device_Model`, `Serial_Number`, `Set_Up` and friends next to your existing ones. It no longer changes Status on its own either.
- Alert on/off switches are kept in the plugin settings. Old `Renewal_Alerts`, `Billing_Alerts` and `Seat_Alerts` properties still work and are removed from the note when you change the switch.
- "Record charge" and "Renew" keep their history in the plugin settings instead of adding `Billing_History`, `Last_Billed`, `Renewal_History`, `Last_Renewed` and `Renewal_Review_Required`. The last dates are shown in the plugin settings.
- "Renew" only adds a follow-up task to the note if you turn that on in settings (off by default), and the task now has proper line breaks.
- License assignments reuse the employee's existing checkbox, and bundle assignments no longer add empty `Nextiva_Licenses` lists or unticked checkboxes nobody asked for.

### Note types
- New `type` property: `computer`, `peripheral`, `equipment`, `service`, `contract`, `license`, `invoice`, `order` or `employee`. Notes that use tags keep working.
- New "New note" button and "Create a new operations note" command: pick what you're tracking and get a note with exactly the properties that type needs.
- Settings has a "Note types & properties" guide listing the properties for each type.
- Cards show a small label for peripherals, office equipment, services, contracts and licenses.
- Services aren't flagged for a missing renewal date (their contract has it), and neither are licenses billed on a schedule.
- License bundles (several license types making up one bill) work for any vendor, not only notes named Nextiva.

### Fixes
- Per-note settings follow a note when it is renamed or moved.
- Fixed `[[link]]` matching and spaces in the icon search.
