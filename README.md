# Operations Manager for Obsidian

A beta Obsidian plugin for tracking equipment, licenses, renewals, and invoices in Markdown notes.

## Installation with BRAT

1. Install the BRAT community plugin in Obsidian.
2. Choose **Add Beta Plugin**.
3. Enter `theshayneb/operations-manager`.
4. Enable **Operations Manager** in Community plugins.

## Note types

Give each note a `Type` property so the plugin knows what it is:

| Type | For | Shows on |
|---|---|---|
| `computer` | A laptop or desktop assigned to an employee. Add the `equip_ug` tag for replacements in that project. | Equipment |
| `peripheral` | Mouse, keyboard, monitor, dock… | Equipment |
| `equipment` | Printers, postage machines | Equipment |
| `service` | Off-site storage, shredding, IT company. Billing can be One-off for pay-per-use services. | Licenses & Renewals |
| `contract` | A lease or yearly agreement with a renewal date | Licenses & Renewals |
| `license` | Per-seat software, or a bundle of license types making up one bill | Licenses & Renewals |
| `invoice` | A bill | Invoices |
| `order` | A purchase | Orders |
| `employee` | A person, with one checkbox per license they hold | – |

The **New note** button (or the "Create a new operations note" command) creates a note with the right properties for each type. **Settings → Note types & properties** lists them all. A lowercase `type` property works too, and notes that use tags instead keep working. To add `Type` to the notes you already have, run the **Add Type to existing notes** command, check its suggestions, and confirm. Then run **Check notes for missing properties** to add any properties each note is missing and see which values still need filling in.

**Status:** Experimental. Back up your vault before use. BRAT installation requires a published GitHub release with `main.js`, `manifest.json`, and `styles.css` attached.
