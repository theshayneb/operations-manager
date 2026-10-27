function xlsxWorkbook(columns, records) {
  const xmlEscape = (v) =>
    String(v ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  const colName = (n) => {
    let a = "";
    for (let x = n + 1; x; x = Math.floor((x - 1) / 26)) a = String.fromCharCode(65 + ((x - 1) % 26)) + a;
    return a;
  };
  const data = [columns, ...records];
  const rows = data
    .map(
      (values, i) =>
        '<row r="' +
        (i + 1) +
        '">' +
        values
          .map((v, j) => {
            const ref = colName(j) + (i + 1),
              style = i === 0 ? ' s="1"' : "";
            const numeric =
              i > 0 && columns[j] === "Cost" && v !== "" && v != null && Number.isFinite(Number(v));
            return numeric
              ? '<c r="' + ref + '" s="2"><v>' + Number(v) + "</v></c>"
              : '<c r="' + ref + '" t="inlineStr"' + style + "><is><t>" + xmlEscape(v) + "</t></is></c>";
          })
          .join("") +
        "</row>",
    )
    .join("");
  const widths = columns
    .map(
      (c, i) =>
        '<col min="' +
        (i + 1) +
        '" max="' +
        (i + 1) +
        '" width="' +
        Math.min(48, Math.max(13, c.length + 3)) +
        '" customWidth="1"/>',
    )
    .join("");
  const sheet =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:' +
    colName(columns.length - 1) +
    data.length +
    '"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>' +
    widths +
    "</cols><sheetData>" +
    rows +
    '</sheetData><autoFilter ref="A1:' +
    colName(columns.length - 1) +
    data.length +
    '"/></worksheet>';
  const files = {
    "[Content_Types].xml":
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    "_rels/.rels":
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    "xl/workbook.xml":
      '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Operations Records" sheetId="1" r:id="rId1"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels":
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    "xl/styles.xml":
      '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF244466"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>',
    "xl/worksheets/sheet1.xml": sheet,
  };
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = crcTable[(c ^ b) & 255] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const local = [],
    central = [];
  let offset = 0;
  const u16 = (b, n, p) => b.writeUInt16LE(n, p),
    u32 = (b, n, p) => b.writeUInt32LE(n >>> 0, p);
  for (const [name, content] of Object.entries(files)) {
    const nameBuf = Buffer.from(name),
      body = Buffer.from(content),
      sum = crc(body);
    const lh = Buffer.alloc(30);
    u32(lh, 0x04034b50, 0);
    u16(lh, 20, 4);
    u16(lh, 0, 6);
    u16(lh, 0, 8);
    u32(lh, sum, 14);
    u32(lh, body.length, 18);
    u32(lh, body.length, 22);
    u16(lh, nameBuf.length, 26);
    local.push(lh, nameBuf, body);
    const ch = Buffer.alloc(46);
    u32(ch, 0x02014b50, 0);
    u16(ch, 20, 4);
    u16(ch, 20, 6);
    u32(ch, sum, 16);
    u32(ch, body.length, 20);
    u32(ch, body.length, 24);
    u16(ch, nameBuf.length, 28);
    u32(ch, offset, 42);
    central.push(ch, nameBuf);
    offset += lh.length + nameBuf.length + body.length;
  }
  const centralSize = central.reduce((n, b) => n + b.length, 0);
  const end = Buffer.alloc(22);
  u32(end, 0x06054b50, 0);
  u16(end, Object.keys(files).length, 8);
  u16(end, Object.keys(files).length, 10);
  u32(end, centralSize, 12);
  u32(end, offset, 16);
  return Buffer.concat([...local, ...central, end]);
}
const { Plugin, ItemView, Modal, Setting, Notice, normalizePath } = require("obsidian");
const VIEW = "operations-manager-view";
// The loaded plugin settings, for helpers that only receive a note's frontmatter.
let OM_SETTINGS = null;
const dateISO = (d) => d.toISOString().slice(0, 10);
function addMonths(iso, months) {
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return null;
  const target = new Date(Date.UTC(y, m - 1 + months, 1));
  const last = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();
  target.setUTCDate(Math.min(d, last));
  return dateISO(target);
}
function tagSet(fm, file, app) {
  const values = [fm.tags].flatMap((v) =>
    Array.isArray(v) ? v : typeof v === "string" ? v.split(/[,\s]+/) : [],
  );
  return new Set(values.map((x) => String(x).trim().replace(/^#/, "").toLowerCase()).filter(Boolean));
}
function omTerminated(fm) {
  const value = fm.Termination_Date ?? fm.termination_date;
  return (
    value !== undefined &&
    value !== null &&
    String(value).trim() !== "" &&
    String(value).trim().toLowerCase() !== "null"
  );
}
function omLicenseAssigned(value) {
  return (
    value === true ||
    String(value ?? "")
      .trim()
      .toLowerCase() === "checked"
  );
}
function iso(v) {
  if (v instanceof Date) return dateISO(v);
  if (typeof v === "string") return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : null;
  return null;
}
function omMonthToMonth(fm, file) {
  const fields = [fm.Billing_Cycle, fm.Renewal_Type, fm.Renewal_Term, fm.Contract_Type, fm.Renewal];
  return fields.some((v) => /^(month[ -]?to[ -]?month|monthly|m2m)$/i.test(String(v ?? "").trim()));
}
function inferredRenewal(fm) {
  const explicit = iso(fm.Next_Renewal);
  if (explicit) return { date: explicit, estimated: false };
  const ordered = iso(fm.Ordered);
  const term = Number(fm.Renewal);
  if (!ordered || !Number.isInteger(term) || term <= 0 || term > 36500) return null;
  const d = new Date(ordered + "T00:00:00Z");
  if (!Number.isFinite(d.getTime())) return null;
  d.setUTCDate(d.getUTCDate() + term);
  return { date: dateISO(d), estimated: true };
}
function daysUntil(d) {
  return Math.round(
    (Date.parse(d + "T00:00:00Z") - Date.parse(dateISO(new Date()) + "T00:00:00Z")) / 86400000,
  );
}
// Note types. A note's `Type` property (or lowercase `type`) says what it is; tags keep working for notes without one.
// Each property: [name, kind, what to enter, default value for new notes]. Kinds: text, link, date,
// number, checkbox, list. Properties marked optional are listed in the guide but not added to new notes.
const OM_TYPE_ALIASES = {
  computer: ["computer", "laptop", "desktop", "workstation", "pc", "device"],
  peripheral: ["peripheral", "monitor", "keyboard", "mouse", "dock", "headset", "accessory"],
  equipment: ["equipment", "office equipment", "printer", "copier", "postage machine", "machine"],
  service: ["service", "vendor"],
  contract: ["contract", "agreement", "lease"],
  license: ["license", "licence", "subscription", "software", "license bundle"],
  invoice: ["invoice", "bill"],
  order: ["order", "purchase", "purchase order"],
  employee: ["employee", "person", "staff"],
};
const OM_TYPE_LABELS = {
  computer: "Computer",
  peripheral: "Peripheral",
  equipment: "Office equipment",
  service: "Service",
  contract: "Contract",
  license: "License",
  invoice: "Bill",
  order: "Order",
  employee: "Employee",
};
const OM_STATUS = "Status";
const OM_TEMPLATES = [
  {
    id: "computer",
    type: "computer",
    label: "Computer",
    description: "A laptop or desktop assigned to one employee. Shows on the Equipment tab.",
    props: [
      [OM_STATUS, "text", "deployed, spare, retired, to do, ordered, delivered…", "deployed"],
      ["Employee", "link", "Who has it, as a link: [[Jane Doe]]."],
      ["Model", "text", "Make and model."],
      ["Serial", "text", "Serial number."],
      ["Set_Up", "date", "Date it was set up for the employee."],
      ["Notes", "text", "Anything else."],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
      ["Ordered", "date", "Date it was ordered.", undefined, true],
    ],
  },
  {
    id: "replacement",
    type: "computer",
    tags: ["equip_ug"],
    label: "Replacement computer (equip_ug)",
    description:
      'A new computer replacing an old one as part of the equip_ug project. The equip_ug tag groups these by phase on the Dashboard ("Equipment by Phase").',
    props: [
      [OM_STATUS, "text", "to do, ordered, delivered, complete…", "to do"],
      [
        "Approved",
        "text",
        'Who approved it and the phase, e.g. Approved by Sam · Phase 2. The "Phase …" part is the group.',
      ],
      ["Employee", "link", "Who it is for, as a link: [[Jane Doe]]."],
      ["To_Replace", "link", "The old computer's note, as a link."],
      ["Model", "text", "Make and model."],
      ["Serial", "text", "Serial number."],
      ["Ordered", "date", "Date it was ordered."],
      ["Set_Up", "date", "Date it was set up for the employee."],
      ["Old_Device_Returned", "checkbox", "Ticked once the old computer is back.", false],
      ["Notes", "text", "Anything else."],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
    ],
  },
  {
    id: "peripheral",
    type: "peripheral",
    label: "Peripheral",
    description:
      "A mouse, keyboard, monitor, dock or headset assigned to an employee. Shows on the Equipment tab.",
    props: [
      [OM_STATUS, "text", "deployed, spare, retired…", "deployed"],
      ["Employee", "link", "Who has it, as a link: [[Jane Doe]]."],
      ["Model", "text", "Make and model."],
      ["Serial", "text", "Serial number, if it has one."],
      ["Notes", "text", "Anything else."],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
    ],
  },
  {
    id: "equipment",
    type: "equipment",
    label: "Office equipment",
    description:
      "A shared machine such as a printer or postage machine. Its lease or service agreement goes in a separate Contract note.",
    props: [
      [OM_STATUS, "text", "deployed, spare, retired…", "deployed"],
      ["Vendor", "text", "Who supplies or services it."],
      ["Model", "text", "Make and model."],
      ["Serial", "text", "Serial number."],
      ["Notes", "text", "Location, and anything else."],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
    ],
  },
  {
    id: "service",
    type: "service",
    label: "Service",
    description:
      "An ongoing vendor service such as off-site storage, shredding or your IT company. Shows on the Licenses & Renewals tab and warns before each charge. Its yearly agreement goes in a separate Contract note.",
    props: [
      [OM_STATUS, "text", "active, cancelled…", "active"],
      ["Vendor", "text", "Company name."],
      ["Cost", "number", "Amount charged each billing cycle. Not needed for One-off: each bill has its own amount."],
      [
        "Billing",
        "text",
        "Monthly, Quarterly, Semiannually, Annually, or One-off for services that only charge when used.",
        "Monthly",
      ],
      ["Next_Billing", "date", "Date of the next charge. Leave blank for One-off."],
      ["Link", "text", "Website or customer portal."],
      ["Email", "text", "Contact or billing email address."],
      ["Account_Number", "text", "Your account number with the vendor."],
      [
        "Attachments",
        "list",
        "Files such as signed forms or price sheets: drag them in, or add [[links]].",
        [],
      ],
      ["Notes", "text", "Contacts, and anything else."],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
    ],
  },
  {
    id: "contract",
    type: "contract",
    label: "Contract",
    description:
      "An agreement with a renewal date: a printer lease, or the yearly contract for a service. Shows on the Licenses & Renewals tab and warns before it renews.",
    props: [
      [OM_STATUS, "text", "active, cancelled…", "active"],
      ["Vendor", "text", "Company name."],
      ["Next_Renewal", "date", "When it next renews or expires."],
      [
        "Renewal",
        "number",
        "Length of each term in days, e.g. 365 (or month-to-month, or variable). Renew moves Next_Renewal forward by this many days.",
        365,
      ],
      [
        "Ordered",
        "date",
        "Start date. With Renewal, gives an estimated renewal date while Next_Renewal is empty.",
        undefined,
        true,
      ],
      ["Cost", "number", "Amount per billing cycle (or per term)."],
      ["Notes", "text", "What it covers, and anything else."],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
      [
        "Billing",
        "text",
        "Only if the contract itself is billed, e.g. a lease: Monthly, Quarterly…",
        undefined,
        true,
      ],
      [
        "Next_Billing",
        "date",
        "Only if the contract itself is billed: date of the next charge.",
        undefined,
        true,
      ],
    ],
  },
  {
    id: "license",
    type: "license",
    label: "Per-seat license",
    description:
      "Software with seats assigned to individual employees, such as ChatGPT. Each employee who has a seat gets a ticked checkbox property named after License_Field.",
    props: [
      [OM_STATUS, "text", "active, cancelled…", "active"],
      ["Vendor", "text", "Company name."],
      ["License_Seats", "number", "Number of seats purchased."],
      [
        "License_Field",
        "text",
        "Name of the checkbox property on employee notes. Defaults to the note's name.",
      ],
      ["Cost", "number", "Amount charged each billing cycle."],
      ["Billing", "text", "Monthly, Quarterly, Semiannually or Annually.", "Monthly"],
      ["Next_Billing", "date", "Date of the next charge."],
      ["Next_Renewal", "date", "When the subscription renews, if it has a term."],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
    ],
  },
  {
    id: "bundle",
    type: "license",
    label: "License bundle (e.g. phone service)",
    description:
      'One vendor bill made up of several license types assigned to employees, like your phone service. After creating it, use "Manage license types" on the Licenses & Renewals tab to add each type, its monthly price, seat count and employee checkbox property. The yearly agreement goes in a separate Contract note.',
    props: [
      [OM_STATUS, "text", "active, cancelled…", "active"],
      ["Vendor", "text", "Company name."],
      ["Cost", "number", "Total amount charged each billing cycle."],
      ["Billing", "text", "Monthly, Quarterly, Semiannually or Annually.", "Monthly"],
      ["Next_Billing", "date", "Date of the next charge."],
      ["License_Types", "list", 'Filled in by "Manage license types".', []],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
    ],
  },
  {
    id: "invoice",
    type: "invoice",
    label: "Bill / invoice",
    description:
      "A single bill: for a service, an office machine or licenses. Shows on the Invoices tab until it is marked paid.",
    props: [
      [OM_STATUS, "text", "unpaid or paid.", "unpaid"],
      ["Vendor", "text", "Company name."],
      ["Date", "date", "Invoice date."],
      ["Due_Date", "date", "Payment due date."],
      ["Cost", "number", "Amount."],
      ["Notes", "text", 'What it is for. Shown as "Services".'],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
      ["Paid_Date", "date", "Set by Mark paid.", undefined, true],
    ],
  },
  {
    id: "order",
    type: "order",
    label: "Order",
    description:
      "A purchase that is not part of the equip_ug project. Shows on the Orders tab until complete.",
    props: [
      [OM_STATUS, "text", "to do, ordered, delivered, complete…", "to do"],
      ["Vendor", "text", "Who it was ordered from."],
      ["Approved", "text", "Who approved it and when."],
      ["Ordered", "date", "Date it was ordered."],
      ["Employee", "link", "Who it is for, as a link: [[Jane Doe]]."],
      ["Notes", "text", "Anything else."],
      ["Related", "list", "Links to connected notes, one per item: [[Office printer]], [[Printer lease]]…", []],
    ],
  },
  {
    id: "employee",
    type: "employee",
    label: "Employee",
    description:
      "A person. New employee notes get an unticked checkbox for every per-seat license and license-bundle type that uses a checkbox property; tick the ones they have.",
    props: [
      [
        "Termination_Date",
        "date",
        "Fill in when they leave: their licenses stop being counted.",
        undefined,
        true,
      ],
    ],
  },
];
function omNoteType(fm) {
  const raw = String(fm?.Type ?? fm?.type ?? "")
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, " ");
  if (!raw) return null;
  for (const [type, aliases] of Object.entries(OM_TYPE_ALIASES)) if (aliases.includes(raw)) return type;
  return null;
}
// Services keep their renewal date in a separate contract note, and a typed license that is billed on a
// schedule may have no term, so only nag about a missing renewal date where one is expected.
function omNeedsRenewalDate(fm) {
  const type = omNoteType(fm);
  if (type === "service") return false;
  if (type === "license" && omBillingInfo(fm)) return false;
  return true;
}
function omIsEmployee(e) {
  return (
    omNoteType(e.fm) === "employee" ||
    tagSet(e.fm).has("employee") ||
    e.file.path.split("/").some((p) => /^employees?$/i.test(p))
  );
}
function classified(file, fm) {
  const type = omNoteType(fm);
  if (type === "computer" || type === "peripheral" || type === "equipment") return "device";
  if (type === "service" || type === "contract" || type === "license") return "renewal";
  if (type === "invoice" || type === "order") return type;
  if (type === "employee") return null;
  const t = tagSet(fm),
    p = file.path.toLowerCase();
  if (t.has("ops-project") || (t.has("project") && p.includes("project"))) return "project";
  if (t.has("invoice") || ("/" + p).includes("/invoices/")) return "invoice";
  if (
    t.has("license") ||
    t.has("licenses") ||
    t.has("subscription") ||
    t.has("renewal") ||
    t.has("software_license") ||
    (t.has("printer") && (fm.Renewal != null || fm.Next_Renewal != null || fm.Agreement_Types != null)) ||
    fm.Agreement_Types != null ||
    fm.License_Seats != null ||
    fm.License_Field != null ||
    fm.License_Types != null ||
    fm.Next_Renewal != null ||
    (fm.Next_Billing != null && fm.Billing != null) ||
    (fm.License_Type != null &&
      /maintenance|lease|agreement|software|subscription/i.test(String(fm.License_Type)))
  )
    return "renewal";
  if (t.has("equip_up") || t.has("equip_upgrade") || t.has("equip_ug")) return "order";
  if (t.has("order")) return "order";
  if (t.has("device")) return "device";
  return null;
}
function status(fm) {
  return String(fm.Status ?? fm.status ?? "")
    .trim()
    .toLowerCase();
}

function omLinkValue(v) {
  if (v == null) return "";
  if (Array.isArray(v)) return omLinkValue(v[0]);
  if (typeof v === "object") return String(v.path ?? v.link ?? "");
  return String(v).replace(/^\[\[/, "").replace(/\]\]$/, "").split("|")[0].trim();
}
function omKey(v) {
  return omLinkValue(v).replace(/\\/g, "/").replace(/\.md$/i, "").toLowerCase();
}
function omMatches(ref, file) {
  const k = omKey(ref);
  return (
    !!k &&
    (k === omKey(file.path) || k === omKey(file.basename) || k === omKey(file.path.split("/").slice(-1)[0]))
  );
}
function omRelated(fm, field) {
  const v = fm[field];
  return Array.isArray(v) ? v : v == null ? [] : [v];
}
class OMDeviceWorkflowModal extends Modal {
  constructor(app, plugin, record) {
    super(app);
    this.plugin = plugin;
    this.record = record;
    const fm = record.fm;
    const text = (v) => (v == null ? "" : Array.isArray(v) ? v.map(String).join(", ") : String(v));
    this.status = text(fm.Status ?? fm.status) || "in progress";
    this.employee = text(fm.Employee ?? fm.Assigned_To);
    this.serial = text(fm.Serial_Number ?? fm.Serial);
    this.model = text(fm.Device_Model ?? fm.Model);
    this.orderDate = iso(fm.Ordered ?? fm.Order_Date) ?? "";
    this.received = iso(fm.Set_Up ?? fm.Set_up) ?? "";
    this.oldReturned = !!fm.Old_Device_Returned;
    this.initial = {
      status: this.status,
      employee: this.employee,
      serial: this.serial,
      model: this.model,
      orderDate: this.orderDate,
      received: this.received,
      oldReturned: this.oldReturned,
    };
  }
  onOpen() {
    const c = this.contentEl;
    c.createEl("h2", { text: "Update equipment workflow" });
    c.createEl("p", { text: this.record.file.basename });
    const choices = [
      "in progress",
      "awaiting approval",
      "ordered",
      "shipped",
      "delivered",
      "complete",
      "cancelled",
    ];
    const current = this.status;
    new Setting(c)
      .setName("Status")
      .setDesc("Only the fields you change are saved to the note.")
      .addDropdown((d) => {
        for (const option of [...new Set([...choices, current])]) d.addOption(option, option);
        d.setValue(this.status);
        d.onChange((v) => (this.status = v));
      });
    const field = (label, val, set) =>
      new Setting(c).setName(label).addText((t) => t.setValue(val).onChange(set));
    field("Employee", this.employee, (v) => (this.employee = v));
    field("Device model", this.model, (v) => (this.model = v));
    field("Serial number", this.serial, (v) => (this.serial = v));
    const dateField = (label, value, save) =>
      new Setting(c).setName(label).addText((t) => {
        t.inputEl.type = "date";
        t.setValue(value);
        t.onChange(save);
      });
    dateField("Ordered", this.orderDate, (v) => (this.orderDate = v));
    dateField("Set_Up", this.received, (v) => (this.received = v));
    new Setting(c)
      .setName("Old device returned")
      .addToggle((t) => t.setValue(this.oldReturned).onChange((v) => (this.oldReturned = v)));
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Save to original note")
        .setCta()
        .onClick(async () => {
          for (const v of [this.orderDate, this.received])
            if (v && !iso(v)) {
              new Notice("Dates must use YYYY-MM-DD");
              return;
            }
          try {
            const before = this.initial;
            await this.app.fileManager.processFrontMatter(this.record.file, (fm) => {
              // Only write what changed, under the property name the note already uses. Names are
              // listed in the order they are read; new properties use the short name.
              const has = (k) => Object.prototype.hasOwnProperty.call(fm, k);
              const set = (names, value, old, fallback = names[0]) => {
                if (value === old) return;
                const k = names.find(has) ?? fallback;
                if (value === "" && !has(k)) return;
                fm[k] = Array.isArray(fm[k])
                  ? value
                      .split(",")
                      .map((x) => x.trim())
                      .filter(Boolean)
                  : value;
              };
              set(["Status", "status"], this.status, before.status);
              set(["Employee", "Assigned_To"], this.employee, before.employee);
              set(["Device_Model", "Model"], this.model, before.model, "Model");
              set(["Serial_Number", "Serial"], this.serial, before.serial, "Serial");
              set(["Ordered", "Order_Date"], this.orderDate, before.orderDate);
              set(["Set_Up", "Set_up"], this.received, before.received);
              if (this.oldReturned !== before.oldReturned) fm.Old_Device_Returned = this.oldReturned;
              delete fm.Order_Status;
            });
            new Notice("Equipment workflow updated");
            this.close();
            this.plugin.refresh();
          } catch (e) {
            new Notice("Could not save: " + e.message);
          }
        }),
    );
  }
}
class OMEmployeeLicenseModal extends Modal {
  constructor(app, plugin, license, employees) {
    super(app);
    this.plugin = plugin;
    this.license = license;
    this.employees = employees;
    this.search = "";
  }
  onOpen() {
    this.render();
  }
  render() {
    const c = this.contentEl;
    c.empty();
    const field = this.plugin.licenseField(this.license);
    c.createEl("h2", { text: "Employee assignments: " + field });
    new Setting(c).setName("Find employee").addText((t) =>
      t.setValue(this.search).onChange((v) => {
        this.search = v;
        this.render();
      }),
    );
    const matching = this.employees
      .filter((e) => e.file.basename.toLowerCase().includes(this.search.toLowerCase()))
      .sort((a, b) =>
        a.file.basename.localeCompare(b.file.basename, undefined, { sensitivity: "base", numeric: true }),
      );
    for (const e of matching) {
      const active = Object.entries(e.fm).some(
        ([k, v]) => k.toLowerCase() === field.toLowerCase() && omLicenseAssigned(v),
      );
      new Setting(c)
        .setName(e.file.basename)
        .setDesc(active ? "Assigned" : "Not assigned")
        .addButton((b) =>
          b.setButtonText(active ? "Unassign" : "Assign").onClick(async () => {
            try {
              if (!active) {
                const counts = this.plugin.licenseInventory(this.license);
                if (counts.total !== null && counts.assigned >= counts.total) {
                  new Notice("No available seats. Increase purchased seats or unassign an employee first.");
                  return;
                }
              }
              await this.app.fileManager.processFrontMatter(e.file, (fm) => {
                // Reuse the employee's existing checkbox (any capitalisation) instead of adding a second one.
                const key = Object.keys(fm).find((k) => k.toLowerCase() === field.toLowerCase()) ?? field;
                fm[key] = !active;
                e.fm = { ...e.fm, [key]: !active };
              });
              this.plugin.refresh();
              this.render();
            } catch (err) {
              new Notice("Could not update employee: " + err.message);
            }
          }),
        );
    }
    if (!matching.length) c.createEl("p", { text: "No matching employee notes." });
  }
}
function omNextivaTypes(fm) {
  const raw = fm.License_Types;
  return (Array.isArray(raw) ? raw : [])
    .filter((x) => x && typeof x === "object")
    .map((x) => ({
      name: String(x.Name ?? x.name ?? "").trim(),
      purchased: Number(x.Purchased ?? x.purchased),
      rate: Number(x.Monthly_Rate ?? x.monthly_rate),
      field: String(x.Employee_Field ?? x.employee_field ?? "").trim(),
    }))
    .filter(
      (x) =>
        x.name && Number.isInteger(x.purchased) && x.purchased >= 0 && Number.isFinite(x.rate) && x.rate >= 0,
    );
}
// A note whose bill is made up of several license types (License_Types). Notes named after Nextiva
// count even before any types are added.
function omIsBundle(r) {
  return /nextiva/i.test(r.file.basename) || Array.isArray(r.fm.License_Types);
}
function omEmployeeLicenseValue(fm, type) {
  if (type.field) {
    const key = Object.keys(fm).find((k) => k.toLowerCase() === type.field.toLowerCase());
    return key ? omLicenseAssigned(fm[key]) : false;
  }
  const raw = fm.Nextiva_Licenses;
  return (Array.isArray(raw) ? raw : typeof raw === "string" ? [raw] : []).some(
    (v) => String(v).trim().toLowerCase() === type.name.toLowerCase(),
  );
}
function omNextivaSummary(plugin, license) {
  const types = omNextivaTypes(license.fm);
  const people = plugin
    .employeeNotes()
    .filter(
      (e) =>
        omIsEmployee(e) ||
        Array.isArray(e.fm.Nextiva_Licenses) ||
        types.some(
          (t) => t.field && Object.keys(e.fm).some((k) => k.toLowerCase() === t.field.toLowerCase()),
        ),
    );
  const rows = types.map((t) => {
    const holders = people.filter((e) => omEmployeeLicenseValue(e.fm, t));
    const assigned = holders.length;
    return {
      ...t,
      holders,
      assigned,
      available: t.purchased - assigned,
      cost: Math.round(t.purchased * t.rate * 100) / 100,
    };
  });
  return { rows, total: Math.round(rows.reduce((n, x) => n + x.cost, 0) * 100) / 100, people };
}
class OMNextivaTypesModal extends Modal {
  constructor(app, plugin, license) {
    super(app);
    this.plugin = plugin;
    this.license = license;
    this.rows = omNextivaTypes(license.fm).map((x) => ({ ...x }));
  }
  onOpen() {
    this.render();
  }
  render() {
    const c = this.contentEl;
    c.empty();
    c.createEl("h2", { text: "License types: " + this.license.file.basename });
    c.createEl("p", {
      text: "Each price is the monthly rate per purchased license. Charges include unused purchased licenses.",
    });
    for (const [i, row] of this.rows.entries()) {
      const line = c.createDiv({ cls: "ops-project-card" });
      new Setting(line)
        .setName("License type")
        .addText((t) => t.setValue(row.name).onChange((v) => (row.name = v)));
      new Setting(line)
        .setName("Purchased seats")
        .addText((t) => t.setValue(String(row.purchased)).onChange((v) => (row.purchased = v)));
      new Setting(line)
        .setName("Monthly price per seat")
        .addText((t) => t.setValue(String(row.rate)).onChange((v) => (row.rate = v)));
      new Setting(line)
        .setName("Employee frontmatter property")
        .setDesc(
          "Optional. Match assignments using this existing property on employee notes (true or checked). Leave blank to use Nextiva_Licenses.",
        )
        .addText((t) =>
          t
            .setPlaceholder("e.g. Nextiva_Voice")
            .setValue(row.field ?? "")
            .onChange((v) => (row.field = v)),
        );
      new Setting(line).addButton((b) =>
        b.setButtonText("Remove type").onClick(() => {
          this.rows.splice(i, 1);
          this.render();
        }),
      );
    }
    new Setting(c)
      .addButton((b) =>
        b.setButtonText("Add license type").onClick(() => {
          this.rows.push({ name: "", purchased: 0, rate: 0 });
          this.render();
        }),
      )
      .addButton((b) =>
        b
          .setButtonText("Save license types")
          .setCta()
          .onClick(async () => {
            const rows = this.rows.map((x) => ({
              Name: String(x.name).trim(),
              Purchased: Number(x.purchased),
              Monthly_Rate: Number(x.rate),
              ...(String(x.field ?? "").trim() ? { Employee_Field: String(x.field).trim() } : {}),
            }));
            if (
              rows.some(
                (x) =>
                  !x.Name ||
                  !Number.isInteger(x.Purchased) ||
                  x.Purchased < 0 ||
                  !Number.isFinite(x.Monthly_Rate) ||
                  x.Monthly_Rate < 0,
              ) ||
              new Set(rows.map((x) => x.Name.toLowerCase())).size !== rows.length
            ) {
              new Notice("Enter unique names, nonnegative whole seat counts and valid monthly prices");
              return;
            }
            try {
              await this.app.fileManager.processFrontMatter(this.license.file, (fm) => {
                fm.License_Types = rows;
              });
              this.license.fm.License_Types = rows;
              this.close();
              this.plugin.refresh();
              this.plugin.licenseSettingsTab?.display();
              new Notice("License types saved");
            } catch (e) {
              new Notice("Could not save license types: " + e.message);
            }
          }),
      );
  }
}
class OMNextivaAssignmentsModal extends Modal {
  constructor(app, plugin, license) {
    super(app);
    this.plugin = plugin;
    this.license = license;
    this.search = "";
  }
  onOpen() {
    this.render();
  }
  render() {
    const c = this.contentEl;
    c.empty();
    c.createEl("h2", { text: "Employee license assignments" });
    c.createEl("p", {
      text: "Assignments update existing employee-note frontmatter. Purchased seats, not assignments, determine the bill.",
    });
    new Setting(c).setName("Find employee").addText((t) =>
      t.setValue(this.search).onChange((v) => {
        this.search = v;
        this.render();
      }),
    );
    const summary = omNextivaSummary(this.plugin, this.license);
    for (const employee of summary.people
      .filter((e) => e.file.basename.toLowerCase().includes(this.search.toLowerCase()))
      .slice(0, 80)) {
      const details = c.createEl("details");
      details.createEl("summary", { text: employee.file.basename });
      const selected = new Set(
        summary.rows.filter((t) => omEmployeeLicenseValue(employee.fm, t)).map((t) => t.name.toLowerCase()),
      );
      for (const type of summary.rows)
        new Setting(details)
          .setName(type.name)
          .setDesc(type.field ? "Property: " + type.field : "Property: Nextiva_Licenses")
          .addToggle((t) =>
            t.setValue(selected.has(type.name.toLowerCase())).onChange((v) => {
              if (v) selected.add(type.name.toLowerCase());
              else selected.delete(type.name.toLowerCase());
            }),
          );
      new Setting(details).addButton((b) =>
        b.setButtonText("Save assignments").onClick(async () => {
          for (const type of summary.rows) {
            const was = omEmployeeLicenseValue(employee.fm, type),
              now = selected.has(type.name.toLowerCase());
            if (!was && now && type.available <= 0) {
              new Notice("No available " + type.name + " seats. Increase purchased seats first.");
              return;
            }
          }
          try {
            await this.app.fileManager.processFrontMatter(employee.file, (fm) => {
              const current = Array.isArray(fm.Nextiva_Licenses)
                ? fm.Nextiva_Licenses.map(String)
                : typeof fm.Nextiva_Licenses === "string"
                  ? [fm.Nextiva_Licenses]
                  : [];
              const unmapped = new Set(current.map((x) => x.toLowerCase()));
              for (const type of summary.rows) {
                const enabled = selected.has(type.name.toLowerCase());
                if (type.field) {
                  const key = Object.keys(fm).find((k) => k.toLowerCase() === type.field.toLowerCase());
                  if (key) fm[key] = enabled;
                  else if (enabled) fm[type.field] = true;
                } else {
                  if (enabled) unmapped.add(type.name.toLowerCase());
                  else unmapped.delete(type.name.toLowerCase());
                }
              }
              const previouslyNamed = new Map(current.map((x) => [x.toLowerCase(), x]));
              const named = [...unmapped].map(
                (x) =>
                  previouslyNamed.get(x) ?? summary.rows.find((t) => t.name.toLowerCase() === x)?.name ?? x,
              );
              if (named.length) fm.Nextiva_Licenses = named;
              else delete fm.Nextiva_Licenses;
            });
            const updated = this.app.metadataCache.getFileCache(employee.file)?.frontmatter;
            employee.fm = updated ?? employee.fm;
            this.plugin.refresh();
            this.render();
            new Notice("Assignments saved for " + employee.file.basename);
          } catch (error) {
            new Notice("Unable to save assignments: " + error.message);
          }
        }),
      );
    }
  }
}
class OMContractDateModal extends Modal {
  constructor(app, plugin, license) {
    super(app);
    this.plugin = plugin;
    this.license = license;
    this.date = iso(license.fm.Next_Renewal) ?? "";
  }
  onOpen() {
    const c = this.contentEl;
    c.createEl("h2", { text: "Contract renewal: " + this.license.file.basename });
    c.createEl("p", { text: "This date is independent of your billing schedule." });
    new Setting(c)
      .setName("Next contract renewal date")
      .setDesc("YYYY-MM-DD")
      .addText((t) =>
        t
          .setPlaceholder("2027-08-04")
          .setValue(this.date)
          .onChange((value) => (this.date = value.trim())),
      );
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Save renewal date")
        .setCta()
        .onClick(async () => {
          if (this.date && !iso(this.date)) {
            new Notice("Enter a valid YYYY-MM-DD date");
            return;
          }
          try {
            await this.app.fileManager.processFrontMatter(this.license.file, (fm) => {
              if (this.date) fm.Next_Renewal = this.date;
              else delete fm.Next_Renewal;
            });
            this.license.fm.Next_Renewal = this.date;
            this.close();
            this.plugin.refresh();
            this.plugin.licenseSettingsTab?.display();
            new Notice("Contract renewal date saved");
          } catch (error) {
            new Notice("Could not save renewal date: " + error.message);
          }
        }),
    );
  }
}
class OMAlertRulesModal extends Modal {
  constructor(app, plugin, license) {
    super(app);
    this.plugin = plugin;
    this.license = license;
    this.rules = {};
    for (const kind of ["renewal", "billing", "seats"])
      this.rules[kind] = omAlertRules(plugin, license.file, kind).map((x) => ({ ...x }));
  }
  onOpen() {
    this.render();
  }
  render() {
    const c = this.contentEl;
    c.empty();
    c.createEl("h2", { text: "Alert colors & thresholds: " + this.license.file.basename });
    c.createEl("p", {
      cls: "setting-item-description",
      text: "Customize this license only. The first matching threshold decides the color; No alert hides that event. These settings do not add properties to your notes.",
    });
    for (const [kind, title, unit] of [
      ["renewal", "Contract renewal", "days remaining"],
      ["billing", "Upcoming charges", "days remaining"],
      ...(omTracksSeats(this.license.fm) ? [["seats", "Available seats", "seats remaining"]] : []),
    ]) {
      const group = c.createDiv({ cls: "ops-alert-rule-group" });
      group.createEl("h3", { text: title });
      group.createEl("p", {
        cls: "setting-item-description",
        text: "Use ascending thresholds. The last tier sets when alerts begin; choose No alert to suppress a tier.",
      });
      for (const [index, row] of this.rules[kind].entries()) {
        const line = new Setting(group).setName(
          ["Critical", "Warning", "Early notice"][index] + " (" + unit + ")",
        );
        line.addText((t) =>
          t
            .setPlaceholder("Threshold")
            .setValue(String(row.limit))
            .onChange((v) => (row.limit = v)),
        );
        line.addDropdown((d) => {
          for (const color of OM_ALERT_COLORS)
            d.addOption(color, color === "none" ? "No alert" : color[0].toUpperCase() + color.slice(1));
          d.setValue(row.color).onChange((v) => (row.color = v));
        });
      }
    }
    new Setting(c)
      .addButton((b) =>
        b.setButtonText("Restore defaults").onClick(() => {
          for (const kind of ["renewal", "billing", "seats"])
            this.rules[kind] = OM_ALERT_DEFAULTS[kind].map((x) => ({ ...x }));
          this.render();
        }),
      )
      .addButton((b) =>
        b
          .setButtonText("Save alert rules")
          .setCta()
          .onClick(async () => {
            for (const kind of ["renewal", "billing", "seats"]) {
              const rows = this.rules[kind].map((x) => ({ limit: Number(x.limit), color: x.color }));
              if (
                rows.some(
                  (x, i) =>
                    !Number.isInteger(x.limit) ||
                    x.limit < 0 ||
                    x.limit > 3650 ||
                    !OM_ALERT_COLORS.includes(x.color) ||
                    (i > 0 && x.limit < rows[i - 1].limit),
                )
              ) {
                new Notice("Use whole-number thresholds in ascending order from 0 to 3650");
                return;
              }
              this.rules[kind] = rows;
            }
            try {
              const key = this.license.file.path;
              this.plugin.settings.licenseAlertRules = {
                ...(this.plugin.settings.licenseAlertRules || {}),
                [key]: this.rules,
              };
              await this.plugin.saveData(this.plugin.settings);
              this.plugin.refresh();
              this.close();
              new Notice("Alert rules saved for " + this.license.file.basename);
            } catch (e) {
              new Notice("Could not save alert rules: " + e.message);
            }
          }),
      );
  }
}
class OMLicenseFieldModal extends Modal {
  constructor(app, plugin, license) {
    super(app);
    this.plugin = plugin;
    this.license = license;
    this.field = String(license.fm.License_Field ?? license.fm.License_Name ?? license.file.basename).trim();
  }
  onOpen() {
    const c = this.contentEl;
    c.createEl("h2", { text: "Employee field: " + this.license.file.basename });
    c.createEl("p", {
      text: "Use an existing property on employee notes to count assigned seats. The value should be true or checked. No employee notes are changed by saving this mapping.",
    });
    new Setting(c)
      .setName("Employee frontmatter property")
      .addText((t) => t.setValue(this.field).onChange((v) => (this.field = v.trim())));
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Save field mapping")
        .setCta()
        .onClick(async () => {
          if (!this.field) {
            new Notice("Enter a frontmatter property name");
            return;
          }
          try {
            await this.app.fileManager.processFrontMatter(this.license.file, (fm) => {
              fm.License_Field = this.field;
            });
            this.license.fm.License_Field = this.field;
            this.close();
            this.plugin.refresh();
            this.plugin.licenseSettingsTab?.display();
            new Notice("Employee field mapping saved");
          } catch (error) {
            new Notice("Could not save mapping: " + error.message);
          }
        }),
    );
  }
}
class OMSeatCountModal extends Modal {
  constructor(app, plugin, license) {
    super(app);
    this.plugin = plugin;
    this.license = license;
    this.count = String(license.fm.License_Seats ?? "");
  }
  onOpen() {
    const c = this.contentEl;
    c.createEl("h2", { text: "Purchased seats: " + this.plugin.licenseField(this.license) });
    new Setting(c)
      .setName("Number of seats purchased")
      .addText((t) => t.setValue(this.count).onChange((v) => (this.count = v)));
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Save")
        .setCta()
        .onClick(async () => {
          const n = Number(this.count);
          if (!Number.isInteger(n) || n < 0) {
            new Notice("Enter a nonnegative whole number");
            return;
          }
          try {
            await this.app.fileManager.processFrontMatter(this.license.file, (fm) => {
              fm.License_Seats = n;
            });
            this.license.fm.License_Seats = n;
            this.close();
            this.plugin.refresh();
            if (this.plugin.licenseSettingsTab) this.plugin.licenseSettingsTab.display();
            new Notice("Seat count saved");
          } catch (e) {
            new Notice(e.message);
          }
        }),
    );
  }
}
class OMProjectModal extends Modal {
  constructor(app, plugin) {
    super(app);
    this.plugin = plugin;
    this.name = "";
  }
  onOpen() {
    const c = this.contentEl;
    c.createEl("h2", { text: "Create master project" });
    new Setting(c).setName("Project name").addText((t) => t.onChange((v) => (this.name = v)));
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Create project")
        .setCta()
        .onClick(async () => {
          if (!this.name.trim()) {
            new Notice("Project name required");
            return;
          }
          const folder = this.plugin.settings.folder;
          const path = normalizePath(folder + "/" + this.name.trim().replace(/[\\\\/:*?"<>|]/g, "-") + ".md");
          try {
            if (!this.app.vault.getAbstractFileByPath(folder)) await this.app.vault.createFolder(folder);
            if (this.app.vault.getAbstractFileByPath(path)) {
              new Notice("Project already exists");
              return;
            }
            const content =
              "---\\ntags:\\n  - ops-project\\nStatus: planned\\nProject_Owner: " +
              JSON.stringify(this.owner) +
              "\\nTarget_Date: " +
              JSON.stringify(this.target) +
              "\\n---\\n# " +
              this.name +
              "\\n\\n## Scope\\n\\n## Decisions\\n\\n## Notes\\n";
            await this.app.vault.create(path, content);
            new Notice("Project created");
            this.close();
            this.plugin.refresh();
          } catch (e) {
            new Notice(e.message);
          }
        }),
    );
  }
}
class OMLinkModal extends Modal {
  constructor(app, plugin, project) {
    super(app);
    this.plugin = plugin;
    this.project = project;
    this.search = "";
  }
  onOpen() {
    this.render();
  }
  render() {
    const c = this.contentEl;
    c.empty();
    c.createEl("h2", { text: "Link records to " + this.project.basename });
    new Setting(c).setName("Search device, license or order").addText((t) =>
      t.setValue(this.search).onChange((v) => {
        this.search = v;
        this.render();
      }),
    );
    const options = this.plugin
      .records()
      .filter(
        (r) =>
          r.file.path !== this.project.path &&
          ["device", "renewal", "order"].includes(r.kind) &&
          r.file.basename.toLowerCase().includes(this.search.toLowerCase()),
      )
      .slice(0, 40);
    for (const r of options) {
      const line = new Setting(c).setName(r.file.basename).setDesc(r.kind + " • " + r.file.path);
      line.addButton((b) =>
        b.setButtonText(this.plugin.linkedTo(r.fm, this.project) ? "Unlink" : "Link").onClick(async () => {
          try {
            await this.app.fileManager.processFrontMatter(r.file, (fm) => {
              const list = omRelated(fm, "Operations_Projects").map(omLinkValue).filter(Boolean);
              const linked = list.some((x) => omMatches(x, this.project));
              fm.Operations_Projects = linked
                ? list.filter((x) => !omMatches(x, this.project))
                : [...list, this.project.path];
              if (!fm.Operations_Projects.length) delete fm.Operations_Projects;
            });
            new Notice("Project links updated");
            this.plugin.refresh();
            this.render();
          } catch (e) {
            new Notice("Could not update: " + e.message);
          }
        }),
      );
    }
  }
}

const OM_ALERT_DEFAULTS = {
  renewal: [
    { limit: 30, color: "red" },
    { limit: 60, color: "orange" },
    { limit: 90, color: "yellow" },
  ],
  billing: [
    { limit: 30, color: "red" },
    { limit: 60, color: "orange" },
    { limit: 90, color: "yellow" },
  ],
  seats: [
    { limit: 2, color: "red" },
    { limit: 5, color: "orange" },
    { limit: 10, color: "none" },
  ],
};
const OM_ALERT_COLORS = ["red", "orange", "yellow", "green", "none"];
function omAlertRules(plugin, file, kind) {
  const raw = plugin.settings.licenseAlertRules?.[file.path]?.[kind];
  const defaults = OM_ALERT_DEFAULTS[kind];
  if (!Array.isArray(raw) || raw.length !== 3) return defaults;
  const rows = raw.map((v, i) => ({ limit: Number(v?.limit), color: String(v?.color ?? "") }));
  return rows.every(
    (v, i) =>
      Number.isInteger(v.limit) &&
      v.limit >= 0 &&
      v.limit <= 3650 &&
      OM_ALERT_COLORS.includes(v.color) &&
      (i === 0 || v.limit >= rows[i - 1].limit),
  )
    ? rows
    : defaults;
}
function omAlertColor(plugin, file, kind, value) {
  if (value === null || value === undefined || !Number.isFinite(value)) return null;
  const rules = omAlertRules(plugin, file, kind);
  for (const rule of rules) if (value <= rule.limit) return rule.color === "none" ? null : rule.color;
  return null;
}
function omAlertDays(plugin, file, kind) {
  return omAlertRules(plugin, file, kind)[2].limit;
}
function omSeatCount(plugin, r) {
  const types = omNextivaTypes(r.fm);
  return types.length
    ? omNextivaSummary(plugin, r).rows.reduce((sum, v) => sum + v.available, 0)
    : plugin.licenseInventory(r).available;
}
// Alert switches are kept in the plugin settings. A Renewal_Alerts / Billing_Alerts / Seat_Alerts
// property written to a note by older versions is still honoured until the switch is changed.
function omAlertSwitch(fm, file, kind, field) {
  const saved = file ? OM_SETTINGS?.licenseAlerts?.[file.path]?.[kind] : undefined;
  if (typeof saved === "boolean") return saved;
  return fm[field] !== false && String(fm[field] ?? "").toLowerCase() !== "false";
}
function omRenewalAlertsEnabled(fm, file) {
  return omAlertSwitch(fm, file, "renewal", "Renewal_Alerts");
}
function omBillingAlertsEnabled(fm, file) {
  return omAlertSwitch(fm, file, "billing", "Billing_Alerts");
}
function omTracksSeats(fm) {
  const raw = fm.Agreement_Types ?? fm.Agreement_Type ?? fm.License_Type ?? [];
  const kinds = (Array.isArray(raw) ? raw : String(raw).split(","))
    .map((v) => String(v).trim().toLowerCase())
    .filter(Boolean);
  if (omNextivaTypes(fm).length > 0) return true;
  // With a Type, the Type decides: licenses count seats (a bundle once its license types are added);
  // services and contracts don't. Agreement_Types only matters for older notes without a Type.
  const type = omNoteType(fm);
  if (type === "license") return !Array.isArray(fm.License_Types);
  if (type) return false;
  if (kinds.length)
    return kinds.some((v) =>
      ["software seats", "software seat", "seat", "seats", "software", "per-user", "per user"].includes(v),
    );
  return (
    fm.License_Seats !== undefined && fm.License_Seats !== null && String(fm.License_Seats).trim() !== ""
  );
}
function omSeatAlertsEnabled(fm, file) {
  return omAlertSwitch(fm, file, "seats", "Seat_Alerts");
}
function omRenewalEstimate(fm) {
  const raw = fm.Cost;
  const amount =
    typeof raw === "number"
      ? raw
      : Number(
          String(raw ?? "")
            .trim()
            .replace(/[$,]/g, ""),
        );
  if (
    raw === null ||
    raw === undefined ||
    String(raw).trim() === "" ||
    !Number.isFinite(amount) ||
    amount < 0
  )
    return null;
  return { amount: Math.round((amount + Number.EPSILON) * 100) / 100 };
}
function omBillingCycle(fm) {
  const value = String(fm.Billing ?? "")
    .trim()
    .toLowerCase();
  const months = {
    monthly: 1,
    month: 1,
    quarterly: 3,
    quarter: 3,
    annually: 12,
    annual: 12,
    yearly: 12,
    year: 12,
    semiannually: 6,
    "semi-annually": 6,
    semiannual: 6,
    biannually: 6,
    biannual: 6,
  };
  // One-off services only charge when used, so they have no billing schedule or charge alerts.
  if (/^(one[ -]?off|one[ -]?time|as[ -]needed|per[ -]use|on[ -]demand)$/.test(value))
    return { months: 0, label: "One-off", oneOff: true };
  return Object.prototype.hasOwnProperty.call(months, value)
    ? { months: months[value], label: value.charAt(0).toUpperCase() + value.slice(1) }
    : null;
}
function omBillingInfo(fm) {
  const cycle = omBillingCycle(fm);
  const date = iso(fm.Next_Billing);
  if (!cycle || cycle.oneOff || !date) return null;
  return {
    date,
    days: daysUntil(date),
    months: cycle.months,
    label: cycle.label,
    amount: omRenewalEstimate(fm)?.amount ?? null,
  };
}
function omTermCost(fm) {
  const charge = omRenewalEstimate(fm),
    cycle = omBillingCycle(fm),
    term = Number(fm.Renewal);
  if (!charge || !cycle || cycle.oneOff || !Number.isInteger(term) || term <= 0 || term > 36500) return null;
  const periods = Math.max(1, Math.round((term * 12) / (365.25 * cycle.months)));
  return { total: Math.round(charge.amount * periods * 100) / 100, periods };
}
function omUSD(n) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(n);
}
function omDashboardAction(r, plugin) {
  const fm = r.fm;
  if (r.equipUp) {
    const phase = String(fm.Approved ?? fm.approved ?? "Unassigned phase");
    const ordered = iso(fm.Ordered ?? fm.Order_Date),
      setup = iso(fm.Set_Up ?? fm.Set_up);
    const label = !ordered
      ? "Confirm approval and order equipment"
      : !setup
        ? "Follow up on delivery and schedule Set_Up"
        : "Confirm completion and update Status";
    return {
      priority: !ordered ? 20 : !setup ? 40 : 65,
      date: ordered ?? "",
      label: label + " · Phase: " + phase,
      kind: "equipment",
    };
  }
  if (r.kind === "invoice") {
    const due = iso(fm.Due_Date),
      created = iso(fm.Date),
      age = created ? -daysUntil(created) : 0;
    return {
      priority: due && daysUntil(due) < 0 ? 5 : age > 7 ? 15 : due && daysUntil(due) <= 7 ? 25 : 50,
      date: due ?? created ?? "",
      label:
        due && daysUntil(due) < 0
          ? "Check overdue payment · Due " + due
          : age > 7
            ? "Follow up on unpaid invoice · Dated " + created
            : due
              ? "Check payment · Due " + due
              : "Check payment · No due date",
      kind: "invoice",
    };
  }
  if (r.kind === "renewal") {
    const renewal = omMonthToMonth(fm, r.file) ? null : inferredRenewal(fm);
    const renewalDays = renewal ? daysUntil(renewal.date) : null;
    const renewalDue =
      omRenewalAlertsEnabled(fm, r.file) && !!omAlertColor(plugin, r.file, "renewal", renewalDays);
    const billing = omBillingAlertsEnabled(fm, r.file) ? omBillingInfo(fm) : null;
    const billingDue = billing && !!omAlertColor(plugin, r.file, "billing", billing.days);
    const urgency = (d) => (d < 0 ? 0 : d <= 14 ? 10 : d <= 30 ? 30 : 45);
    const preferBilling = billingDue && (!renewalDue || billing.date < renewal.date);
    const seatDue =
      omTracksSeats(fm) &&
      omSeatAlertsEnabled(fm, r.file) &&
      !!omAlertColor(plugin, r.file, "seats", omSeatCount(plugin, r));
    const label = preferBilling
      ? "Upcoming charge · " +
        billing.date +
        (billing.days < 0
          ? " · Overdue " + -billing.days + " days"
          : " · " + billing.days + " days remaining")
      : renewalDue
        ? "Renew license or agreement package · " +
          renewal.date +
          (renewalDays < 0 ? " · Overdue " + -renewalDays + " days" : " · " + renewalDays + " days remaining")
        : seatDue
          ? "Review license availability · " + omSeatCount(plugin, r) + " seats remaining"
          : "Review license record";
    return {
      priority: Math.min(
        renewalDue ? urgency(renewalDays) : Infinity,
        billingDue ? urgency(billing.days) : Infinity,
        55,
      ),
      date: preferBilling ? billing.date : renewalDue ? renewal.date : "",
      label,
      kind: "renewal",
    };
  }
  return { priority: 70, date: "", label: r.reason || "Follow up", kind: r.kind };
}
function omLicenseActiveAlerts(plugin, r) {
  if (r.kind !== "renewal") return [];
  const fm = r.fm,
    alerts = [];
  const renewal = omMonthToMonth(fm, r.file) ? null : inferredRenewal(fm);
  if (renewal && omRenewalAlertsEnabled(fm, r.file)) {
    const days = daysUntil(renewal.date),
      color = omAlertColor(plugin, r.file, "renewal", days);
    if (color)
      alerts.push({
        kind: "renewal",
        color,
        date: renewal.date,
        text:
          "Contract renewal: " +
          renewal.date +
          " · " +
          (days < 0 ? "Overdue " + -days + " days" : days + " days remaining"),
      });
  }
  const billing = omBillingAlertsEnabled(fm, r.file) ? omBillingInfo(fm) : null;
  if (billing) {
    const color = omAlertColor(plugin, r.file, "billing", billing.days);
    if (color)
      alerts.push({
        kind: "billing",
        color,
        date: billing.date,
        text:
          "Upcoming charge: " +
          billing.date +
          " · " +
          (billing.days < 0 ? "Overdue " + -billing.days + " days" : billing.days + " days remaining") +
          (billing.amount === null ? "" : " · " + omUSD(billing.amount)),
      });
  }
  if (omTracksSeats(fm) && omSeatAlertsEnabled(fm, r.file)) {
    const available = omSeatCount(plugin, r),
      color = omAlertColor(plugin, r.file, "seats", available);
    if (color) {
      const types = omNextivaTypes(fm);
      const purchased = types.length
        ? types.reduce((sum, t) => sum + t.purchased, 0)
        : plugin.licenseInventory(r).total;
      alerts.push({
        kind: "seats",
        color,
        date: null,
        text:
          "Seat availability: " +
          available +
          " of " +
          (purchased === null ? "unknown" : purchased) +
          " seats available" +
          (available <= 0 ? " · Review assignments or purchased seats" : ""),
      });
    }
  }
  return alerts;
}
function omDashboardDisplayColor(plugin, r) {
  if (r.kind === "renewal") {
    const colors = omLicenseActiveAlerts(plugin, r).map((a) => a.color);
    return ["red", "orange", "yellow", "green"].find((c) => colors.includes(c)) ?? null;
  }
  const due = r.kind === "invoice" ? iso(r.fm.Due_Date) : null;
  const days = due ? daysUntil(due) : null;
  const stale = r.kind === "invoice" && r.attention && iso(r.fm.Date) && daysUntil(iso(r.fm.Date)) < -7;
  return stale || (days !== null && days <= 30) ? "red" : days !== null && days <= 60 ? "orange" : null;
}
function omDashboardSortDate(plugin, r) {
  if (r.kind === "renewal")
    return (
      omLicenseActiveAlerts(plugin, r)
        .filter((a) => a.date)
        .map((a) => a.date)
        .sort()[0] ?? "9999-12-31"
    );
  return iso(r.fm.Due_Date) ?? iso(r.fm.Date) ?? "9999-12-31";
}
class OperationsView extends ItemView {
  constructor(leaf, plugin) {
    super(leaf);
    this.plugin = plugin;
  }
  getViewType() {
    return VIEW;
  }
  getDisplayText() {
    return "Operations Manager";
  }
  getIcon() {
    return "layout-dashboard";
  }
  async onOpen() {
    await this.render();
  }
  renderInvoices(root, records) {
    let sort = "date";
    const controls = root.createDiv({ cls: "ops-equipment-controls" });
    new Setting(controls).setName("Sort invoices by").addDropdown((d) =>
      d
        .addOption("date", "Date (newest first)")
        .addOption("vendor", "Vendor (A-Z)")
        .setValue(sort)
        .onChange((v) => {
          sort = v;
          draw();
        }),
    );
    const area = root.createDiv({ cls: "ops-invoice-groups" });
    const draw = () => {
      area.empty();
      const ordered = [...records].sort((a, b) =>
        sort === "vendor"
          ? String(a.fm.Vendor ?? "").localeCompare(String(b.fm.Vendor ?? ""), undefined, {
              numeric: true,
              sensitivity: "base",
            }) || a.file.basename.localeCompare(b.file.basename)
          : String(b.fm.Date ?? "").localeCompare(String(a.fm.Date ?? "")) ||
            a.file.basename.localeCompare(b.file.basename),
      );
      for (const [label, paid] of [
        ["Unpaid", false],
        ["Paid", true],
      ]) {
        const items = ordered.filter(
          (r) =>
            (status(r.fm) === "complete" ||
              status(r.fm) === "paid" ||
              String(r.fm.Paid_Date ?? "").trim() !== "") === paid,
        );
        const key = "invoices:status:" + (paid ? "paid" : "unpaid");
        const section = area.createEl("details", { cls: "ops-equipment-phase ops-invoice-section" });
        section.open = this.plugin.settings.collapsedSections?.[key] === true;
        section.addEventListener("toggle", () => {
          if (this.plugin.settings.collapsedSections?.[key] === section.open) return;
          this.plugin.settings.collapsedSections = {
            ...(this.plugin.settings.collapsedSections || {}),
            [key]: section.open,
          };
          this.plugin.saveData(this.plugin.settings);
        });
        section.createEl("summary", { text: label + " (" + items.length + ")" });
        const list = section.createDiv({ cls: "ops-equipment-list" });
        if (!items.length) {
          list.createEl("small", { text: "No " + label.toLowerCase() + " invoices." });
          continue;
        }
        for (const r of items) {
          const card = list.createDiv({ cls: "ops-attention-item ops-equipment-card" });
          this.plugin.renderCardIcon(card, "invoice");
          const info = card.createDiv({ cls: "ops-item-info" });
          const title = info.createEl("button", { cls: "ops-item-title", text: r.file.basename });
          title.onclick = () => this.app.workspace.getLeaf(false).openFile(r.file);
          info.createEl("small", {
            text: "Date: " + (r.fm.Date ?? "Not set") + " · Vendor: " + (r.fm.Vendor ?? "Not set"),
          });
          info.createEl("small", { text: r.reason || "" });
          if (r.fm.Notes !== undefined && r.fm.Notes !== null && String(r.fm.Notes).trim()) {
            const notes = info.createDiv({ cls: "ops-equipment-note" });
            notes.createSpan({ text: "Services: " });
            omCardLinkedText(notes, r.fm.Notes, this.app, r.file.path);
          }
          omRenderRelated(info, r.fm, this.app, r.file.path);
          if (!paid) {
            const btn = card.createEl("button", { cls: "ops-item-action", text: "Mark paid" });
            btn.onclick = () => new PaidModal(this.app, this.plugin, r.file).open();
          }
        }
      }
    };
    draw();
  }

  renderLicenseOverviewInline(info, r) {
    const plugin = this.plugin;
    const types = omNextivaTypes(r.fm);
    const multi = types.length > 0;
    const inv = plugin.licenseInventory(r);
    const data = multi ? omNextivaSummary(plugin, r) : null;
    const purchased = multi ? data.rows.reduce((n, x) => n + x.purchased, 0) : inv.total;
    const assigned = multi ? data.rows.reduce((n, x) => n + x.assigned, 0) : inv.assigned;
    const available = purchased === null ? null : purchased - assigned;
    const amount = omRenewalEstimate(r.fm)?.amount ?? null;
    const period = omBillingCycle(r.fm)?.label?.toLowerCase() ?? "billing cycle";
    const seatBased = omTracksSeats(r.fm);
    const line = seatBased
      ? purchased === null
        ? "Purchased seats not set"
        : purchased + " purchased · " + assigned + " assigned · " + available + " available"
      : "";
    const costLine = amount === null ? "" : omUSD(amount) + " / " + period;
    if (line || costLine)
      info.createEl("small", {
        cls: "ops-license-summary",
        text: [line, costLine].filter(Boolean).join(" · "),
      });
    if (!seatBased && !omIsBundle(r)) return;
    const key = "licenses:inline:" + r.file.path;
    const details = info.createEl("details", { cls: "ops-license-inline-details" });
    details.open = plugin.settings.collapsedSections?.[key] === true;
    details.addEventListener("toggle", () => {
      if (plugin.settings.collapsedSections?.[key] === details.open) return;
      plugin.settings.collapsedSections = {
        ...(plugin.settings.collapsedSections || {}),
        [key]: details.open,
      };
      plugin.saveData(plugin.settings);
    });
    details.createEl("summary", { text: "License details" });
    const content = details.createDiv({ cls: "ops-license-inline-body" });
    if (omIsBundle(r) && !multi) {
      content.createEl("p", { text: "No license types configured yet." });
      const add = content.createEl("button", { text: "Add license types" });
      add.onclick = () => new OMNextivaTypesModal(this.app, plugin, r).open();
    }
    if (multi) {
      content.createEl("p", {
        text:
          "Calculated license-type subtotal: " +
          omUSD(data.total) +
          " (comparison only; billed amount comes from Cost).",
      });
      const table = content.createEl("table");
      const header = table.createEl("tr");
      for (const title of ["Type", "Purchased", "Assigned", "Available", "Rate", "Calculated subtotal"])
        header.createEl("th", { text: title });
      for (const type of data.rows) {
        const tr = table.createEl("tr");
        for (const value of [
          type.name,
          type.purchased,
          type.assigned,
          type.available,
          omUSD(type.rate),
          omUSD(type.cost),
        ])
          tr.createEl("td", { text: String(value) });
        const people = content.createDiv({ cls: "ops-license-assignees" });
        people.createSpan({ text: type.name + ": " });
        if (!type.holders.length) people.createSpan({ text: "No employees assigned" });
        for (const [i, employee] of type.holders.entries()) {
          if (i) people.createSpan({ text: ", " });
          const link = people.createEl("button", { text: employee.file.basename });
          link.onclick = () => this.app.workspace.getLeaf(false).openFile(employee.file);
        }
      }
      const edit = content.createEl("button", { text: "Manage license types" });
      edit.onclick = () => new OMNextivaTypesModal(this.app, plugin, r).open();
      const assign = content.createEl("button", { text: "Manage assignments" });
      assign.onclick = () => new OMNextivaAssignmentsModal(this.app, plugin, r).open();
    } else if (seatBased) {
      if (amount !== null && purchased !== null && purchased > 0)
        content.createEl("p", {
          text: "Derived average per purchased seat: " + omUSD(amount / purchased) + " / " + period,
        });
      const employees = inv.employees.filter((e) =>
        Object.entries(e.fm).some(
          ([field, value]) => field.toLowerCase() === inv.field.toLowerCase() && omLicenseAssigned(value),
        ),
      );
      content.createEl("p", { text: "Assigned employees: " + employees.length });
      for (const employee of employees) {
        const button = content.createEl("button", { text: employee.file.basename });
        button.onclick = () => this.app.workspace.getLeaf(false).openFile(employee.file);
      }
      const assign = content.createEl("button", { text: "Manage assignments" });
      assign.onclick = () => new OMEmployeeLicenseModal(this.app, plugin, r, inv.employees).open();
    }
  }

  renderOrders(root, records) {
    root.createEl("h2", { cls: "ops-equipment-heading", text: "Orders by Status" });
    const groups = new Map();
    for (const r of records) {
      const raw = String(r.fm.Status ?? r.fm.status ?? "").trim();
      const name = raw || (r.fm.Ordered ? "ordered" : "not ordered");
      const label = name.toLowerCase() === "in progress" && !r.fm.Ordered ? "not ordered" : name;
      if (!groups.has(label)) groups.set(label, []);
      groups.get(label).push(r);
    }
    const rank = (name) => {
      const value = name.trim().toLowerCase();
      return value === "not ordered"
        ? 0
        : value === "in progress"
          ? 1
          : value === "ordered"
            ? 2
            : value === "complete"
              ? 3
              : 4;
    };
    const names = [...groups.keys()].sort(
      (a, b) => rank(a) - rank(b) || a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }),
    );
    if (!names.length) root.createEl("p", { text: "No orders found." });
    for (const name of names) {
      const items = groups
        .get(name)
        .sort((a, b) =>
          a.file.basename.localeCompare(b.file.basename, undefined, { numeric: true, sensitivity: "base" }),
        );
      const key = "orders:status:" + name.toLowerCase();
      const details = root.createEl("details", { cls: "ops-equipment-phase" });
      details.open = this.plugin.settings.collapsedSections?.[key] === true;
      details.addEventListener("toggle", () => {
        if (this.plugin.settings.collapsedSections?.[key] === details.open) return;
        this.plugin.settings.collapsedSections = {
          ...(this.plugin.settings.collapsedSections || {}),
          [key]: details.open,
        };
        this.plugin.saveData(this.plugin.settings);
      });
      details.createEl("summary", {
        text: name.charAt(0).toUpperCase() + name.slice(1) + " (" + items.length + ")",
      });
      const list = details.createDiv({ cls: "ops-equipment-list" });
      for (const r of items) {
        const card = list.createDiv({ cls: "ops-attention-item ops-equipment-card" });
        this.plugin.renderCardIcon(card, "order");
        const info = card.createDiv({ cls: "ops-item-info" });
        omCardTitle(info, r, this.app);
        const emp = r.fm.Employee ?? r.fm.Assigned_To;
        if (emp) {
          const line = info.createEl("small");
          omCardLinkedText(line, emp, this.app, r.file.path);
        }
        const approval = r.fm.Approved;
        if (approval !== undefined && approval !== null && String(approval).trim()) {
          const line = info.createDiv({ cls: "ops-equipment-note" });
          line.createSpan({ text: "Approved: " });
          omCardLinkedText(line, approval, this.app, r.file.path);
        }
        const notes = r.fm.Notes;
        if (notes !== undefined && notes !== null && String(notes).trim()) {
          const line = info.createDiv({ cls: "ops-equipment-note" });
          line.createSpan({ text: "Notes: " });
          omCardLinkedText(line, notes, this.app, r.file.path);
        }
        omRenderRelated(info, r.fm, this.app, r.file.path);
        if (tagSet(r.fm).has("equip_ug") || tagSet(r.fm).has("equip_upgrade")) {
          const update = card.createEl("button", { cls: "ops-item-action", text: "Update" });
          update.onclick = () => new OMDeviceWorkflowModal(this.app, this.plugin, r).open();
        }
      }
    }
  }

  async render() {
    const root = this.contentEl;
    const scroll = root.querySelector(".view-content") || root;
    const previousScroll = scroll.scrollTop;
    root.empty();
    root.addClass("ops-command");
    const rows = this.plugin.records();
    const nav = root.createDiv({ cls: "ops-header" });
    const actions = nav.createDiv({ cls: "ops-actions" });
    for (const [name, fn] of [
      ["Refresh", () => this.render()],
      ["New note", () => new OMNewNoteModal(this.app, this.plugin).open()],
      ["Export Excel (.xlsx)", () => new OMExportExcelModal(this.app, this.plugin).open()],
    ]) {
      const btn = actions.createEl("button", { text: name });
      btn.onclick = fn;
    }
    const tabs = root.createDiv({ cls: "ops-tabs" });
    const pages = [
      ["Dashboard", "dashboard"],
      ["Equipment", "device"],
      ["Licenses & Renewals", "renewal"],
      ["Invoices", "invoice"],
      ["Orders", "order"],
    ];
    this.page = this.page || "dashboard";
    for (const [label, key] of pages) {
      const t = tabs.createEl("button", {
        text: label,
        cls: key === this.page ? "ops-tab active" : "ops-tab",
      });
      t.onclick = () => {
        this.page = key;
        this.render();
      };
    }
    const due = rows.filter((r) => r.kind === "renewal" && r.attention).length;
    const orders = rows.filter((r) => r.kind === "order" && !r.done).length;
    const invoices = rows.filter((r) => r.kind === "invoice" && !r.done).length;
    const stats = root.createDiv({ cls: "ops-stats" });
    for (const [num, label, key] of [
      [due, "Renewals / charges due", "renewal"],
      [orders, "Open orders", "order"],
      [invoices, "Unpaid invoices", "invoice"],
    ]) {
      const stat = stats.createDiv({ cls: "ops-stat" });
      const statIcon = stat.createSpan({ cls: "ops-stat-icon" });
      const iconName = this.plugin.cardIcon(key);
      if (iconName !== "none") require("obsidian").setIcon(statIcon, iconName);
      stat.createEl("strong", { text: String(num) });
      stat.createEl("small", { text: label });
      stat.onclick = () => {
        this.page = key;
        this.render();
      };
    }
    const selected =
      this.page === "dashboard"
        ? rows
            .filter(
              (r) =>
                r.kind !== "project" &&
                r.attention &&
                !omTerminated(r.fm) &&
                ((r.kind !== "order" && r.kind !== "device") || r.equipUp),
            )
            .sort((a, b) => {
              const x = omDashboardAction(a, this.plugin),
                y = omDashboardAction(b, this.plugin);
              return (
                x.priority - y.priority ||
                (x.date || "9999").localeCompare(y.date || "9999") ||
                a.file.basename.localeCompare(b.file.basename)
              );
            })
        : rows.filter(
            (r) =>
              !omTerminated(r.fm) &&
              (this.page === "order"
                ? r.kind === "order"
                : this.page === "device"
                  ? r.kind === "device" || (r.kind === "order" && omIsDeviceNote(r.fm))
                  : r.kind === this.page),
          );
    if (this.page === "invoice") {
      this.renderInvoices(root, selected);
      return;
    }
    if (this.page === "order") {
      this.renderOrders(root, selected);
      return;
    }
    if (this.page === "device") {
      this.plugin.renderEquipmentPhases(
        root,
        selected.filter((r) => r.kind !== "invoice"),
        "status",
      );
      return;
    }
    if (this.page === "renewal") {
      const sort = this.licenseSort || "filename";
      new Setting(root.createDiv({ cls: "ops-equipment-controls" }))
        .setName("Sort licenses by")
        .addDropdown((d) =>
          d
            .addOption("filename", "Filename (A-Z)")
            .addOption("seats", "Seats left (fewest first)")
            .addOption("renewal", "Next renewal date (soonest first)")
            .setValue(sort)
            .onChange((v) => {
              this.licenseSort = v;
              this.render();
            }),
        );
      const compareName = (a, b) =>
        a.file.basename.localeCompare(b.file.basename, undefined, { numeric: true, sensitivity: "base" });
      selected.sort((a, b) => {
        if (sort === "seats") {
          const av = this.plugin.licenseInventory(a).available,
            bv = this.plugin.licenseInventory(b).available;
          return (av === null ? Infinity : av) - (bv === null ? Infinity : bv) || compareName(a, b);
        }
        if (sort === "renewal") {
          const ad = omMonthToMonth(a.fm, a.file) ? null : inferredRenewal(a.fm)?.date,
            bd = omMonthToMonth(b.fm, b.file) ? null : inferredRenewal(b.fm)?.date;
          return (ad || "9999-12-31").localeCompare(bd || "9999-12-31") || compareName(a, b);
        }
        return compareName(a, b);
      });
    }
    if (this.page === "renewal")
      root.createEl("h2", { cls: "ops-equipment-heading", text: "Licenses & Renewals" });
    else if (this.page !== "dashboard")
      root.createEl("h3", { text: pages.find((p) => p[1] === this.page)?.[0] || "Records" });
    const list = root.createDiv({ cls: "ops-attention" });
    const actionRows = this.page === "dashboard" ? selected.filter((r) => !r.equipUp) : selected;
    const dashboard = this.page === "dashboard";
    const seatAlerts = dashboard
      ? rows.filter(
          (r) =>
            r.kind === "renewal" &&
            omTracksSeats(r.fm) &&
            omSeatAlertsEnabled(r.fm, r.file) &&
            !!omAlertColor(this.plugin, r.file, "seats", omSeatCount(this.plugin, r)),
        )
      : [];
    const grouped = new Map();
    for (const r of actionRows) {
      grouped.set(r.file.path, { r, seat: null });
    }
    for (const r of seatAlerts) {
      const inv = this.plugin.licenseInventory(r);
      const types = omNextivaTypes(r.fm);
      const seat = types.length
        ? {
            ...inv,
            available: omSeatCount(this.plugin, r),
            assigned: omNextivaSummary(this.plugin, r).rows.reduce((sum, x) => sum + x.assigned, 0),
            total: types.reduce((sum, x) => sum + x.purchased, 0),
          }
        : inv;
      if (grouped.has(r.file.path)) grouped.get(r.file.path).seat = seat;
      else grouped.set(r.file.path, { r, seat });
    }
    const entries = [...grouped.values()].filter(
      (e) =>
        !dashboard ||
        (!["follow up", "follow-up"].includes(status(e.r.fm)) &&
          !tagSet(e.r.fm).has("follow_up") &&
          !tagSet(e.r.fm).has("follow-up")),
    );
    if (dashboard)
      entries.sort((a, b) => {
        const order = { red: 0, orange: 1, yellow: 2, green: 3 };
        const ac = omDashboardDisplayColor(this.plugin, a.r),
          bc = omDashboardDisplayColor(this.plugin, b.r);
        const rankA = ac === null ? 4 : order[ac],
          rankB = bc === null ? 4 : order[bc];
        return (
          rankA - rankB ||
          omDashboardSortDate(this.plugin, a.r).localeCompare(omDashboardSortDate(this.plugin, b.r)) ||
          a.r.file.basename.localeCompare(b.r.file.basename, undefined, {
            numeric: true,
            sensitivity: "base",
          })
        );
      });
    if (
      !entries.length &&
      !(
        dashboard &&
        this.app.vault.getMarkdownFiles().some((file) => {
          const fm = this.app.metadataCache.getFileCache(file)?.frontmatter;
          return (
            fm &&
            !tagSet(fm).has("archive") &&
            !tagSet(fm).has("meeting_notes") &&
            !tagSet(fm).has("call_notes") &&
            (["follow up", "follow-up"].includes(status(fm)) ||
              tagSet(fm).has("follow_up") ||
              tagSet(fm).has("follow-up"))
          );
        })
      ) &&
      !(
        dashboard &&
        this.app.vault.getMarkdownFiles().some((file) => {
          const fm = this.app.metadataCache.getFileCache(file)?.frontmatter;
          return (
            fm &&
            !tagSet(fm).has("archive") &&
            !tagSet(fm).has("meeting_notes") &&
            !tagSet(fm).has("call_notes") &&
            tagSet(fm).has("equip_ug") &&
            status(fm) !== "complete"
          );
        })
      )
    )
      list.createEl("p", { text: "Nothing currently needs follow-up." });
    for (const { r, seat } of entries) {
      const color = dashboard ? omDashboardDisplayColor(this.plugin, r) : null;
      const item = list.createDiv({ cls: "ops-attention-item" + (color ? " ops-alert-" + color : "") });
      const icon = this.plugin.renderCardIcon(item, r.kind);
      const info = item.createDiv({ cls: "ops-item-info" });
      omCardTitle(info, r, this.app);
      if (
        dashboard &&
        r.kind !== "renewal" &&
        r.kind !== "invoice" &&
        (r.kind === "device" || omIsDeviceNote(r.fm))
      )
        omDashboardDeviceDetails(info, r, this.app, false);
      else if (dashboard && r.kind === "renewal") {
        const active = omLicenseActiveAlerts(this.plugin, r);
        if (active.length) {
          for (const alert of active)
            info.createEl("small", {
              cls: "ops-active-alert ops-active-alert-" + alert.color,
              text: alert.text,
            });
        } else if (
          r.attention &&
          omRenewalAlertsEnabled(r.fm, r.file) &&
          omNeedsRenewalDate(r.fm) &&
          !omMonthToMonth(r.fm, r.file) &&
          !inferredRenewal(r.fm)
        )
          info.createEl("small", { text: "Action needed: Set contract renewal date" });
      } else if (!dashboard || r.attention)
        info.createEl("small", {
          text: dashboard ? omDashboardAction(r, this.plugin).label : r.reason || r.kind,
        });
      if (this.page === "renewal" && r.kind === "renewal") {
        try {
          this.renderLicenseOverviewInline(info, r);
        } catch (error) {
          console.error("Operations Manager: license details failed for " + r.file.path, error);
          info.createEl("small", {
            text: "Could not display license details: " + (error?.message || String(error)),
          });
        }
        omRenderRelated(info, r.fm, this.app, r.file.path);
      }
      if (!dashboard)
        item.createSpan({
          cls: "ops-item-status",
          text: r.attention ? "Action needed" : r.done ? "Complete" : "Tracked",
        });
      if (
        dashboard &&
        r.kind === "renewal" &&
        omBillingAlertsEnabled(r.fm, r.file) &&
        omBillingInfo(r.fm) &&
        !!omAlertColor(this.plugin, r.file, "billing", omBillingInfo(r.fm).days)
      ) {
        const charge = item.createEl("button", { cls: "ops-item-action", text: "Record charge" });
        charge.onclick = () => new OMRecordChargeModal(this.app, this.plugin, r).open();
      }
      if (
        r.kind === "renewal" &&
        omRenewalAlertsEnabled(r.fm, r.file) &&
        inferredRenewal(r.fm) &&
        !omMonthToMonth(r.fm, r.file) &&
        !!omAlertColor(this.plugin, r.file, "renewal", daysUntil(inferredRenewal(r.fm).date))
      ) {
        const btn = item.createEl("button", { cls: "ops-item-action", text: "Renew" });
        btn.onclick = () => new RenewModal(this.app, this.plugin, r.file, r.fm).open();
      }
      if (dashboard && (r.kind === "device" || omIsDeviceNote(r.fm))) {
        const btn = item.createEl("button", { cls: "ops-item-action", text: "Update" });
        btn.onclick = () => new OMDeviceWorkflowModal(this.app, this.plugin, r).open();
      }
      if (r.kind === "invoice" && r.attention) {
        const btn = item.createEl("button", { cls: "ops-item-action", text: "Mark paid" });
        btn.onclick = () => new PaidModal(this.app, this.plugin, r.file).open();
      }
    }
    if (this.page === "dashboard") {
      const existing = new Set(entries.map((e) => e.r.file.path));
      const followups = this.app.vault
        .getMarkdownFiles()
        .filter(
          (file) =>
            !file.path
              .split("/")
              .slice(0, -1)
              .some((p) => p.toLowerCase() === "templates"),
        )
        .map((file) => ({ file, fm: this.app.metadataCache.getFileCache(file)?.frontmatter }))
        .filter(
          (x) =>
            x.fm &&
            !omTerminated(x.fm) &&
            !["archive", "archived"].includes(status(x.fm)) &&
            !["archive", "meeting_notes", "call_notes"].some((t) => tagSet(x.fm).has(t)) &&
            (status(x.fm) === "follow up" ||
              status(x.fm) === "follow-up" ||
              tagSet(x.fm).has("follow_up") ||
              tagSet(x.fm).has("follow-up")) &&
            !existing.has(x.file.path),
        )
        .sort((a, b) =>
          a.file.basename.localeCompare(b.file.basename, undefined, { numeric: true, sensitivity: "base" }),
        );
      for (const r of followups) {
        const card = list.createDiv({ cls: "ops-attention-item" });
        this.plugin.renderCardIcon(card, "followup");
        const info = card.createDiv({ cls: "ops-item-info" });
        const btn = info.createEl("button", { cls: "ops-item-title", text: r.file.basename });
        btn.onclick = () => this.app.workspace.getLeaf(false).openFile(r.file);
        omDashboardDeviceDetails(info, r, this.app, true);
        if (omIsDeviceNote(r.fm)) {
          const update = card.createEl("button", { text: "Update", cls: "ops-item-action" });
          update.onclick = () =>
            new OMDeviceWorkflowModal(this.app, this.plugin, { ...r, kind: "device" }).open();
        }
      }
    }
    if (this.page === "dashboard") {
      const upgrades = this.app.vault
        .getMarkdownFiles()
        .filter(
          (file) =>
            !file.path
              .split("/")
              .slice(0, -1)
              .some((part) => part.toLowerCase() === "templates"),
        )
        .map((file) => ({ file, fm: this.app.metadataCache.getFileCache(file)?.frontmatter }))
        .filter(
          (x) =>
            x.fm &&
            !tagSet(x.fm).has("archive") &&
            !tagSet(x.fm).has("meeting_notes") &&
            !tagSet(x.fm).has("call_notes") &&
            tagSet(x.fm).has("equip_ug") &&
            status(x.fm) !== "complete",
        )
        .map((x) => ({ ...x, kind: "order", equipUp: true, attention: true }));
      if (upgrades.length) this.plugin.renderEquipmentPhases(root, upgrades);
    }
    root.createEl("p", {
      cls: "ops-footnote",
      text: "Data from your existing Markdown notes. Export creates a formatted Excel workbook. No source notes are modified by browsing or exporting.",
    });
    scroll.scrollTop = previousScroll;
  }
}
class OMExportExcelModal extends Modal {
  constructor(app, plugin) {
    super(app);
    this.plugin = plugin;
    this.selection = "tracked";
    this.allProperties = false;
  }
  onOpen() {
    this.render();
  }
  render() {
    const c = this.contentEl;
    c.empty();
    c.createEl("h2", { text: "Export to Excel" });
    c.createEl("p", {
      cls: "setting-item-description",
      text: "Choose which notes to export. Tag exports match YAML frontmatter tags only, even for notes outside the Operations Manager tabs. Inline tags in the note body are ignored.",
    });
    const tags = this.plugin.exportTagNames();
    new Setting(c)
      .setName("What to export")
      .setDesc("Filter only by YAML frontmatter tags, such as #device, #license and #person.")
      .addDropdown((d) => {
        d.addOption("tracked", "Operations Manager records (current export)");
        d.addOption("all", "All Markdown notes");
        for (const tag of tags) d.addOption("tag:" + tag, "#" + tag);
        d.setValue(this.selection).onChange((value) => {
          this.selection = value;
          this.allProperties = value !== "tracked";
          this.render();
        });
      });
    new Setting(c)
      .setName("Include every frontmatter property")
      .setDesc(
        "Creates a column for each property found across the selected notes, including custom fields. Nested values are kept as readable text.",
      )
      .addToggle((t) =>
        t.setValue(this.allProperties).onChange((value) => {
          this.allProperties = value;
          this.updateCount();
        }),
      );
    this.countEl = c.createEl("p", { cls: "ops-footnote" });
    this.updateCount();
    new Setting(c)
      .addButton((b) => b.setButtonText("Cancel").onClick(() => this.close()))
      .addButton((b) =>
        b
          .setButtonText("Export .xlsx")
          .setCta()
          .onClick(async () => {
            const notes = this.plugin.exportSourceNotes(this.selection);
            if (!notes.length) {
              new Notice("No matching notes to export");
              return;
            }
            if (
              await this.plugin.exportXLSX(notes, {
                selection: this.selection,
                allProperties: this.allProperties,
              })
            )
              this.close();
          }),
      );
  }
  updateCount() {
    if (!this.countEl) return;
    const notes = this.plugin.exportSourceNotes(this.selection);
    const props = new Set();
    if (this.allProperties)
      for (const note of notes)
        for (const key of Object.keys(note.fm || {})) if (key !== "position") props.add(key);
    this.countEl.setText(
      notes.length +
        " notes selected" +
        (this.allProperties
          ? " · " + props.size + " frontmatter properties found"
          : " · standard Operations Manager columns"),
    );
  }
}
function omYamlValue(kind, v) {
  if (kind === "list") {
    const items = Array.isArray(v) ? v : v == null || v === "" ? [] : [v];
    return items.length ? items.map((x) => "\n  - " + JSON.stringify(x)).join("") : " []";
  }
  if (v === undefined || v === null || v === "") return "";
  if (kind === "checkbox") return " " + (v === true || v === "true");
  if (kind === "number" && Number.isFinite(Number(v))) return " " + Number(v);
  if (kind === "date") return iso(String(v)) ? " " + v : "";
  const text = String(v);
  // Plain words stay unquoted; anything YAML could misread (links, numbers, true/false…) is quoted.
  return /^[A-Za-z][\w .()-]*$/.test(text) && !/^(true|false|yes|no|on|off|null)$/i.test(text)
    ? " " + text.trim()
    : " " + JSON.stringify(text);
}
// Checkbox properties employee notes use for per-seat licenses and license-bundle types.
function omEmployeeLicenseFields(plugin) {
  const fields = new Map();
  for (const r of plugin.records().filter((x) => x.kind === "renewal")) {
    const types = omNextivaTypes(r.fm);
    if (types.length) {
      for (const t of types) if (t.field) fields.set(t.field.toLowerCase(), t.field);
    } else if (omTracksSeats(r.fm)) {
      const field = plugin.licenseField(r);
      fields.set(field.toLowerCase(), field);
    }
  }
  return [...fields.values()].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}
// Type label for a record: its Type property, or the closest match for older tag-based notes.
function omTypeLabel(r) {
  const type = omNoteType(r.fm || {});
  if (type) return OM_TYPE_LABELS[type];
  return (
    { device: "Equipment", renewal: "License", invoice: "Bill", order: "Order", project: "Project" }[
      r.kind
    ] ?? r.kind
  );
}
// Best guess at a Type for a note that has none, from its tags, folder and properties.
function omSuggestType(file, fm, licenseFields) {
  const tags = tagSet(fm);
  const folders = file.path.split("/").slice(0, -1);
  if (tags.has("employee")) return { type: "employee", why: "tag employee" };
  if (folders.some((p) => /^employees?$/i.test(p)))
    return { type: "employee", why: "in an Employees folder" };
  if (tags.has("equip_ug") || tags.has("equip_upgrade")) return { type: "computer", why: "tag equip_ug" };
  const kind = classified(file, fm);
  if (kind === "invoice")
    return { type: "invoice", why: tags.has("invoice") ? "tag invoice" : "in an Invoices folder" };
  if (kind === "order") return { type: "order", why: "tag order" };
  if (kind === "device")
    return tags.has("printer")
      ? { type: "equipment", why: "tags device and printer" }
      : { type: "computer", why: "tag device: change it for peripherals and office equipment" };
  if (kind === "renewal") {
    if (omNextivaTypes(fm).length || Array.isArray(fm.License_Types))
      return { type: "license", why: "has License_Types (a license bundle)" };
    const raw = fm.Agreement_Types ?? fm.Agreement_Type ?? fm.License_Type ?? [];
    const kinds = (Array.isArray(raw) ? raw : String(raw).split(",")).map((v) =>
      String(v).trim().toLowerCase(),
    );
    if (omTracksSeats(fm)) return { type: "license", why: "counts seats" };
    if (tags.has("printer") || kinds.some((k) => /lease|maintenance|agreement/.test(k)))
      return { type: "contract", why: "lease or maintenance agreement" };
    if (inferredRenewal(fm) || fm.Renewal != null)
      return { type: "contract", why: "has a renewal date: change it to Service or License if needed" };
    if (omBillingCycle(fm)) return { type: "service", why: "billed on a schedule with no renewal date" };
    return { type: "license", why: "tag license: check it" };
  }
  const keys = new Set(Object.keys(fm).map((k) => k.toLowerCase()));
  if (licenseFields.some((f) => keys.has(f.toLowerCase())))
    return { type: "employee", why: "has license checkboxes" };
  return null;
}
class OMAddTypeModal extends Modal {
  constructor(app, plugin) {
    super(app);
    this.plugin = plugin;
    this.filter = "all";
    const licenseFields = omEmployeeLicenseFields(plugin);
    this.rows = [];
    for (const file of app.vault.getMarkdownFiles()) {
      if (
        file.path
          .split("/")
          .slice(0, -1)
          .some((p) => p.toLowerCase() === "templates")
      )
        continue;
      const fm = app.metadataCache.getFileCache(file)?.frontmatter;
      // Notes that already have a type property of any kind are left alone.
      if (!fm || Object.keys(fm).some((k) => k.toLowerCase() === "type")) continue;
      const guess = omSuggestType(file, fm, licenseFields);
      if (guess) this.rows.push({ file, type: guess.type, why: guess.why, include: true });
    }
    const order = Object.keys(OM_TYPE_LABELS);
    this.rows.sort(
      (a, b) =>
        order.indexOf(a.type) - order.indexOf(b.type) ||
        a.file.basename.localeCompare(b.file.basename, undefined, { numeric: true, sensitivity: "base" }),
    );
  }
  onOpen() {
    this.modalEl.addClass("ops-add-type-modal");
    this.render();
  }
  visible() {
    return this.rows.filter((r) => this.filter === "all" || r.type === this.filter);
  }
  render() {
    const c = this.contentEl;
    c.empty();
    c.createEl("h2", { text: "Add Type to existing notes" });
    if (!this.rows.length) {
      c.createEl("p", { text: "Every operations note already has a Type. Nothing to do." });
      return;
    }
    c.createEl("p", {
      cls: "setting-item-description",
      text: "Check each suggestion, change any that are wrong, and untick notes you want to skip. Only a Type property is added; nothing else in the note changes. Notes that already have a Type are not listed.",
    });
    new Setting(c)
      .setName("Show")
      .addDropdown((d) => {
        d.addOption("all", "All suggestions (" + this.rows.length + ")");
        for (const [type, label] of Object.entries(OM_TYPE_LABELS)) {
          const n = this.rows.filter((r) => r.type === type).length;
          if (n) d.addOption(type, label + " (" + n + ")");
        }
        d.setValue(this.filter).onChange((v) => {
          this.filter = v;
          this.render();
        });
      })
      .addButton((b) =>
        b.setButtonText("Tick shown").onClick(() => {
          for (const r of this.visible()) r.include = true;
          this.render();
        }),
      )
      .addButton((b) =>
        b.setButtonText("Untick shown").onClick(() => {
          for (const r of this.visible()) r.include = false;
          this.render();
        }),
      );
    const list = c.createDiv({ cls: "ops-add-type-list" });
    for (const row of this.visible()) {
      new Setting(list)
        .setName(row.file.basename)
        .setDesc(row.file.path + " · " + row.why)
        .addDropdown((d) => {
          for (const [type, label] of Object.entries(OM_TYPE_LABELS)) d.addOption(type, label);
          d.setValue(row.type).onChange((v) => (row.type = v));
        })
        .addToggle((t) =>
          t.setValue(row.include).onChange((v) => {
            row.include = v;
            this.updateCount();
          }),
        );
    }
    new Setting(c)
      .addButton((b) => b.setButtonText("Cancel").onClick(() => this.close()))
      .addButton((b) => {
        this.applyButton = b;
        b.setCta().onClick(() => this.apply());
      });
    this.updateCount();
  }
  updateCount() {
    const n = this.rows.filter((r) => r.include).length;
    this.applyButton?.setButtonText("Add Type to " + n + (n === 1 ? " note" : " notes"));
  }
  async apply() {
    const chosen = this.rows.filter((r) => r.include);
    if (!chosen.length) {
      new Notice("No notes ticked");
      return;
    }
    let done = 0;
    const failed = [];
    for (const row of chosen) {
      try {
        await this.app.fileManager.processFrontMatter(row.file, (fm) => {
          if (Object.keys(fm).some((k) => k.toLowerCase() === "type")) return;
          // Put Type first, like new notes, keeping every other property as it was.
          const rest = { ...fm };
          for (const k of Object.keys(fm)) delete fm[k];
          fm.Type = row.type;
          Object.assign(fm, rest);
        });
        done++;
      } catch (e) {
        failed.push(row.file.basename);
      }
    }
    new Notice(
      "Added Type to " +
        done +
        (done === 1 ? " note" : " notes") +
        (failed.length ? ". Could not update: " + failed.join(", ") : ""),
    );
    this.close();
    this.plugin.refresh();
  }
  onClose() {
    this.contentEl.empty();
  }
}
// Older property names the plugin still reads in place of a template's name.
const OM_PROP_ALIASES = {
  Status: ["status"],
  Model: ["Device_Model"],
  Serial: ["Serial_Number"],
  Employee: ["Assigned_To"],
  Ordered: ["Order_Date"],
  Set_Up: ["Set_up"],
};
// Properties earlier versions wrote that nothing reads any more (history and alert switches now
// live in the plugin settings; a note's Type replaces Agreement_Types).
const OM_LEGACY_PROPS = [
  "Billing_History",
  "Last_Billed",
  "Renewal_History",
  "Last_Renewed",
  "Renewal_Review_Required",
  "Order_Status",
  "Renewal_Alerts",
  "Billing_Alerts",
  "Seat_Alerts",
  "Agreement_Types",
  "Agreement_Type",
  "Renewal_Term_Months",
];
// The key the plugin actually reads for a property, in the order the readers try them
// (Device_Model and Serial_Number win over Model and Serial).
function omFindKey(fm, name) {
  const aliases = OM_PROP_ALIASES[name] || [];
  const order = name === "Model" || name === "Serial" ? [...aliases, name] : [name, ...aliases];
  const names = order.map((n) => n.toLowerCase());
  for (const n of names) {
    const key = Object.keys(fm).find((k) => k.toLowerCase() === n);
    if (key) return key;
  }
  return null;
}
function omFilled(v) {
  return v !== undefined && v !== null && !(Array.isArray(v) && !v.length) && String(v).trim() !== "";
}
// Renewal terms that can't be advanced automatically.
const OM_VARIABLE_TERM = /variable|manual|negotiat|tbd|unknown|custom/i;
// One-off billing has no fixed amount or schedule: each use gets its own bill.
const OM_NOT_FOR_ONE_OFF = ["Cost", "Next_Billing"];
function omTemplateFor(file, fm) {
  const type = omNoteType(fm);
  if (!type) return null;
  const tags = tagSet(fm);
  const id =
    type === "computer" && (tags.has("equip_ug") || tags.has("equip_upgrade"))
      ? "replacement"
      : type === "license" && omIsBundle({ file, fm })
        ? "bundle"
        : type;
  return OM_TEMPLATES.find((t) => t.id === id) ?? null;
}
// Everything worth knowing about one typed note: properties it is missing, values that are empty or
// in the wrong format, and leftover properties from older versions.
function omCheckNote(file, fm, licenseFields) {
  const t = omTemplateFor(file, fm);
  if (!t) return null;
  const oneOff = !!omBillingCycle(fm)?.oneOff;
  const missing = [],
    values = [],
    legacy = [];
  const get = (name) => {
    const k = omFindKey(fm, name);
    return k ? fm[k] : undefined;
  };
  for (const [name, kind, , fallback, optional] of t.props) {
    const key = omFindKey(fm, name);
    if (!key) {
      if (!optional && !(oneOff && OM_NOT_FOR_ONE_OFF.includes(name)))
        missing.push({ name, kind, value: kind === "checkbox" ? false : kind === "list" ? [] : null });
      continue;
    }
    const v = fm[key];
    if (!omFilled(v)) continue;
    if (kind === "date" && !iso(v)) values.push(key + ' "' + v + "\" isn't a YYYY-MM-DD date");
    if (name === "Renewal") {
      if (!/^\d+$/.test(String(v).trim()) && !omMonthToMonth(fm) && !OM_VARIABLE_TERM.test(String(v)))
        values.push('Renewal "' + v + '" should be a number of days, month-to-month, or variable');
    } else if (kind === "number" && !Number.isFinite(Number(String(v).replace(/[$,\s]/g, ""))))
      values.push(key + ' "' + v + "\" isn't a number");
    if (name === "Billing" && !omBillingCycle(fm))
      values.push('Billing "' + v + '" should be Monthly, Quarterly, Semiannually, Annually or One-off');
  }
  if (t.id === "employee" && !omTerminated(fm))
    for (const field of licenseFields)
      if (!Object.keys(fm).some((k) => k.toLowerCase() === field.toLowerCase()))
        missing.push({ name: field, kind: "checkbox", value: false });
  const needs = (name, why) => {
    if (!omFilled(get(name))) values.push(name + " needs a value" + (why ? ": " + why : ""));
  };
  const cycle = omBillingCycle(fm);
  if (["computer", "replacement", "peripheral", "equipment", "order"].includes(t.id)) needs("Status");
  if (t.id === "replacement") needs("Approved", 'put the "Phase …" here so it is grouped');
  if (t.id === "service") {
    if (!oneOff) needs("Cost");
    needs("Billing", "needed for charge alerts");
    if (cycle && !cycle.oneOff) needs("Next_Billing", "needed for charge alerts");
  }
  if (t.id === "contract" && !inferredRenewal(fm)) needs("Next_Renewal", "needed for renewal alerts");
  if (t.id === "license") needs("License_Seats", "needed to count seats");
  if ((t.id === "license" || t.id === "bundle") && cycle && !cycle.oneOff)
    needs("Next_Billing", "needed for charge alerts");
  if (t.id === "bundle" && !omNextivaTypes(fm).length)
    values.push('No license types yet: add them with "Manage license types"');
  if (t.id === "invoice") {
    needs("Date");
    needs("Due_Date");
    needs("Cost");
  }
  for (const name of OM_LEGACY_PROPS) {
    const key = Object.keys(fm).find((k) => k.toLowerCase() === name.toLowerCase());
    if (key) legacy.push(key);
  }
  if (Array.isArray(fm.Nextiva_Licenses) && !fm.Nextiva_Licenses.length) legacy.push("Nextiva_Licenses");
  for (const [name, aliases] of Object.entries(OM_PROP_ALIASES)) {
    const present = [name, ...aliases].filter((n) => Object.prototype.hasOwnProperty.call(fm, n));
    if (present.length > 1)
      values.push(
        "Has both " +
          present.join(" and ") +
          "; " +
          omFindKey(fm, name) +
          " is the one used. Delete the other.",
      );
  }
  return { file, fm, template: t, missing, values, legacy };
}
class OMCheckNotesModal extends Modal {
  constructor(app, plugin) {
    super(app);
    this.plugin = plugin;
    this.filter = "all";
    this.scan();
  }
  scan() {
    const licenseFields = omEmployeeLicenseFields(this.plugin);
    this.checked = 0;
    this.rows = [];
    for (const file of this.app.vault.getMarkdownFiles()) {
      if (
        file.path
          .split("/")
          .slice(0, -1)
          .some((p) => p.toLowerCase() === "templates")
      )
        continue;
      const fm = this.app.metadataCache.getFileCache(file)?.frontmatter;
      if (!fm || ["archive", "archived"].includes(status(fm)) || tagSet(fm).has("archive")) continue;
      const result = omCheckNote(file, fm, licenseFields);
      if (!result) continue;
      this.checked++;
      if (result.missing.length || result.values.length || result.legacy.length)
        this.rows.push({ ...result, addMissing: result.missing.length > 0, removeLegacy: false });
    }
    const order = OM_TEMPLATES.map((t) => t.id);
    this.rows.sort(
      (a, b) =>
        order.indexOf(a.template.id) - order.indexOf(b.template.id) ||
        a.file.basename.localeCompare(b.file.basename, undefined, { numeric: true, sensitivity: "base" }),
    );
  }
  onOpen() {
    this.modalEl.addClass("ops-add-type-modal");
    this.render();
  }
  visible() {
    return this.rows.filter((r) => this.filter === "all" || r.template.id === this.filter);
  }
  render() {
    const c = this.contentEl;
    c.empty();
    c.createEl("h2", { text: "Check notes" });
    const fine = this.checked - this.rows.length;
    c.createEl("p", {
      cls: "setting-item-description",
      text:
        this.checked +
        " notes with a Type checked · " +
        fine +
        " have everything their type needs · " +
        this.rows.length +
        " listed below. Notes without a Type aren't checked: use Add Type to existing notes first.",
    });
    if (!this.rows.length) return;
    c.createEl("p", {
      cls: "setting-item-description",
      text: "Adding missing properties adds them empty, for you to fill in. Values that are empty or in the wrong format have to be fixed in the note itself. Removing old properties moves any alert switch or charge and renewal history into the plugin settings first.",
    });
    new Setting(c)
      .setName("Show")
      .addDropdown((d) => {
        d.addOption("all", "All (" + this.rows.length + ")");
        for (const t of OM_TEMPLATES) {
          const n = this.rows.filter((r) => r.template.id === t.id).length;
          if (n) d.addOption(t.id, t.label + " (" + n + ")");
        }
        d.setValue(this.filter).onChange((v) => {
          this.filter = v;
          this.render();
        });
      })
      .addButton((b) =>
        b.setButtonText("Tick all old-property removals shown").onClick(() => {
          for (const r of this.visible()) if (r.legacy.length) r.removeLegacy = true;
          this.render();
        }),
      );
    const list = c.createDiv({ cls: "ops-add-type-list ops-check-list" });
    for (const row of this.visible()) {
      const card = list.createDiv({ cls: "ops-check-note" });
      const head = card.createDiv({ cls: "ops-item-head" });
      const title = head.createEl("button", { cls: "ops-item-title", text: row.file.basename });
      title.onclick = () => this.app.workspace.getLeaf(false).openFile(row.file);
      head.createSpan({ cls: "ops-type-chip", text: row.template.label });
      if (row.values.length) {
        const ul = card.createEl("ul", { cls: "ops-check-issues" });
        for (const v of row.values) ul.createEl("li", { text: v });
      }
      if (row.missing.length)
        new Setting(card)
          .setName("Add missing properties")
          .setDesc(row.missing.map((m) => m.name).join(", "))
          .addToggle((t) =>
            t.setValue(row.addMissing).onChange((v) => {
              row.addMissing = v;
              this.updateCount();
            }),
          );
      if (row.legacy.length)
        new Setting(card)
          .setName("Remove old properties")
          .setDesc(row.legacy.join(", "))
          .addToggle((t) =>
            t.setValue(row.removeLegacy).onChange((v) => {
              row.removeLegacy = v;
              this.updateCount();
            }),
          );
    }
    new Setting(c)
      .addButton((b) => b.setButtonText("Close").onClick(() => this.close()))
      .addButton((b) => {
        this.applyButton = b;
        b.setCta().onClick(() => this.apply());
      });
    this.updateCount();
  }
  updateCount() {
    const n = this.rows.filter(
      (r) => (r.addMissing && r.missing.length) || (r.removeLegacy && r.legacy.length),
    ).length;
    this.applyButton?.setButtonText("Update " + n + (n === 1 ? " note" : " notes"));
    this.applyButton?.setDisabled(n === 0);
  }
  async apply() {
    const plugin = this.plugin;
    const chosen = this.rows.filter(
      (r) => (r.addMissing && r.missing.length) || (r.removeLegacy && r.legacy.length),
    );
    let done = 0;
    const failed = [];
    for (const row of chosen) {
      try {
        const path = row.file.path;
        if (row.removeLegacy && row.legacy.length) {
          // Keep alert choices and history before the note properties go.
          const alerts = { ...(plugin.settings.licenseAlerts?.[path] || {}) };
          for (const [kind, field] of [
            ["renewal", "Renewal_Alerts"],
            ["billing", "Billing_Alerts"],
            ["seats", "Seat_Alerts"],
          ])
            if (typeof alerts[kind] !== "boolean" && row.legacy.includes(field))
              alerts[kind] = omAlertSwitch(row.fm, null, kind, field);
          if (Object.keys(alerts).length)
            plugin.settings.licenseAlerts = { ...(plugin.settings.licenseAlerts || {}), [path]: alerts };
          for (const [kind, field] of [
            ["billing", "Billing_History"],
            ["renewal", "Renewal_History"],
          ]) {
            const raw = row.fm[field];
            for (const d of (Array.isArray(raw) ? raw : raw ? [raw] : []).map(iso).filter(Boolean))
              await plugin.addHistory(row.file, kind, d);
          }
          await plugin.saveData(plugin.settings);
        }
        await this.app.fileManager.processFrontMatter(row.file, (fm) => {
          // Renewal (days) replaces Renewal_Term_Months: carry a month count over as days.
          const months = Number(fm.Renewal_Term_Months);
          if (row.removeLegacy && Number.isInteger(months) && months > 0 && !omFilled(fm.Renewal))
            fm.Renewal = Math.round((months * 365) / 12);
          if (row.addMissing)
            for (const m of row.missing)
              if (!Object.keys(fm).some((k) => k.toLowerCase() === m.name.toLowerCase()))
                fm[m.name] = m.value;
          if (row.removeLegacy) for (const key of row.legacy) delete fm[key];
        });
        done++;
      } catch (e) {
        failed.push(row.file.basename);
      }
    }
    new Notice(
      "Updated " +
        done +
        (done === 1 ? " note" : " notes") +
        (failed.length ? ". Could not update: " + failed.join(", ") : ""),
    );
    plugin.refresh();
    // Show what is left once Obsidian has re-read the notes.
    window.setTimeout(() => {
      this.scan();
      this.render();
    }, 500);
  }
  onClose() {
    this.contentEl.empty();
  }
}
class OMNewNoteModal extends Modal {
  constructor(app, plugin, templateId = "computer") {
    super(app);
    this.plugin = plugin;
    this.select(templateId);
    this.title = "";
  }
  select(id) {
    this.template = OM_TEMPLATES.find((t) => t.id === id) ?? OM_TEMPLATES[0];
    this.values = {};
    this.folder = this.plugin.settings.typeFolders?.[this.template.id] ?? this.plugin.settings.folder;
  }
  onOpen() {
    this.render();
  }
  render() {
    const c = this.contentEl;
    c.empty();
    c.createEl("h2", { text: "New operations note" });
    new Setting(c).setName("What are you tracking?").addDropdown((d) => {
      for (const t of OM_TEMPLATES) d.addOption(t.id, t.label);
      d.setValue(this.template.id).onChange((v) => {
        this.select(v);
        this.render();
      });
    });
    c.createEl("p", { cls: "setting-item-description", text: this.template.description });
    new Setting(c).setName("Note name").addText((t) => {
      t.setValue(this.title).onChange((v) => (this.title = v));
      window.setTimeout(() => t.inputEl.focus(), 0);
    });
    new Setting(c)
      .setName("Folder")
      .addText((t) => t.setValue(this.folder).onChange((v) => (this.folder = v.trim())));
    for (const [name, kind, help, , optional] of this.template.props) {
      if (optional || kind === "list") continue;
      const line = new Setting(c).setName(name).setDesc(help);
      if (kind === "checkbox")
        line.addToggle((t) => t.setValue(false).onChange((v) => (this.values[name] = v)));
      else
        line.addText((t) => {
          if (kind === "date") t.inputEl.type = "date";
          if (kind === "number") t.inputEl.type = "number";
          t.onChange((v) => (this.values[name] = v.trim()));
        });
    }
    c.createEl("p", {
      cls: "setting-item-description",
      text: "Leave anything you don't know yet blank: the property is still added so you can fill it in later.",
    });
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Create note")
        .setCta()
        .onClick(() => this.create()),
    );
  }
  async create() {
    const title = this.title.trim();
    if (!title) {
      new Notice("Enter a note name");
      return;
    }
    const t = this.template;
    for (const [name, kind] of t.props)
      if (kind === "date" && this.values[name] && !iso(this.values[name])) {
        new Notice(name + " must be a YYYY-MM-DD date");
        return;
      }
    const folder = normalizePath(this.folder || "/");
    const path = normalizePath(
      (folder === "/" ? "" : folder + "/") + title.replace(/[\\/:*?"<>|#^[\]]/g, "-") + ".md",
    );
    if (this.app.vault.getAbstractFileByPath(path)) {
      new Notice("A note with that name already exists");
      return;
    }
    const lines = ["Type: " + t.type];
    if (t.tags) lines.push("tags:" + omYamlValue("list", t.tags));
    for (const [name, kind, , fallback, optional] of t.props) {
      if (optional) continue;
      const billing = this.values.Billing ?? t.props.find((p) => p[0] === "Billing")?.[3];
      if (OM_NOT_FOR_ONE_OFF.includes(name) && omBillingCycle({ Billing: billing })?.oneOff) continue;
      let value = this.values[name] ?? fallback;
      if (name === "License_Field" && !value) value = title;
      lines.push(name + ":" + omYamlValue(kind, value));
    }
    if (t.id === "employee")
      for (const field of omEmployeeLicenseFields(this.plugin)) lines.push(field + ": false");
    try {
      if (folder !== "/") {
        let current = "";
        for (const part of folder.split("/")) {
          current = current ? current + "/" + part : part;
          if (!this.app.vault.getAbstractFileByPath(current)) await this.app.vault.createFolder(current);
        }
      }
      const file = await this.app.vault.create(path, "---\n" + lines.join("\n") + "\n---\n");
      if ((this.plugin.settings.typeFolders?.[t.id] ?? this.plugin.settings.folder) !== this.folder) {
        this.plugin.settings.typeFolders = {
          ...(this.plugin.settings.typeFolders || {}),
          [t.id]: this.folder,
        };
        await this.plugin.saveData(this.plugin.settings);
      }
      this.close();
      await this.app.workspace.getLeaf(false).openFile(file);
      this.plugin.refresh();
    } catch (e) {
      new Notice("Could not create note: " + e.message);
    }
  }
  onClose() {
    this.contentEl.empty();
  }
}
class OMBillingModal extends Modal {
  constructor(app, plugin, record) {
    super(app);
    this.plugin = plugin;
    this.record = record;
    this.cycle = String(record.fm.Billing ?? "");
    this.date = iso(record.fm.Next_Billing) ?? "";
  }
  onOpen() {
    const c = this.contentEl;
    c.createEl("h2", { text: "Billing schedule: " + this.record.file.basename });
    c.createEl("p", {
      text: "Billing is separate from contract renewal. Cost is the total amount charged each billing cycle, regardless of License_Seats.",
    });
    new Setting(c).setName("Billing frequency").addDropdown((d) => {
      for (const [value, label] of [
        ["", "Not set"],
        ["Monthly", "Monthly"],
        ["Quarterly", "Quarterly"],
        ["Semiannually", "Every 6 months"],
        ["Annually", "Annually"],
        ["One-off", "One-off (only charged when used)"],
      ])
        d.addOption(value, label);
      if (this.cycle && !["Monthly", "Quarterly", "Semiannually", "Annually", "One-off"].includes(this.cycle))
        d.addOption(this.cycle, this.cycle);
      d.setValue(this.cycle).onChange((v) => (this.cycle = v));
    });
    new Setting(c)
      .setName("Next billing date")
      .setDesc("YYYY-MM-DD")
      .addText((t) =>
        t
          .setPlaceholder("2026-11-01")
          .setValue(this.date)
          .onChange((v) => (this.date = v.trim())),
      );
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Save billing")
        .setCta()
        .onClick(async () => {
          if (this.cycle && !omBillingCycle({ Billing: this.cycle })) {
            new Notice("Choose a supported billing frequency");
            return;
          }
          if (this.date && !iso(this.date)) {
            new Notice("Enter a valid YYYY-MM-DD date");
            return;
          }
          try {
            await this.app.fileManager.processFrontMatter(this.record.file, (fm) => {
              if (this.cycle) fm.Billing = this.cycle;
              else delete fm.Billing;
              if (this.date) fm.Next_Billing = this.date;
              else delete fm.Next_Billing;
            });
            this.record.fm.Billing = this.cycle;
            this.record.fm.Next_Billing = this.date;
            this.close();
            this.plugin.refresh();
            if (this.plugin.licenseSettingsTab) this.plugin.licenseSettingsTab.display();
            new Notice("Billing schedule saved");
          } catch (e) {
            new Notice("Unable to save billing: " + e.message);
          }
        }),
    );
  }
}
class OMRecordChargeModal extends Modal {
  constructor(app, plugin, record) {
    super(app);
    this.plugin = plugin;
    this.record = record;
  }
  onOpen() {
    const current = omBillingInfo(this.record.fm);
    const c = this.contentEl;
    c.createEl("h2", { text: "Record billing charge" });
    if (!current) {
      c.createEl("p", { text: "Set Billing and Next_Billing in plugin settings first." });
      return;
    }
    const next = addMonths(current.date, current.months);
    c.createEl("p", { text: "Current charge date: " + current.date });
    c.createEl("p", { text: "Next charge date: " + next });
    if (current.amount !== null) c.createEl("p", { text: "Estimated charge: " + omUSD(current.amount) });
    c.createEl("p", {
      text: "Confirm only after this scheduled charge has occurred. The contract renewal date will not change.",
    });
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Record charge")
        .setCta()
        .onClick(async () => {
          try {
            await this.app.fileManager.processFrontMatter(this.record.file, (fm) => {
              if (iso(fm.Next_Billing) !== current.date)
                throw Error("Billing date changed; reopen this dialog");
              fm.Next_Billing = next;
            });
            await this.plugin.addHistory(this.record.file, "billing", current.date);
            this.record.fm.Next_Billing = next;
            this.close();
            this.plugin.refresh();
            if (this.plugin.licenseSettingsTab) this.plugin.licenseSettingsTab.display();
            new Notice("Charge recorded; next billing date updated");
          } catch (e) {
            new Notice("Unable to record charge: " + e.message);
          }
        }),
    );
  }
}
class RenewModal extends Modal {
  constructor(app, plugin, file, fm) {
    super(app);
    this.plugin = plugin;
    this.file = file;
    this.fm = fm;
  }
  onOpen() {
    const c = this.contentEl;
    c.createEl("h2", { text: "Complete renewal: " + this.file.basename });
    const current = inferredRenewal(this.fm);
    if (!current) {
      c.createEl("p", {
        text: "Manual review required: enter a confirmed Next_Renewal date, or a valid Ordered date and Renewal term in days.",
      });
      return;
    }
    const days = Number(this.fm.Renewal);
    const variable = OM_VARIABLE_TERM.test(
      String(this.fm.Renewal ?? "") +
        " " +
        String(this.fm.Renewal_Term ?? "") +
        " " +
        String(this.fm.Renewal_Type ?? ""),
    );
    const fixedDays = Number.isInteger(days) && days > 0 && days <= 36500;
    if (variable || !fixedDays) {
      c.createEl("p", {
        text: "Manual review required: Renewal is variable, missing, or not a positive number of days. No automatic date advancement will occur.",
      });
      return;
    }
    const old = current.date;
    const d = new Date(old + "T00:00:00Z");
    d.setUTCDate(d.getUTCDate() + days);
    const next = dateISO(d);
    c.createEl("p", { text: "Current renewal: " + old + (current.estimated ? " (estimated)" : "") });
    c.createEl("p", { text: "Next renewal after completion: " + next });
    c.createEl("p", {
      text:
        "This sets Next_Renewal in the note and keeps the completed date in the plugin's renewal history." +
        (this.plugin.settings.renewalFollowUpTask
          ? " A follow-up task is added to the end of the note."
          : ""),
    });
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Confirm completion")
        .setCta()
        .onClick(async () => {
          try {
            await this.app.fileManager.processFrontMatter(this.file, (fm) => {
              const check = inferredRenewal(fm);
              if (!check || check.date !== old) throw Error("Renewal date changed; reopen this dialog");
              fm.Next_Renewal = next;
            });
            await this.plugin.addHistory(this.file, "renewal", old);
            if (this.plugin.settings.renewalFollowUpTask) {
              const note = await this.app.vault.read(this.file);
              const task = "- [ ] Review renewal for " + this.file.basename + " 📅 " + next;
              if (!note.includes(task))
                await this.app.vault.append(
                  this.file,
                  (note.endsWith("\n") ? "" : "\n") + "\n## Renewal follow-up\n" + task + "\n",
                );
            }
            new Notice("Renewal completed; next renewal date saved");
            this.close();
            this.plugin.refresh();
          } catch (e) {
            new Notice("Renewal needs review: " + e.message);
          }
        }),
    );
  }
}
class PaidModal extends Modal {
  constructor(app, plugin, file) {
    super(app);
    this.plugin = plugin;
    this.file = file;
    this.date = dateISO(new Date());
  }
  onOpen() {
    const c = this.contentEl;
    c.createEl("h2", { text: "Mark invoice paid" });
    new Setting(c)
      .setName("Payment date")
      .addText((t) => t.setValue(this.date).onChange((v) => (this.date = v)));
    new Setting(c).addButton((b) =>
      b
        .setButtonText("Confirm payment")
        .setCta()
        .onClick(async () => {
          if (!iso(this.date)) {
            new Notice("Use YYYY-MM-DD");
            return;
          }
          try {
            await this.app.fileManager.processFrontMatter(this.file, (fm) => {
              fm.Status = "complete";
              fm.Paid_Date = this.date;
            });
            new Notice("Invoice marked paid");
            this.close();
            this.plugin.refresh();
          } catch (e) {
            new Notice(e.message);
          }
        }),
    );
  }
}
// Card title, with a small label for note types that share a tab with other types.
function omCardTitle(info, r, app) {
  const row = info.createDiv({ cls: "ops-item-head" });
  const title = row.createEl("button", { cls: "ops-item-title", text: r.file.basename });
  title.onclick = () => app.workspace.getLeaf(false).openFile(r.file);
  const type = omNoteType(r.fm);
  if (["peripheral", "equipment", "service", "contract", "license"].includes(type))
    row.createSpan({ cls: "ops-type-chip", text: OM_TYPE_LABELS[type] });
  return title;
}
function omIsDeviceNote(fm) {
  return (
    ["computer", "peripheral", "equipment"].includes(omNoteType(fm)) ||
    tagSet(fm).has("device") ||
    tagSet(fm).has("equip_ug") ||
    tagSet(fm).has("equip_upgrade")
  );
}
function omCardLinkedText(parent, value, app, sourcePath) {
  const text = Array.isArray(value) ? value.map(String).join(", ") : String(value ?? "");
  const parts = text.split(/(\[\[[^\]]+\]\])/g);
  for (const part of parts) {
    const match = part.match(/^\[\[([^\]]+)\]\]$/);
    if (!match) {
      parent.createSpan({ text: part });
      continue;
    }
    const [target, alias] = match[1].split("|");
    const link = parent.createEl("a", { cls: "internal-link", text: alias || target });
    link.href = "#";
    link.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      app.workspace.openLinkText(target, sourcePath, false);
    };
  }
}
// "Related:" line with the note's Related links, clickable.
function omRenderRelated(info, fm, app, sourcePath) {
  const related = fm.Related ?? fm.related;
  if (!omFilled(related)) return;
  const line = info.createDiv({ cls: "ops-equipment-note" });
  line.createSpan({ text: "Related: " });
  omCardLinkedText(line, related, app, sourcePath);
}
function omDashboardDeviceDetails(info, record, app, followUp) {
  const employee = record.fm.Employee ?? record.fm.Assigned_To;
  const hasEmployee = employee !== undefined && employee !== null && String(employee).trim() !== "";
  if (hasEmployee || followUp) {
    const line = info.createEl("small");
    if (hasEmployee) omCardLinkedText(line, employee, app, record.file.path);
    if (followUp) line.createSpan({ text: (hasEmployee ? " · " : "") + "Follow up" });
  }
  const notes = record.fm.Notes;
  if (notes !== undefined && notes !== null && String(notes).trim()) {
    const noteLine = info.createDiv({ cls: "ops-equipment-note" });
    noteLine.createSpan({ text: "Notes: " });
    omCardLinkedText(noteLine, notes, app, record.file.path);
  }
  omRenderRelated(info, record.fm, app, record.file.path);
}
const OM_CARD_ICONS = {
  renewal: "clock",
  invoice: "file-text",
  device: "laptop",
  order: "package",
  followup: "flag",
};
const OM_LEGACY_ICONS = {
  "◷": "clock",
  "▤": "file-text",
  "▭": "laptop",
  "▣": "package",
  "◉": "circle-dot",
  "◆": "diamond",
  "★": "star",
  "⚑": "flag",
  "✓": "check",
  "!": "alert-circle",
  "•": "dot",
  "—": "none",
};
class OMIconPicker extends require("obsidian").SuggestModal {
  constructor(app, plugin, kind, label, onPick) {
    super(app);
    this.plugin = plugin;
    this.kind = kind;
    this.label = label;
    this.onPick = onPick;
    this.setPlaceholder("Search Lucide icons by name...");
    this.setInstructions([
      { command: "↑↓", purpose: "Navigate" },
      { command: "↵", purpose: "Select" },
    ]);
    this.icons = ["none", ...require("obsidian").getIconIds().sort()];
  }
  getSuggestions(query) {
    const q = query.trim().toLowerCase().replace(/\s+/g, "-");
    return this.icons.filter((id) => id.includes(q)).slice(0, 100);
  }
  renderSuggestion(id, el) {
    el.addClass("ops-icon-suggestion");
    const preview = el.createSpan({ cls: "ops-icon-preview" });
    if (id !== "none") require("obsidian").setIcon(preview, id);
    el.createSpan({ text: id === "none" ? "No icon" : id.replace(/-/g, " ") });
  }
  onChooseSuggestion(id) {
    this.onPick(id);
  }
}
class OMEmbeddedDashboard extends require("obsidian").MarkdownRenderChild {
  constructor(el, plugin) {
    super(el);
    this.contentEl = el;
    this.plugin = plugin;
    this.app = plugin.app;
    this.page = "dashboard";
  }
  async onload() {
    this.plugin.embeddedDashboards.add(this);
    await this.render();
  }
  onunload() {
    this.plugin.embeddedDashboards.delete(this);
  }
  async render() {
    return OperationsView.prototype.render.call(this);
  }
  renderInvoices(root, records) {
    return OperationsView.prototype.renderInvoices.call(this, root, records);
  }
  renderOrders(root, records) {
    return OperationsView.prototype.renderOrders.call(this, root, records);
  }
  renderLicenseOverviewInline(info, record) {
    return OperationsView.prototype.renderLicenseOverviewInline.call(this, info, record);
  }
}
module.exports = class OperationsManager extends Plugin {
  async onload() {
    this.embeddedDashboards = new Set();
    this.registerMarkdownCodeBlockProcessor("operations-manager", (_source, el, ctx) => {
      ctx.addChild(new OMEmbeddedDashboard(el, this));
    });
    this.settings = Object.assign(
      {
        folder: "Operations Records",
        renewalWarningDays: 90,
        cardIcons: {},
        collapsedSections: {},
        licenseAlerts: {},
        licenseHistory: {},
        typeFolders: {},
        renewalFollowUpTask: false,
      },
      await this.loadData(),
    );
    this.settings.renewalWarningDays = 90;
    OM_SETTINGS = this.settings;
    this.registerView(VIEW, (leaf) => new OperationsView(leaf, this));
    this.addRibbonIcon("layout-dashboard", "Open Operations Manager", () => this.activate());
    this.addCommand({
      id: "open-dashboard",
      name: "Open Operations Manager",
      callback: () => this.activate(),
    });
    this.addCommand({
      id: "check-notes",
      name: "Check notes for missing properties",
      callback: () => new OMCheckNotesModal(this.app, this).open(),
    });
    this.addCommand({
      id: "add-type",
      name: "Add Type to existing notes",
      callback: () => new OMAddTypeModal(this.app, this).open(),
    });
    this.addCommand({
      id: "new-note",
      name: "Create a new operations note",
      callback: () => new OMNewNoteModal(this.app, this).open(),
    });
    this.addSettingTab(
      new (class extends require("obsidian").PluginSettingTab {
        display() {
          this.plugin.licenseSettingsTab = this;
          this.containerEl.empty();
          this.containerEl.createEl("h2", { text: "Operations Manager" });
          this.containerEl.createEl("h3", { text: "General" });
          new Setting(this.containerEl)
            .setName("Folder for new records")
            .setDesc(
              "Where new notes are created, unless you pick a different folder for a type when creating one.",
            )
            .addText((t) =>
              t.setValue(this.plugin.settings.folder).onChange(async (v) => {
                this.plugin.settings.folder = v;
                await this.plugin.saveData(this.plugin.settings);
              }),
            );
          new Setting(this.containerEl)
            .setName("Add a follow-up task when renewing")
            .setDesc(
              'When on, "Renew" also adds a "Renewal follow-up" task to the end of the note. Off by default.',
            )
            .addToggle((t) =>
              t.setValue(!!this.plugin.settings.renewalFollowUpTask).onChange(async (v) => {
                this.plugin.settings.renewalFollowUpTask = v;
                await this.plugin.saveData(this.plugin.settings);
              }),
            );
          this.plugin.renderTypeGuide(this.containerEl);
          const iconSection = this.containerEl.createEl("details", { cls: "ops-settings-icon-section" });
          iconSection.createEl("summary", { text: "Customize card icons" });
          iconSection.createEl("p", {
            cls: "setting-item-description",
            text: "Optional. Choose a Lucide icon for each Dashboard card type.",
          });
          for (const [kind, label] of [
            ["renewal", "Renewals & licenses"],
            ["invoice", "Invoices"],
            ["device", "Equipment"],
            ["order", "Orders"],
            ["followup", "Follow up"],
          ]) {
            const setting = new Setting(iconSection).setName(label);
            const selected = this.plugin.cardIcon(kind);
            const preview = setting.controlEl.createSpan({ cls: "ops-icon-preview" });
            if (selected !== "none") require("obsidian").setIcon(preview, selected);
            setting.addButton((b) =>
              b.setButtonText(selected === "none" ? "No icon" : selected).onClick(() =>
                new OMIconPicker(this.app, this.plugin, kind, label, async (icon) => {
                  this.plugin.settings.cardIcons = {
                    ...(this.plugin.settings.cardIcons || {}),
                    [kind]: icon,
                  };
                  await this.plugin.saveData(this.plugin.settings);
                  this.plugin.refresh();
                  this.display();
                }).open(),
              ),
            );
          }
          this.plugin.renderLicenseInventory(
            this.containerEl,
            this.plugin
              .records()
              .filter((r) => r.kind === "renewal")
              .sort((a, b) =>
                a.file.basename.localeCompare(b.file.basename, undefined, {
                  numeric: true,
                  sensitivity: "base",
                }),
              ),
          );
        }
      })(this.app, this),
    );
    // Renewal dates estimated from Ordered + Renewal are worked out when displayed; they are never
    // written back to the note.
    const refreshSoon = require("obsidian").debounce(() => this.refresh(), 300, true);
    this.registerEvent(this.app.metadataCache.on("changed", () => refreshSoon()));
    this.registerEvent(
      this.app.vault.on("rename", (file, oldPath) => this.renameSettingsKeys(oldPath, file.path)),
    );
  }
  // Per-note settings (alert switches and rules, history, open sections) are keyed by file path,
  // so follow the note when it is renamed or moved.
  async renameSettingsKeys(oldPath, newPath) {
    let changed = false;
    for (const name of ["licenseAlertRules", "licenseAlerts", "licenseHistory"]) {
      const map = this.settings[name];
      if (map && Object.prototype.hasOwnProperty.call(map, oldPath)) {
        map[newPath] = map[oldPath];
        delete map[oldPath];
        changed = true;
      }
    }
    const sections = this.settings.collapsedSections || {};
    for (const key of Object.keys(sections))
      if (key.endsWith(":" + oldPath)) {
        sections[key.slice(0, -oldPath.length) + newPath] = sections[key];
        delete sections[key];
        changed = true;
      }
    if (changed) await this.saveData(this.settings);
  }
  cardIcon(kind) {
    const value = this.settings.cardIcons?.[kind] ?? OM_CARD_ICONS[kind] ?? "circle";
    return OM_LEGACY_ICONS[value] ?? value;
  }
  renderCardIcon(parent, kind) {
    const name = this.cardIcon(kind);
    const holder = parent.createSpan({ cls: "ops-item-icon" });
    if (name !== "none") require("obsidian").setIcon(holder, name);
    return holder;
  }
  records() {
    const result = [];
    for (const file of this.app.vault.getMarkdownFiles()) {
      if (
        file.path
          .split("/")
          .slice(0, -1)
          .some((part) => part.toLowerCase() === "templates")
      )
        continue;
      const fm = this.app.metadataCache.getFileCache(file)?.frontmatter;
      if (!fm) continue;
      if (omTerminated(fm)) continue;
      if (tagSet(fm, file, this.app).has("call_notes") || tagSet(fm, file, this.app).has("meeting_notes"))
        continue;
      const tags = tagSet(fm, file, this.app);
      const kind = tags.has("equip_ug") ? "order" : classified(file, fm);
      if (!kind) continue;
      const st = status(fm);
      const done = [
        "paid",
        "complete",
        "completed",
        "deployed",
        "closed",
        "retired",
        "unused",
        "in stock",
        "spare",
        "available",
        "cancelled",
        "canceled",
      ].includes(st);
      if (["archive", "archived"].includes(st) || tags.has("archive")) continue;
      let reason = "",
        attention = false;
      const monthly = kind === "renewal" && omMonthToMonth(fm, file);
      const renewal = monthly ? null : inferredRenewal(fm);
      if (kind === "renewal") {
        if (monthly) {
          reason = "Month-to-month · No expiration";
        } else if (renewal) {
          const days = daysUntil(renewal.date);
          const qualifier = renewal.estimated ? "Estimated renewal" : "Renewal";
          reason = qualifier + " " + renewal.date;
          if (omRenewalAlertsEnabled(fm, file) && omAlertColor(this, file, "renewal", days)) {
            attention = true;
            reason += days < 0 ? " (past due by " + -days + " days)" : " (in " + days + " days)";
          }
        } else if (!omNeedsRenewalDate(fm)) {
          reason = "";
        } else {
          reason = "Renewal date needs review: enter Next_Renewal or Ordered and Renewal (days)";
          attention = omRenewalAlertsEnabled(fm, file);
        }
        const billing = omBillingAlertsEnabled(fm, file) ? omBillingInfo(fm) : null;
        if (billing && omAlertColor(this, file, "billing", billing.days)) {
          attention = true;
          if (
            !renewal ||
            !omRenewalAlertsEnabled(fm, file) ||
            !omAlertColor(this, file, "renewal", daysUntil(renewal.date))
          )
            reason =
              "Upcoming charge " +
              billing.date +
              (billing.days < 0
                ? " (past due by " + -billing.days + " days)"
                : " (in " + billing.days + " days)");
        }
      }
      if (kind === "invoice") {
        const paid =
          ["paid", "complete", "completed"].includes(st) ||
          fm.Paid === true ||
          !!iso(fm.Payment_Date ?? fm.Paid_Date);
        const due = iso(fm.Due_Date);
        const invoiceDate = iso(fm.Date);
        reason = paid
          ? "Paid"
          : (invoiceDate ? "Invoice date " + invoiceDate + " · " : "") +
            (due ? "Unpaid · Due " + due : "Unpaid · Due date not recorded");
        attention = !paid;
      }
      if (kind === "project") {
        reason = st || "Planned";
        attention = false;
      }
      if (kind === "order") {
        reason = st || "Status not set";
        attention = !done;
      }
      if (kind === "device") {
        reason = st || "Status not set";
        attention = !done && ["in progress", "pending", "ordered", "to do", "awaiting setup"].includes(st);
      }
      if (
        (kind === "order" || kind === "device") &&
        (tags.has("equip_ug") || tags.has("equip_upgrade") || tags.has("equip_ug"))
      ) {
        const ordered = iso(fm.Ordered ?? fm.Order_Date);
        const setup = iso(fm.Set_Up ?? fm.Set_up);
        attention = tags.has("equip_ug") ? st !== "complete" : !ordered || !setup;
        reason =
          st === "complete"
            ? "Complete"
            : !ordered
              ? "Not ordered — Ordered date missing"
              : !setup
                ? "Ordered " + ordered + " — Set_Up date missing"
                : "Follow-up needed · Status: " + (st || "not set");
      }
      result.push({ file, fm, kind, attention, reason, done, equipUp: tags.has("equip_ug") });
    }
    return result.sort(
      (a, b) => Number(b.attention) - Number(a.attention) || a.file.basename.localeCompare(b.file.basename),
    );
  }
  renderTypeGuide(root) {
    const section = root.createDiv({ cls: "ops-type-guide" });
    section.createEl("h3", { text: "Note types & properties" });
    section.createEl("p", {
      cls: "setting-item-description",
      text: "Give each note a Type property so the plugin knows what it is, then fill in the properties listed for that type. A lowercase type property works too, and notes that use tags (device, equip_ug, license, invoice, order) instead keep working.",
    });
    const addType = section.createEl("button", { text: "Add Type to existing notes…" });
    addType.onclick = () => new OMAddTypeModal(this.app, this).open();
    const check = section.createEl("button", { text: "Check notes for missing properties…" });
    check.onclick = () => new OMCheckNotesModal(this.app, this).open();
    for (const t of OM_TEMPLATES) {
      const details = section.createEl("details", { cls: "ops-settings-license" });
      details.createEl("summary", { cls: "ops-settings-license-summary", text: t.label });
      const body = details.createDiv({ cls: "ops-settings-license-body" });
      body.createEl("p", { cls: "setting-item-description", text: t.description });
      body.createEl("p", {
        cls: "ops-settings-facts",
        text: "Type: " + t.type + (t.tags ? " · tags: " + t.tags.join(", ") : ""),
      });
      const table = body.createEl("table", { cls: "ops-type-guide-table" });
      const head = table.createEl("tr");
      for (const label of ["Property", "Kind", "What to enter"]) head.createEl("th", { text: label });
      for (const [name, kind, help, , optional] of t.props) {
        const tr = table.createEl("tr");
        tr.createEl("td", { text: name });
        tr.createEl("td", { text: kind });
        tr.createEl("td", { text: (optional ? "Optional. " : "") + help });
      }
      if (t.id === "employee") {
        const fields = omEmployeeLicenseFields(this);
        const tr = table.createEl("tr");
        tr.createEl("td", { text: fields.length ? fields.join(", ") : "One per license" });
        tr.createEl("td", { text: "checkbox" });
        tr.createEl("td", { text: "Ticked when this person has that license seat." });
      }
      const create = body.createEl("button", { text: "Create a note of this type" });
      create.onclick = () => new OMNewNoteModal(this.app, this, t.id).open();
    }
  }
  async addHistory(file, kind, date) {
    const path = file.path;
    const entry = { ...(this.settings.licenseHistory?.[path] || {}) };
    const list = Array.isArray(entry[kind]) ? [...entry[kind]] : [];
    if (!list.includes(date)) list.push(date);
    entry[kind] = list.sort();
    this.settings.licenseHistory = { ...(this.settings.licenseHistory || {}), [path]: entry };
    await this.saveData(this.settings);
  }
  // Most recent recorded date, from the plugin history or from history older versions wrote to the note.
  lastHistory(license, kind) {
    const field = kind === "billing" ? "Billing_History" : "Renewal_History";
    const legacy = license.fm[field];
    const dates = [
      ...(this.settings.licenseHistory?.[license.file.path]?.[kind] || []),
      ...(Array.isArray(legacy) ? legacy : legacy ? [legacy] : []),
    ]
      .map(iso)
      .filter(Boolean)
      .sort();
    return dates[dates.length - 1] ?? null;
  }
  licenseField(license) {
    return String(license.fm.License_Field ?? license.fm.License_Name ?? license.file.basename).trim();
  }
  employeeNotes() {
    return this.app.vault
      .getMarkdownFiles()
      .filter(
        (file) =>
          !file.path
            .split("/")
            .slice(0, -1)
            .some((p) => p.toLowerCase() === "templates"),
      )
      .map((file) => ({ file, fm: this.app.metadataCache.getFileCache(file)?.frontmatter }))
      .filter(
        (e) =>
          e.fm &&
          !omTerminated(e.fm) &&
          !tagSet(e.fm).has("call_notes") &&
          !tagSet(e.fm).has("meeting_notes") &&
          !tagSet(e.fm).has("archive"),
      );
  }
  licenseInventory(license) {
    const field = this.licenseField(license);
    const employees = this.employeeNotes().filter(
      (e) => Object.keys(e.fm).some((k) => k.toLowerCase() === field.toLowerCase()) || omIsEmployee(e),
    );
    employees.sort((a, b) =>
      a.file.basename.localeCompare(b.file.basename, undefined, { sensitivity: "base", numeric: true }),
    );
    const assigned = employees.filter((e) =>
      Object.entries(e.fm).some(([k, v]) => k.toLowerCase() === field.toLowerCase() && omLicenseAssigned(v)),
    ).length;
    const raw = license.fm.License_Seats;
    const total =
      raw !== undefined && raw !== null && raw !== "" && Number.isInteger(Number(raw)) && Number(raw) >= 0
        ? Number(raw)
        : null;
    return { field, employees, assigned, total, available: total === null ? null : total - assigned };
  }
  renderLicenseInventory(root, records) {
    const section = root.createDiv({ cls: "ops-settings-licenses" });
    section.createEl("h3", { text: "Services, contracts & licenses" });
    section.createEl("p", {
      cls: "setting-item-description",
      text: "Select one to manage its alerts, billing, renewal date, seats and assignments. Your notes remain the source of truth.",
    });
    const toolbar = section.createDiv({ cls: "ops-settings-toolbar" });
    const search = toolbar.createEl("input", {
      type: "search",
      placeholder: "Find a service, contract or license…",
      cls: "ops-settings-search",
    });
    search.setAttribute("aria-label", "Search licenses");
    const count = toolbar.createSpan({ cls: "ops-settings-count" });
    const list = section.createDiv({ cls: "ops-settings-list" });
    const show = () => {
      list.empty();
      const filtered = records.filter((r) =>
        r.file.basename.toLowerCase().includes(search.value.trim().toLowerCase()),
      );
      count.setText(filtered.length + " of " + records.length);
      if (!filtered.length) {
        list.createEl("p", { cls: "setting-item-description", text: "No matching licenses." });
        return;
      }
      for (const license of filtered) {
        const inv = this.licenseInventory(license),
          multi = omNextivaTypes(license.fm),
          hasTypes = multi.length > 0,
          seatBased = omTracksSeats(license.fm);
        const summary = hasTypes ? omNextivaSummary(this, license) : null;
        const bought = hasTypes ? summary.rows.reduce((n, x) => n + x.purchased, 0) : inv.total;
        const used = hasTypes ? summary.rows.reduce((n, x) => n + x.assigned, 0) : inv.assigned;
        const available = bought === null ? null : bought - used;
        const charge = omRenewalEstimate(license.fm)?.amount ?? null;
        const frequency = omBillingCycle(license.fm)?.label?.toLowerCase() ?? "cycle";
        const key = "settings:license:" + license.file.path;
        const details = list.createEl("details", { cls: "ops-settings-license" });
        details.open = this.settings.collapsedSections?.[key] === true;
        details.addEventListener("toggle", () => {
          if (this.settings.collapsedSections?.[key] === details.open) return;
          this.settings.collapsedSections = {
            ...(this.settings.collapsedSections || {}),
            [key]: details.open,
          };
          this.saveData(this.settings);
        });
        const head = details.createEl("summary", { cls: "ops-settings-license-summary" });
        const heading = head.createSpan({ cls: "ops-settings-license-name", text: license.file.basename });
        const chips = head.createSpan({ cls: "ops-settings-license-meta" });
        if (seatBased && bought !== null)
          chips.createSpan({ text: available + " available / " + bought + " purchased" });
        if (charge !== null) chips.createSpan({ text: omUSD(charge) + " / " + frequency });
        const body = details.createDiv({ cls: "ops-settings-license-body" });
        const actions = body.createDiv({ cls: "ops-settings-actions" });
        const open = actions.createEl("button", { text: "Open source note" });
        open.onclick = () => this.app.workspace.getLeaf(false).openFile(license.file);
        const alerts = body.createDiv({ cls: "ops-settings-group" });
        alerts.createEl("h4", { text: "Dashboard alerts" });
        alerts.createEl("p", {
          cls: "setting-item-description",
          text: "Choose what deserves attention. Renewal colors keep the existing 90-day warning system.",
        });
        for (const [kind, field, label, description, enabled] of [
          [
            "renewal",
            "Renewal_Alerts",
            "Contract renewals",
            "Notify when this contract approaches renewal.",
            omRenewalAlertsEnabled(license.fm, license.file),
          ],
          [
            "billing",
            "Billing_Alerts",
            "Upcoming charges",
            "Notify before a scheduled billing charge.",
            omBillingAlertsEnabled(license.fm, license.file),
          ],
          ...(seatBased
            ? [
                [
                  "seats",
                  "Seat_Alerts",
                  "Low seat availability",
                  "Notify when 5 or fewer purchased seats remain.",
                  omSeatAlertsEnabled(license.fm, license.file),
                ],
              ]
            : []),
        ]) {
          new Setting(alerts)
            .setName(label)
            .setDesc(description)
            .addToggle((t) =>
              t.setValue(enabled).onChange(async (value) => {
                try {
                  const path = license.file.path;
                  this.settings.licenseAlerts = {
                    ...(this.settings.licenseAlerts || {}),
                    [path]: { ...(this.settings.licenseAlerts?.[path] || {}), [kind]: value },
                  };
                  await this.saveData(this.settings);
                  // Tidy up the property older versions wrote to the note.
                  if (Object.prototype.hasOwnProperty.call(license.fm, field))
                    await this.app.fileManager.processFrontMatter(license.file, (fm) => {
                      delete fm[field];
                    });
                  this.refresh();
                  new Notice(label + (value ? " enabled" : " disabled"));
                } catch (error) {
                  new Notice("Unable to update alert: " + error.message);
                }
              }),
            );
        }
        const customize = alerts.createEl("button", { text: "Customize alert timing & colors" });
        customize.onclick = () => new OMAlertRulesModal(this.app, this, license).open();
        const billing = body.createDiv({ cls: "ops-settings-group" });
        billing.createEl("h4", { text: "Billing & renewals" });
        const current = omBillingCycle(license.fm),
          next = iso(license.fm.Next_Billing),
          renewal = omMonthToMonth(license.fm, license.file) ? null : inferredRenewal(license.fm);
        billing.createEl("p", {
          cls: "ops-settings-facts",
          text:
            "Billing: " +
            (current?.label ?? "Not set") +
            " · Next charge: " +
            (next ?? "Not set") +
            " · Renewal: " +
            (renewal?.date ?? "Not set"),
        });
        const lastCharge = this.lastHistory(license, "billing"),
          lastRenewal = this.lastHistory(license, "renewal");
        if (lastCharge || lastRenewal)
          billing.createEl("p", {
            cls: "ops-settings-facts",
            text: [
              lastCharge ? "Last charge recorded: " + lastCharge : "",
              lastRenewal ? "Last renewed: " + lastRenewal : "",
            ]
              .filter(Boolean)
              .join(" · "),
          });
        if (charge !== null)
          billing.createEl("p", {
            cls: "ops-settings-facts",
            text: "Expected charge: " + omUSD(charge) + " / " + frequency,
          });
        const schedule = billing.createEl("button", { text: "Edit billing schedule" });
        schedule.onclick = () => new OMBillingModal(this.app, this, license).open();
        const contractDate = billing.createEl("button", {
          text: renewal ? "Edit renewal date" : "Set contract renewal date",
        });
        contractDate.onclick = () => new OMContractDateModal(this.app, this, license).open();
        if (seatBased || omIsBundle(license)) {
          const seats = body.createDiv({ cls: "ops-settings-group" });
          seats.createEl("h4", { text: "Seats & assignments" });
          if (hasTypes || omIsBundle(license)) {
            if (hasTypes) {
              seats.createEl("p", {
                cls: "ops-settings-facts",
                text:
                  multi.length +
                  " license types · " +
                  bought +
                  " purchased · " +
                  used +
                  " assigned · " +
                  available +
                  " available · Calculated subtotal: " +
                  omUSD(summary.total) +
                  " (not recorded Cost)",
              });
              for (const t of multi)
                seats.createEl("p", {
                  cls: "ops-settings-facts",
                  text: t.name + " → " + (t.field || "Nextiva_Licenses"),
                });
            } else
              seats.createEl("p", {
                cls: "ops-settings-facts",
                text: "No license types configured yet. Add each license type, its monthly price, and purchased seat count.",
              });
            const typeBtn = seats.createEl("button", {
              text: hasTypes ? "Manage license types" : "Add license types",
            });
            typeBtn.onclick = () => new OMNextivaTypesModal(this.app, this, license).open();
            if (hasTypes) {
              const employeeBtn = seats.createEl("button", { text: "Assign employee licenses" });
              employeeBtn.onclick = () => new OMNextivaAssignmentsModal(this.app, this, license).open();
            }
          } else {
            seats.createEl("p", {
              cls: "ops-settings-facts",
              text:
                bought === null
                  ? "Purchased seats not set"
                  : bought + " purchased · " + used + " assigned · " + available + " available",
            });
            const set = seats.createEl("button", { text: "Set purchased seats" });
            set.onclick = () => new OMSeatCountModal(this.app, this, license).open();
            seats.createEl("p", {
              cls: "ops-settings-facts",
              text: "Employee property: " + this.licenseField(license),
            });
            const fieldButton = seats.createEl("button", { text: "Choose employee property" });
            fieldButton.onclick = () => new OMLicenseFieldModal(this.app, this, license).open();
            const manage = seats.createEl("button", { text: "Manage employees" });
            manage.onclick = () => new OMEmployeeLicenseModal(this.app, this, license, inv.employees).open();
          }
        }
      }
    };
    search.addEventListener("input", show);
    show();
  }

  renderSeatAvailability(root, records) {
    const seats = records.filter((r) => {
      const raw = r.fm.Agreement_Types ?? r.fm.Agreement_Type ?? r.fm.License_Type ?? [];
      const types = (Array.isArray(raw) ? raw : String(raw).split(",")).map((v) =>
        String(v).trim().toLowerCase(),
      );
      return (
        omNextivaTypes(r.fm).length > 0 ||
        types.some((t) => ["software seats", "seat", "software", "per-user"].includes(t)) ||
        this.licenseInventory(r).total !== null
      );
    });
    if (!seats.length) return;
    root.createEl("h2", { cls: "ops-equipment-heading", text: "License Overview" });
    const list = root.createDiv({ cls: "ops-license-overviews" });
    for (const r of seats) {
      const types = omNextivaTypes(r.fm);
      const multi = types.length > 0;
      const inv = this.licenseInventory(r);
      const breakdown = multi ? omNextivaSummary(this, r) : null;
      const purchased = multi ? breakdown.rows.reduce((n, x) => n + x.purchased, 0) : inv.total;
      const assigned = multi ? breakdown.rows.reduce((n, x) => n + x.assigned, 0) : inv.assigned;
      const available = purchased === null ? null : purchased - assigned;
      const alerts = omSeatAlertsEnabled(r.fm, r.file);
      const severity =
        alerts && available !== null
          ? available <= 2
            ? " ops-seat-warning"
            : available >= 3 && available <= 5
              ? " ops-seat-moderate"
              : ""
          : "";
      const card = list.createDiv({ cls: "ops-license-overview ops-attention-item" + severity });
      const info = card.createDiv({ cls: "ops-item-info" });
      const title = info.createEl("button", { cls: "ops-item-title", text: r.file.basename });
      title.onclick = () => this.app.workspace.getLeaf(false).openFile(r.file);
      const billing = omBillingCycle(r.fm);
      const total = omRenewalEstimate(r.fm)?.amount ?? null;
      const perSeat = total !== null && purchased !== null && purchased > 0 ? total / purchased : null;
      info.createEl("small", {
        text:
          purchased === null
            ? "Purchased seats unknown"
            : purchased + " purchased · " + assigned + " assigned · " + available + " available",
      });
      info.createEl("small", {
        text:
          "Billing total: " +
          (total === null ? "Not set" : omUSD(total)) +
          " / " +
          (billing?.label?.toLowerCase() ?? "billing period") +
          (perSeat !== null ? " · Derived: " + omUSD(perSeat) + " / purchased seat" : ""),
      });
      if (multi) {
        const details = info.createEl("details", { cls: "ops-license-overview-details" });
        const key = "licenses:overview:types:" + r.file.path;
        details.open = this.settings.collapsedSections?.[key] === true;
        details.addEventListener("toggle", () => {
          if (this.settings.collapsedSections?.[key] === details.open) return;
          this.settings.collapsedSections = {
            ...(this.settings.collapsedSections || {}),
            [key]: details.open,
          };
          this.saveData(this.settings);
        });
        details.createEl("summary", { text: "License types (" + breakdown.rows.length + ")" });
        const table = details.createEl("table");
        const header = table.createEl("tr");
        for (const label of ["Type", "Purchased", "Assigned", "Available", "Rate", "Monthly total"])
          header.createEl("th", { text: label });
        for (const item of breakdown.rows) {
          const tr = table.createEl("tr");
          for (const val of [
            item.name,
            item.purchased,
            item.assigned,
            item.available,
            omUSD(item.rate),
            omUSD(item.cost),
          ])
            tr.createEl("td", { text: String(val) });
          if (item.holders.length) {
            const peopleRow = table.createEl("tr");
            const peopleCell = peopleRow.createEl("td");
            peopleCell.colSpan = 6;
            peopleCell.createSpan({ text: "Assigned: " });
            for (const [index, employee] of item.holders.entries()) {
              if (index) peopleCell.createSpan({ text: ", " });
              const link = peopleCell.createEl("a", {
                cls: "internal-link",
                text: employee.file.basename,
                href: "#",
              });
              link.onclick = (e) => {
                e.preventDefault();
                this.app.workspace.getLeaf(false).openFile(employee.file);
              };
            }
          }
        }
        const assign = info.createEl("button", { text: "Manage type assignments" });
        assign.onclick = () => new OMNextivaAssignmentsModal(this.app, this, r).open();
      } else {
        const linked = inv.employees.filter((e) =>
          Object.entries(e.fm).some(
            ([key, value]) => key.toLowerCase() === inv.field.toLowerCase() && omLicenseAssigned(value),
          ),
        );
        const details = info.createEl("details", { cls: "ops-license-overview-details" });
        const key = "licenses:overview:people:" + r.file.path;
        details.open = this.settings.collapsedSections?.[key] === true;
        details.addEventListener("toggle", () => {
          if (this.settings.collapsedSections?.[key] === details.open) return;
          this.settings.collapsedSections = {
            ...(this.settings.collapsedSections || {}),
            [key]: details.open,
          };
          this.saveData(this.settings);
        });
        details.createEl("summary", { text: "Assigned employees (" + linked.length + ")" });
        for (const employee of linked) {
          const link = details.createEl("button", { text: employee.file.basename });
          link.onclick = () => this.app.workspace.getLeaf(false).openFile(employee.file);
        }
        if (!linked.length) details.createEl("small", { text: "No assigned employees." });
        const assign = info.createEl("button", { text: "Manage employee assignments" });
        assign.onclick = () => new OMEmployeeLicenseModal(this.app, this, r, inv.employees).open();
      }
    }
  }
  renderEquipmentPhases(root, records, mode = "phase") {
    const dashboard = mode === "phase";
    const heading = root.createEl(dashboard ? "h2" : "h3", {
      text: dashboard ? "Equipment by Phase" : "Equipment by Status",
    });
    heading.addClass("ops-equipment-heading");
    let sortBy = "filename";
    const controls = root.createDiv({ cls: "ops-equipment-controls" });
    if (!dashboard) {
      new Setting(controls).setName("Sort devices by").addDropdown((d) =>
        d
          .addOption("filename", "Filename")
          .addOption("employee", "Employee name")
          .setValue(sortBy)
          .onChange((v) => {
            sortBy = v;
            draw();
          }),
      );
    }
    const area = root.createDiv({ cls: "ops-equipment-groups" });
    const employeeName = (r) => {
      const v = r.fm.Employee ?? r.fm.Assigned_To ?? "";
      return Array.isArray(v) ? v.map(String).join(", ") : String(v);
    };
    const populated = (v) => v !== undefined && v !== null && String(v).trim() !== "";
    const showValue = (v) =>
      Array.isArray(v)
        ? v.map(String).join(", ")
        : typeof v === "object" && v !== null
          ? JSON.stringify(v)
          : String(v);
    const draw = () => {
      area.empty();
      const groups = new Map();
      for (const r of records) {
        const raw = dashboard ? (r.fm.Approved ?? r.fm.approved) : (r.fm.Status ?? r.fm.status);
        let group = populated(raw) ? String(raw).trim() : dashboard ? "Unassigned phase" : "No status";
        if (dashboard) {
          const match = group.match(/Phase\b.*$/i);
          if (match) group = match[0];
        }
        if (!groups.has(group)) groups.set(group, []);
        groups.get(group).push(r);
      }
      const names = [...groups.keys()].sort((a, b) =>
        !dashboard && a.toLowerCase() === "complete"
          ? 1
          : !dashboard && b.toLowerCase() === "complete"
            ? -1
            : a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }),
      );
      if (!names.length) area.createEl("p", { text: "No equipment notes found." });
      for (const group of names) {
        const items = groups.get(group);
        const section = area.createEl("details", { cls: "ops-equipment-phase" });
        const collapseKey = mode + ":" + group;
        section.open = this.settings.collapsedSections?.[collapseKey] === true;
        section.addEventListener("toggle", () => {
          const next = section.open;
          if (this.settings.collapsedSections?.[collapseKey] === next) return;
          this.settings.collapsedSections = {
            ...(this.settings.collapsedSections || {}),
            [collapseKey]: next,
          };
          this.saveData(this.settings);
        });
        section.createEl("summary", { text: group + " (" + items.length + ")" });
        const list = section.createDiv({ cls: "ops-equipment-list" });
        const ordered = [...items].sort((a, b) =>
          sortBy === "employee"
            ? employeeName(a).localeCompare(employeeName(b), undefined, {
                numeric: true,
                sensitivity: "base",
              }) ||
              a.file.basename.localeCompare(b.file.basename, undefined, {
                numeric: true,
                sensitivity: "base",
              })
            : a.file.basename.localeCompare(b.file.basename, undefined, {
                numeric: true,
                sensitivity: "base",
              }),
        );
        for (const r of ordered) {
          const card = list.createDiv({ cls: "ops-attention-item ops-equipment-card" });
          this.renderCardIcon(card, "device");
          const info = card.createDiv({ cls: "ops-item-info" });
          omCardTitle(info, r, this.app);
          const addLinkedText = (parent, value) => {
            const parts = showValue(value).split(/(\[\[[^\]]+\]\])/g);
            for (const part of parts) {
              const match = part.match(/^\[\[([^\]]+)\]\]$/);
              if (!match) {
                parent.createSpan({ text: part });
                continue;
              }
              const [target, alias] = match[1].split("|");
              const label = alias || target;
              const link = parent.createEl("a", { text: label, cls: "internal-link" });
              link.href = "#";
              link.onclick = (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.app.workspace.openLinkText(target, r.file.path, false);
              };
            }
          };
          const summary = info.createEl("small");
          addLinkedText(summary, employeeName(r) || "No employee");
          summary.createSpan({
            text:
              " · " +
              (populated(r.fm.Ordered ?? r.fm.Order_Date) ? "Ordered" : "Not ordered") +
              " · " +
              (populated(r.fm.Set_Up ?? r.fm.Set_up) ? "Set up" : "Not set up"),
          });
          const note = r.fm.Notes;
          if (populated(note)) {
            const line = info.createDiv({ cls: "ops-equipment-note" });
            line.createSpan({ text: "Notes: " });
            addLinkedText(line, note);
          }
          omRenderRelated(info, r.fm, this.app, r.file.path);
          const replacing = r.fm.To_Replace;
          if (populated(replacing)) {
            const line = info.createDiv({ cls: "ops-equipment-replace" });
            line.createSpan({ text: "To replace: " });
            addLinkedText(line, replacing);
          }
          const edit = card.createEl("button", { text: "Update", cls: "ops-item-action" });
          edit.onclick = () => new OMDeviceWorkflowModal(this.app, this, r).open();
        }
      }
    };
    draw();
  }

  linkedTo(fm, project) {
    return omRelated(fm, "Operations_Projects").some((x) => omMatches(x, project));
  }
  renderProjects(root) {
    const projects = this.records().filter((r) => r.kind === "project");
    const section = root.createDiv({ cls: "ops-projects" });
    section.createEl("h3", { text: "Master projects & linked records" });
    if (!projects.length)
      section.createEl("p", {
        text: "Create a master project, then link existing device, order, and license notes without moving or duplicating them.",
      });
    for (const p of projects) {
      const related = this.records().filter(
        (r) => r.file.path !== p.file.path && this.linkedTo(r.fm, p.file),
      );
      const card = section.createDiv({ cls: "ops-project-card" });
      const heading = card.createDiv({ cls: "ops-project-head" });
      const open = heading.createEl("button", { text: p.file.basename });
      open.onclick = () => this.app.workspace.getLeaf(false).openFile(p.file);
      heading.createSpan({ text: related.length + " linked records" });
      const add = card.createEl("button", { text: "Manage links" });
      add.onclick = () => new OMLinkModal(this.app, this, p.file).open();
      const list = card.createDiv({ cls: "ops-project-links" });
      for (const r of related) {
        const link = list.createEl("button", { text: r.file.basename + " · " + r.kind });
        link.onclick = () => this.app.workspace.getLeaf(false).openFile(r.file);
      }
      if (!related.length) list.createEl("small", { text: "No linked devices, licenses, or orders yet." });
    }
  }
  exportNoteTags(file, fm) {
    return tagSet(fm || {});
  }
  exportTagNames() {
    const tags = new Set();
    for (const file of this.app.vault.getMarkdownFiles()) {
      const fm = this.app.metadataCache.getFileCache(file)?.frontmatter || {};
      for (const tag of this.exportNoteTags(file, fm)) tags.add(tag);
    }
    return [...tags].sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
  }
  exportSourceNotes(selection) {
    if (selection === "tracked") return this.records();
    const results = [];
    for (const file of this.app.vault.getMarkdownFiles()) {
      const fm = this.app.metadataCache.getFileCache(file)?.frontmatter || {};
      if (selection.startsWith("tag:") && !this.exportNoteTags(file, fm).has(selection.slice(4))) continue;
      results.push({ file, fm, kind: classified(file, fm) || "note" });
    }
    return results.sort(
      (a, b) =>
        a.file.basename.localeCompare(b.file.basename, undefined, { numeric: true, sensitivity: "base" }) ||
        a.file.path.localeCompare(b.file.path),
    );
  }
  exportCellValue(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) return dateISO(value);
    if (typeof value === "object") return JSON.stringify(value);
    return String(value);
  }
  async exportXLSX(rows, { selection = "tracked", allProperties = false } = {}) {
    let cols, data;
    if (allProperties) {
      const keys = new Set();
      for (const r of rows)
        for (const key of Object.keys(r.fm || {})) if (key !== "position" && key !== "---") keys.add(key);
      const sorted = [...keys].sort((a, b) =>
        a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }),
      );
      cols = [
        "Name",
        "Note_Path",
        "Tags",
        ...sorted.filter((key) => !["Name", "Note_Path", "Tags"].includes(key)),
      ];
      data = rows.map((r) => {
        const tags = [...this.exportNoteTags(r.file, r.fm)]
          .map((x) => "#" + x)
          .sort()
          .join(", ");
        return [
          r.file.basename,
          r.file.path,
          tags,
          ...cols.slice(3).map((key) => this.exportCellValue(r.fm?.[key])),
        ];
      });
    } else {
      cols = [
        "Type",
        "Name",
        "Status",
        "Approved",
        "Vendor",
        "Employee",
        "Serial",
        "Model",
        "Cost",
        "Ordered",
        "Set_Up",
        "Old_Device_Returned",
        "Next_Renewal",
        "Due_Date",
        "Project",
        "Linked_Projects",
        "Assigned_Licenses",
        "Note_Path",
      ];
      data = rows.map((r) => {
        const f = r.fm || {};
        return [
          omTypeLabel(r),
          r.file.basename,
          f.Status ?? f.status ?? "",
          f.Approved,
          f.Vendor,
          f.Employee ?? f.Assigned_To,
          f.Serial_Number ?? f.Serial,
          f.Device_Model ?? f.Model,
          f.Cost,
          f.Ordered ?? f.Order_Date,
          f.Set_Up ?? f.Set_up ?? f.Received_Date,
          f.Old_Device_Returned,
          inferredRenewal(f)?.date ?? "",
          f.Due_Date,
          f.Project ?? f.Parent_Project,
          omRelated(f, "Operations_Projects").map(omLinkValue).join("; "),
          omRelated(f, "Assigned_Licenses").map(omLinkValue).join("; "),
          r.file.path,
        ];
      });
    }
    const folder = this.settings.folder;
    const name =
      selection === "tracked"
        ? "Operations Records"
        : selection === "all"
          ? "All Vault Notes"
          : "Tag " + selection.slice(4);
    const safe = name.replace(/[\\/:*?"<>|]/g, "-").slice(0, 90);
    const base = folder + "/" + safe + " " + dateISO(new Date());
    let path = normalizePath(base + ".xlsx");
    let i = 2;
    while (this.app.vault.getAbstractFileByPath(path)) path = normalizePath(base + " (" + i++ + ").xlsx");
    try {
      if (!this.app.vault.getAbstractFileByPath(folder)) await this.app.vault.createFolder(folder);
      const bytes = xlsxWorkbook(cols, data);
      const binary = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
      await this.app.vault.createBinary(path, binary);
      new Notice("Exported " + rows.length + " notes to " + path);
      return true;
    } catch (e) {
      new Notice("Excel export failed: " + e.message);
      return false;
    }
  }

  async activate() {
    let leaf = this.app.workspace.getLeavesOfType(VIEW)[0];
    if (!leaf) {
      leaf = this.app.workspace.getRightLeaf(false);
      await leaf.setViewState({ type: VIEW, active: true });
    }
    this.app.workspace.revealLeaf(leaf);
  }
  refresh() {
    for (const embedded of this.embeddedDashboards) {
      Promise.resolve(embedded.render()).catch((err) =>
        console.error("Operations Manager embedded dashboard refresh failed", err),
      );
    }
    for (const leaf of this.app.workspace.getLeavesOfType(VIEW)) {
      const view = leaf.view;
      const el = view.contentEl;
      const scroll = el.querySelector(".view-content") || el;
      const position = scroll.scrollTop;
      Promise.resolve(view.render()).then(() => {
        scroll.scrollTop = position;
      });
    }
  }
};
