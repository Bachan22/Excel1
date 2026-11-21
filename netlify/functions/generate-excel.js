const ExcelJS = require("exceljs");

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function sanitizeText(val) {
  if (val === null || val === undefined) return "";
  let str = String(val).trim();
  if (str.length > 200) {
    str = str.substring(0, 200);
  }
  // Formula injection prevention: prefix with single quote if starting with =, +, -, or @
  if (/^[=+\-@]/.test(str)) {
    str = "'" + str;
  }
  return str;
}

function sanitizeFileName(fileName, date) {
  let name = (fileName || "").replace(/[\\/:*?"<>|]/g, "").trim();
  if (!name) {
    const dateStr = date ? String(date).replace(/\//g, "-").trim() : "export";
    name = `Protons-Widecare-${dateStr}`;
  }
  name = name.replace(/\.xlsx$/i, "");
  return name + ".xlsx";
}

function normalizeColor(hex, fallback) {
  const clean = (hex || "").replace(/^#/, "").trim().toUpperCase();
  return /^[0-9A-F]{6}$/.test(clean) ? clean : fallback;
}

exports.handler = async function (event) {
  // 1. Allow POST only (405 otherwise)
  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      headers: {
        "Content-Type": "application/json",
        "Allow": "POST"
      },
      body: JSON.stringify({ error: "Method not allowed. Only POST is supported." })
    };
  }

  // 2. Limit payload size to ~1MB
  if (event.body && Buffer.byteLength(event.body, "utf8") > 1024 * 1024) {
    return {
      statusCode: 400,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Payload too large. Maximum size is 1MB." })
    };
  }

  // 3. Reject invalid JSON (400)
  let data;
  try {
    data = JSON.parse(event.body || "{}");
  } catch (err) {
    return {
      statusCode: 400,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Invalid JSON body." })
    };
  }

  if (!data || typeof data !== "object") {
    return {
      statusCode: 400,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Request body must be a JSON object." })
    };
  }

  const {
    title = "Protons Engineering",
    company = "Company:Widecare",
    sheetName = "Client Data",
    date = "",
    fileName = "",
    colors = {},
    columns = [],
    rows = [],
    totals = []
  } = data;

  // 4. Validate columns: max 12 columns, at least 1
  if (!Array.isArray(columns) || columns.length === 0) {
    return {
      statusCode: 400,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Columns must be a non-empty array." })
    };
  }
  if (columns.length > 12) {
    return {
      statusCode: 400,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Too many columns. Maximum allowed is 12 columns." })
    };
  }

  // 5. Validate rows: max 500 rows
  if (!Array.isArray(rows)) {
    return {
      statusCode: 400,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Rows must be an array." })
    };
  }
  if (rows.length > 500) {
    return {
      statusCode: 400,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Too many rows. Maximum allowed is 500 rows." })
    };
  }

  try {
    const finalFileName = sanitizeFileName(fileName, date);
    const titleColor = normalizeColor(colors.title, "D35400");
    const companyColor = normalizeColor(colors.company, "E67E22");
    const headerColor = normalizeColor(colors.header, "B03A2E");

    const n = columns.length;
    const last = colLetter(n);
    const fill = (hex) => ({ type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } });
    const thin = { style: "thin", color: { argb: "FF999999" } };
    const border = { top: thin, left: thin, bottom: thin, right: thin };

    const wb = new ExcelJS.Workbook();
    const cleanSheetName = (sheetName || "Sheet1").toString().slice(0, 31).replace(/[\\/?*[\]]/g, "");
    const ws = wb.addWorksheet(cleanSheetName || "Sheet1");

    ws.columns = columns.map((c) => ({
      width: typeof c.width === "number" && c.width > 0 ? c.width : 20
    }));

    // Row 1: Merged title across all columns, size 20 bold white, centred
    ws.mergeCells("A1:" + last + "1");
    const titleCell = ws.getCell("A1");
    titleCell.value = sanitizeText(title);
    titleCell.font = { name: "Calibri", size: 20, bold: true, color: { argb: "FFFFFFFF" } };
    titleCell.fill = fill(titleColor);
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(1).height = 30;

    // Row 2: Company on left half and "DATE:- dd/mm/yyyy" on right half
    const split = Math.max(1, Math.ceil(n / 2));
    const style2 = (cell, align) => {
      cell.font = { name: "Calibri", size: 16, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = fill(companyColor);
      cell.alignment = { horizontal: align, vertical: "middle" };
    };

    const dateDisplay = date ? ("DATE:- " + sanitizeText(date)) : "DATE:- ";
    if (n >= 2) {
      ws.mergeCells("A2:" + colLetter(split) + "2");
      ws.mergeCells(colLetter(split + 1) + "2:" + last + "2");
      const c1 = ws.getCell("A2");
      const c2 = ws.getCell(colLetter(split + 1) + "2");
      c1.value = sanitizeText(company);
      style2(c1, "left");
      c2.value = dateDisplay;
      style2(c2, "right");
    } else {
      const c1 = ws.getCell("A2");
      c1.value = sanitizeText(company) + "   " + dateDisplay;
      style2(c1, "left");
    }
    ws.getRow(2).height = 25;

    // Row 3: Headings, bold white on header colour, centred, thin borders
    const head = ws.getRow(3);
    columns.forEach((c, i) => {
      const cell = head.getCell(i + 1);
      cell.value = sanitizeText(c.key);
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = fill(headerColor);
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = border;
    });
    head.height = 22;

    // Data rows (saved as text numFmt "@" so phone numbers and IDs stay as typed, left aligned, thin borders)
    rows.forEach((r, ri) => {
      const row = ws.getRow(4 + ri);
      columns.forEach((c, i) => {
        const cell = row.getCell(i + 1);
        cell.value = sanitizeText(r[c.key]);
        cell.numFmt = "@";
        cell.font = { name: "Calibri", size: 11 };
        cell.alignment = { horizontal: "left", vertical: "middle" };
        cell.border = border;
      });
    });

    // Totals: blank row, then totals (label in column D, value in column E, bold, light orange label fill FDE8D6)
    if (Array.isArray(totals) && totals.length > 0) {
      const labelCol = n >= 5 ? 4 : 1;
      let rowNo = 4 + rows.length + 1; // leave 1 blank row
      totals.forEach((it) => {
        const row = ws.getRow(rowNo++);
        const a = row.getCell(labelCol);
        const b = row.getCell(labelCol + 1);
        a.value = sanitizeText(it.label);
        b.value = typeof it.value === "number" ? it.value : sanitizeText(it.value);
        a.font = { name: "Calibri", size: 11, bold: true };
        b.font = { name: "Calibri", size: 11, bold: true };
        a.alignment = { horizontal: "left", vertical: "middle" };
        b.alignment = { horizontal: "left", vertical: "middle" };
        a.border = border;
        b.border = border;
        a.fill = fill("FDE8D6");
      });
    }

    const buffer = await wb.xlsx.writeBuffer();
    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${finalFileName}"`,
        "Cache-Control": "no-store"
      },
      body: Buffer.from(buffer).toString("base64"),
      isBase64Encoded: true
    };
  } catch (err) {
    // Return clear JSON error, never return a stack trace
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "Failed to generate Excel file." })
    };
  }
};
