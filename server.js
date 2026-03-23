import express from "express";
import multer from "multer";
import XLSX from "xlsx";
import archiver from "archiver";
import { PDFDocument, cmyk } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const host = normalizeCellValue(process.env.HOST) || "0.0.0.0";
const port = Number(process.env.PORT) || 3010;
const trustProxy = normalizeCellValue(process.env.TRUST_PROXY).toLowerCase() === "true";
const jsonBodyLimit = normalizeCellValue(process.env.JSON_BODY_LIMIT) || "10mb";
const uploadFileSizeMb = Math.max(Number(process.env.UPLOAD_FILE_SIZE_MB) || 20, 1);
const uploadFileSizeBytes = uploadFileSizeMb * 1024 * 1024;

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: uploadFileSizeBytes
  }
});

const templatesDir = path.join(__dirname, "templates");
const publicDir = path.join(__dirname, "public");
const defaultRegularFontPath =
  normalizeCellValue(process.env.DEFAULT_REGULAR_FONT_PATH) || "assets/fonts/Poppins-Regular.ttf";
const defaultBoldFontPath =
  normalizeCellValue(process.env.DEFAULT_BOLD_FONT_PATH) || "assets/fonts/Poppins-Bold.ttf";
const FEISHU_APP_ID = normalizeCellValue(process.env.FEISHU_APP_ID);
const FEISHU_APP_SECRET = normalizeCellValue(process.env.FEISHU_APP_SECRET);
const FEISHU_BASE_URL = "https://open.feishu.cn";
const FEISHU_CONFIRMATION_TEXT = "请确认名片内容有没有问题，如需修改请直接回复。";
const FEISHU_EMAIL_COLUMNS = ["E-Mail Address", "邮箱", "E-mail"];
const FEISHU_MOBILE_COLUMNS = ["Mobile Number", "手机", "Mobile"];
const allowedFrameAncestors =
  normalizeCellValue(process.env.ALLOWED_FRAME_ANCESTORS) ||
  "'self' https://*.feishu.cn https://*.larksuite.com https://*.feishu-pre.cn";
let feishuTokenCache = {
  token: "",
  expiresAt: 0
};

if (trustProxy) {
  app.set("trust proxy", true);
}

app.disable("x-powered-by");

app.use((req, res, next) => {
  res.setHeader(
    "Content-Security-Policy",
    [
      "default-src 'self' data: blob:",
      "connect-src 'self' https://open.feishu.cn",
      "img-src 'self' data: blob:",
      "font-src 'self' data:",
      "style-src 'self' 'unsafe-inline'",
      "script-src 'self' 'unsafe-inline' blob:",
      "worker-src 'self' blob:",
      `frame-ancestors ${allowedFrameAncestors}`
    ].join("; ")
  );
  res.setHeader("Referrer-Policy", "strict-origin-when-cross-origin");
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=()");
  next();
});

app.get("/healthz", (_req, res) => {
  res.json({
    ok: true,
    service: "business-card-web"
  });
});

app.use(express.json({ limit: jsonBodyLimit }));
app.use(express.static(publicDir));

app.get("/vendor/pdfjs/pdf.core.mjs", (req, res) => {
  res.sendFile(path.join(__dirname, "node_modules", "pdfjs-dist", "build", "pdf.mjs"));
});

app.get("/vendor/pdfjs/pdf.mjs", (req, res) => {
  res.sendFile(path.join(__dirname, "node_modules", "pdfjs-dist", "build", "pdf.mjs"));
});

app.get("/vendor/pdfjs/pdf.worker.core.mjs", (req, res) => {
  res.sendFile(path.join(__dirname, "node_modules", "pdfjs-dist", "build", "pdf.worker.mjs"));
});

app.get("/vendor/pdfjs/pdf.worker.mjs", (req, res) => {
  res.sendFile(path.join(__dirname, "node_modules", "pdfjs-dist", "build", "pdf.worker.mjs"));
});

function normalizeCellValue(value) {
  if (value === undefined || value === null) {
    return "";
  }

  return String(value).trim();
}

function stripHtmlTags(value) {
  return normalizeCellValue(value)
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/\s+/g, " ")
    .trim();
}

function extractRichTextValue(value) {
  const text = normalizeCellValue(value);
  if (!text) {
    return "";
  }

  const matches = Array.from(text.matchAll(/<t[^>]*>(.*?)<\/t>/g));
  if (!matches.length) {
    return "";
  }

  return matches
    .map((match) => normalizeCellValue(match[1]))
    .filter(Boolean)
    .join("");
}

function extractWorksheetCellValue(cell) {
  if (!cell) {
    return "";
  }

  const numericValue =
    cell.t === "n" && Number.isFinite(cell.v) ? String(cell.v) : "";
  const scientificDisplay =
    typeof cell.w === "string" && /^[+-]?\d+(?:\.\d+)?E[+-]?\d+$/i.test(cell.w.trim());

  const candidates = [
    cell.l?.display,
    scientificDisplay ? numericValue : "",
    cell.w,
    extractRichTextValue(cell.r),
    stripHtmlTags(cell.h),
    cell.v,
    typeof cell.l?.Target === "string" ? cell.l.Target.replace(/^mailto:/i, "") : ""
  ];

  for (const candidate of candidates) {
    const normalized = normalizeCellValue(candidate);
    if (normalized) {
      return normalized;
    }
  }

  return "";
}

function buildWorksheetHeaders(worksheet, range) {
  const headers = [];
  let emptyCount = 0;

  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const address = XLSX.utils.encode_cell({ r: range.s.r, c: col });
    const headerValue = extractWorksheetCellValue(worksheet[address]);
    if (headerValue) {
      headers.push(headerValue);
      continue;
    }

    headers.push(emptyCount === 0 ? "__EMPTY" : `__EMPTY_${emptyCount}`);
    emptyCount += 1;
  }

  return headers;
}

function worksheetToRows(worksheet) {
  const ref = worksheet?.["!ref"];
  if (!ref) {
    return [];
  }

  const range = XLSX.utils.decode_range(ref);
  const headers = buildWorksheetHeaders(worksheet, range);
  const rows = [];

  for (let rowIndex = range.s.r + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const row = {};
    let hasValue = false;

    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const header = headers[col - range.s.c];
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: col });
      const value = extractWorksheetCellValue(worksheet[address]);
      row[header] = value;
      if (value) {
        hasValue = true;
      }
    }

    if (hasValue) {
      rows.push(row);
    }
  }

  return rows;
}

function sanitizeTextForPdf(value) {
  return normalizeCellValue(value)
    .replace(/➕/g, "+")
    .replace(/＋/g, "+")
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/[–—]/g, "-")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\u00A0/g, " ");
}

function normalizeMultilineValue(value) {
  return sanitizeTextForPdf(value).replace(/\r\n?/g, "\n");
}

function getFieldExcelColumns(field) {
  if (field.excelColumn) {
    return [field.excelColumn];
  }

  if (Array.isArray(field.excelColumns)) {
    return field.excelColumns.filter(Boolean);
  }

  return [];
}

function getFieldPositionKey(field, index) {
  if (field.positionKey) {
    return field.positionKey;
  }

  if (field.excelColumn) {
    return field.excelColumn;
  }

  return `static-${index}`;
}

function getOptionalFieldColumns(field) {
  if (!Array.isArray(field.optionalExcelColumns)) {
    return new Set();
  }

  return new Set(field.optionalExcelColumns.filter(Boolean));
}

function getRequiredColumns(templateConfig) {
  const requiredColumns = new Set([templateConfig.fileNameField]);

  for (const field of templateConfig.fields || []) {
    const optionalColumns = getOptionalFieldColumns(field);
    for (const column of getFieldExcelColumns(field)) {
      if (!optionalColumns.has(column)) {
        requiredColumns.add(column);
      }
    }
  }

  return [...requiredColumns].filter(Boolean);
}

function getTemplatePositionKeys(templateConfig) {
  const keys = new Set();

  (templateConfig.fields || []).forEach((field, index) => {
    if (field.kind === "contact-block") {
      const blockKey = getFieldPositionKey(field, index);
      getFieldExcelColumns(field).forEach((column) => {
        keys.add(column);
        keys.add(`${blockKey}:${column}:label`);
      });
      return;
    }

    keys.add(getFieldPositionKey(field, index));
  });

  return keys;
}

function breakLongWord(word, font, size, maxWidth) {
  const parts = [];
  let current = "";

  for (const char of word) {
    const candidate = current + char;
    if (candidate && font.widthOfTextAtSize(candidate, size) <= maxWidth) {
      current = candidate;
      continue;
    }

    if (current) {
      parts.push(current);
      current = char;
    } else {
      parts.push(char);
      current = "";
    }
  }

  if (current) {
    parts.push(current);
  }

  return parts;
}

function tokenizeTextForWrap(text, font, size, maxWidth) {
  const rawTokens = sanitizeTextForPdf(text)
    .replace(/\s+/g, " ")
    .trim()
    .split(/\s+/)
    .map((token) => token.trim())
    .filter(Boolean);

  const tokens = [];
  for (const token of rawTokens) {
    if (font.widthOfTextAtSize(token, size) <= maxWidth) {
      tokens.push(token);
      continue;
    }

    tokens.push(...breakLongWord(token, font, size, maxWidth));
  }

  return tokens;
}

function lineWidthForTokens(tokens, font, size) {
  if (!tokens.length) {
    return 0;
  }

  const wordsWidth = tokens.reduce(
    (sum, token) => sum + font.widthOfTextAtSize(token, size),
    0
  );
  const spacesWidth = font.widthOfTextAtSize(" ", size) * (tokens.length - 1);
  return wordsWidth + spacesWidth;
}

function wrapTextGreedy(text, font, size, maxWidth) {
  const cleaned = sanitizeTextForPdf(text);
  if (!cleaned.trim()) {
    return [];
  }

  const tokens = tokenizeTextForWrap(cleaned, font, size, maxWidth);
  const lines = [];
  let currentLineTokens = [];

  for (const token of tokens) {
    const candidateTokens = [...currentLineTokens, token];
    if (
      !currentLineTokens.length ||
      lineWidthForTokens(candidateTokens, font, size) <= maxWidth
    ) {
      currentLineTokens = candidateTokens;
      continue;
    }

    lines.push(currentLineTokens.join(" "));
    currentLineTokens = [token];
  }

  if (currentLineTokens.length) {
    lines.push(currentLineTokens.join(" "));
  }

  return lines;
}

function roundToTenths(value) {
  return Math.round(value * 10) / 10;
}

function estimateTextBlockHeight(lineCount, size, lineGap) {
  if (lineCount <= 0) {
    return 0;
  }

  return size + (lineCount - 1) * lineGap;
}

function resolveTextBlockHeight(field) {
  if (Number.isFinite(field.maxHeight) && field.maxHeight > 0) {
    return field.maxHeight;
  }

  if (Number.isFinite(field.maxLines) && field.maxLines > 0) {
    return estimateTextBlockHeight(
      field.maxLines,
      field.size,
      field.lineGap ?? 1.75
    );
  }

  return Number.POSITIVE_INFINITY;
}

function fitMultilineTextToBlock(text, font, field) {
  const maxWidth = field.maxWidth ?? 70;
  const maxHeight = resolveTextBlockHeight(field);
  const baseSize = field.size;
  const baseLineGap = field.lineGap ?? 1.75;
  const minSize = field.minSize ?? Math.max(roundToTenths(baseSize * 0.45), 1.4);

  let bestLayout = {
    lines: wrapTextGreedy(text, font, baseSize, maxWidth),
    size: baseSize,
    lineGap: baseLineGap
  };

  if (
    estimateTextBlockHeight(
      bestLayout.lines.length,
      bestLayout.size,
      bestLayout.lineGap
    ) <= maxHeight
  ) {
    return bestLayout;
  }

  for (
    let size = roundToTenths(baseSize - 0.1);
    size >= minSize;
    size = roundToTenths(size - 0.1)
  ) {
    const scale = size / baseSize;
    const lineGap = roundToTenths(baseLineGap * scale);
    const lines = wrapTextGreedy(text, font, size, maxWidth);

    bestLayout = {
      lines,
      size,
      lineGap
    };

    if (estimateTextBlockHeight(lines.length, size, lineGap) <= maxHeight) {
      return bestLayout;
    }
  }

  return bestLayout;
}

function resolveMultilineTextLayout(text, font, field) {
  if (Number.isFinite(field.fixedSize) && field.fixedSize > 0) {
    return {
      lines: wrapTextGreedy(text, font, field.fixedSize, field.maxWidth ?? 70),
      size: field.fixedSize,
      lineGap: field.fixedLineGap ?? field.lineGap ?? 1.75
    };
  }

  return fitMultilineTextToBlock(text, font, field);
}

async function loadTemplateConfigs() {
  const entries = await fs.readdir(templatesDir, { withFileTypes: true });
  const configs = [];

  for (const entry of entries) {
    if (!entry.isFile() || !entry.name.endsWith(".json")) {
      continue;
    }

    const fullPath = path.join(templatesDir, entry.name);
    const content = await fs.readFile(fullPath, "utf8");
    configs.push(JSON.parse(content));
  }

  return configs.sort((a, b) => a.name.localeCompare(b.name, "zh-CN"));
}

async function getTemplateConfig(templateId) {
  const configs = await loadTemplateConfigs();
  return configs.find((config) => config.id === templateId) || null;
}

function rowsFromWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const rows = worksheetToRows(worksheet);

  const firstRowKeys = Object.keys(rows[0] || {});
  const hasLabelValueShape =
    rows.length > 1 &&
    firstRowKeys.length >= 2 &&
    firstRowKeys.slice(1).every((key) => key.startsWith("__EMPTY"));

  if (hasLabelValueShape) {
    const entries = rows
      .map((row) => {
        const keys = Object.keys(row);
        const labelKey = keys[0];
        const valueKeys = keys.filter((key) => key !== labelKey);
        const valueKey =
          valueKeys.find((key) => normalizeCellValue(row[key])) ?? valueKeys[0];
        return [normalizeCellValue(row[labelKey]), row[valueKey]];
      })
      .filter(([label]) => label);

    return [Object.fromEntries(entries)];
  }

  return rows;
}

function filterEmployeeRows(rows, templateConfig) {
  const requiredColumns = [
    ...new Set([
      templateConfig.fileNameField,
      ...templateConfig.fields.flatMap((field) => getFieldExcelColumns(field))
    ])
  ].filter(Boolean);

  return rows.filter((row) =>
    requiredColumns.some((column) => normalizeCellValue(row[column]))
  );
}

function findMissingColumns(rows, templateConfig) {
  const firstRow = rows[0] || {};
  const present = new Set(Object.keys(firstRow));
  const requiredColumns = getRequiredColumns(templateConfig);

  return requiredColumns.filter((column) => !present.has(column));
}

function formatDynamicFieldValue(excelColumn, value, field = null) {
  const sanitizedValue = normalizeMultilineValue(value);
  if (!sanitizedValue) {
    return sanitizedValue;
  }

  if (field?.disableAutoFormat) {
    return sanitizedValue;
  }

  if (excelColumn === "English Name" || excelColumn === "Job Title in English") {
    return sanitizedValue.toUpperCase();
  }

  if (
    excelColumn === "E-Mail Address" ||
    excelColumn === "邮箱" ||
    excelColumn === "E-mail"
  ) {
    return sanitizedValue.toLowerCase();
  }

  if (
    excelColumn === "Mobile Number" ||
    excelColumn === "WhatsApp Number" ||
    excelColumn === "手机" ||
    excelColumn === "Mobile"
  ) {
    return formatMobileNumberValue(sanitizedValue);
  }

  return sanitizedValue;
}

function formatSingleMobileNumber(value) {
  const rawValue = normalizeCellValue(value);
  const mainlandPattern = /^(?:(?:\+86|0086)[\s-]*)?1\d{10}$|^1\d{10}$/;
  if (!mainlandPattern.test(rawValue.replace(/[^\d+]/g, ""))) {
    return rawValue;
  }

  const digitsOnly = rawValue.replace(/\D/g, "");
  let localNumber = "";

  if (digitsOnly.length === 11 && digitsOnly.startsWith("1")) {
    localNumber = digitsOnly;
  } else if (digitsOnly.length === 13 && digitsOnly.startsWith("861")) {
    localNumber = digitsOnly.slice(2);
  } else if (digitsOnly.length === 15 && digitsOnly.startsWith("00861")) {
    localNumber = digitsOnly.slice(4);
  }

  if (localNumber.length === 11) {
    return `+86 ${localNumber.slice(0, 3)} ${localNumber.slice(3, 7)} ${localNumber.slice(7)}`;
  }

  return rawValue;
}

function containsNonMainlandPhoneFormat(value) {
  const text = normalizeCellValue(value);
  if (!text) {
    return false;
  }

  return /[()]/.test(text) || /(?:\+(?!\s*86\b)|00(?!86\b))/.test(text);
}

function isStrictMainlandMobileText(value) {
  const text = normalizeCellValue(value);
  if (!text) {
    return false;
  }

  const compact = text.replace(/[\s-]/g, "");
  return /^(?:\+86|0086|86)?1\d{10}$/.test(compact);
}

function extractMobileCandidates(value) {
  return Array.from(
    normalizeCellValue(value).matchAll(/(?:\+?86|0086)?[\s-]*1\d{10}|1\d{10}/g)
  )
    .map((match) => match[0])
    .filter(Boolean);
}

function formatMobileLine(line) {
  if (containsNonMainlandPhoneFormat(line)) {
    return normalizeCellValue(line);
  }

  if (isStrictMainlandMobileText(line)) {
    return formatSingleMobileNumber(line);
  }

  const matches = extractMobileCandidates(line);
  if (matches.length > 1) {
    return matches.map((match) => formatSingleMobileNumber(match)).join(" ");
  }

  return normalizeCellValue(line);
}

function formatMobileNumberValue(value) {
  const normalizedValue = normalizeMultilineValue(value);
  if (!normalizedValue) {
    return normalizedValue;
  }

  if (containsNonMainlandPhoneFormat(normalizedValue)) {
    return normalizedValue;
  }

  const explicitLines = normalizedValue
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  if (explicitLines.length > 1) {
    return explicitLines.map((line) => formatMobileLine(line)).join("\n");
  }

  if (isStrictMainlandMobileText(normalizedValue)) {
    return formatSingleMobileNumber(normalizedValue);
  }

  const matches = extractMobileCandidates(normalizedValue);
  const nonSeparatorRemainder = normalizedValue.replace(
    /(?:\+?86|0086)?[\s-]*1\d{10}|[\s,，、;/+()-]/g,
    ""
  );

  if (matches.length > 1 && !nonSeparatorRemainder) {
    return matches.map((match) => formatSingleMobileNumber(match)).join("\n");
  }

  return normalizedValue;
}

function getFieldText(field, row) {
  if (field.kind === "static-text") {
    return sanitizeTextForPdf(field.text);
  }

  const formattedValue = formatDynamicFieldValue(field.excelColumn, row[field.excelColumn], field);

  if (field.textTransform === "svea-cn-name-spacing") {
    const compact = formattedValue.replace(/\s+/g, "");
    if (compact.length === 2) {
      return `${compact[0]} ${compact[1]}`;
    }
    if (compact.length === 3) {
      return `${compact[0]} ${compact.slice(1)}`;
    }
    return compact || formattedValue;
  }

  return formattedValue;
}

function pickFont(fonts, field) {
  if (field?.fontKey && fonts[field.fontKey]) {
    return fonts[field.fontKey];
  }

  if (field?.fontWeight === "bold" && fonts.bold) {
    return fonts.bold;
  }

  return fonts.regular || fonts.bold;
}

function resolveTextY(field, font, size) {
  if (
    Number.isFinite(field.topLineY) &&
    Number.isFinite(field.gapToTopLine)
  ) {
    const textHeight = font.heightAtSize(size, { descender: false });
    return field.topLineY - field.gapToTopLine - textHeight;
  }

  return field.y;
}

function resolveFieldPosition(field, positionOverride = null, font = null, size = null) {
  return {
    x: Number.isFinite(positionOverride?.x) ? positionOverride.x : field.x,
    y: Number.isFinite(positionOverride?.y)
      ? positionOverride.y
      : resolveTextY(field, font, size ?? field.size)
  };
}

function splitRenderedTextLines(value) {
  return normalizeMultilineValue(value)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
}

function normalizePhoneComparisonLines(value) {
  return splitRenderedTextLines(formatMobileNumberValue(value)).map((line) =>
    line.replace(/\D/g, "")
  );
}

function haveSamePhoneValue(leftValue, rightValue) {
  const leftLines = normalizePhoneComparisonLines(leftValue);
  const rightLines = normalizePhoneComparisonLines(rightValue);

  if (!leftLines.length || leftLines.length !== rightLines.length) {
    return false;
  }

  return leftLines.every((line, index) => line === rightLines[index]);
}

function buildContactEntries(field, row) {
  const mobileValue = formatDynamicFieldValue("Mobile Number", row["Mobile Number"]);
  const whatsappValue = formatDynamicFieldValue("WhatsApp Number", row["WhatsApp Number"]);

  if (!mobileValue && !whatsappValue) {
    return [];
  }

  if (!mobileValue && whatsappValue) {
    return [
      {
        fieldKey: "WhatsApp Number",
        label: "W :",
        rawValue: whatsappValue
      }
    ];
  }

  if (mobileValue && !whatsappValue) {
    return [
      {
        fieldKey: "Mobile Number",
        label: "M :",
        rawValue: mobileValue
      }
    ];
  }

  if (haveSamePhoneValue(mobileValue, whatsappValue)) {
    return [
      {
        fieldKey: "Mobile Number",
        label: "M/W :",
        rawValue: mobileValue
      }
    ];
  }

  return [
    {
      fieldKey: "Mobile Number",
      label: "M :",
      rawValue: mobileValue
    },
    {
      fieldKey: "WhatsApp Number",
      label: "W :",
      rawValue: whatsappValue
    }
  ].filter((entry) => entry.rawValue);
}

function drawJustifiedText(page, text, options) {
  const { x, y, size, font, color, maxWidth, justify } = options;
  const value = sanitizeTextForPdf(text);

  if (!justify) {
    page.drawText(value, { x, y, size, font, color });
    return;
  }

  const words = value.split(/\s+/).filter(Boolean);
  if (words.length <= 1) {
    page.drawText(value, { x, y, size, font, color });
    return;
  }

  const totalWordsWidth = words.reduce(
    (sum, word) => sum + font.widthOfTextAtSize(word, size),
    0
  );
  const gapCount = words.length - 1;
  const gapWidth = Math.max((maxWidth - totalWordsWidth) / gapCount, 0);

  let cursorX = x;
  for (const [index, word] of words.entries()) {
    page.drawText(word, {
      x: cursorX,
      y,
      size,
      font,
      color
    });

    cursorX += font.widthOfTextAtSize(word, size);
    if (index < gapCount) {
      cursorX += gapWidth;
    }
  }
}

async function loadTemplateFonts(pdfDoc, templateConfig = {}) {
  pdfDoc.registerFontkit(fontkit);
  const templateFonts = templateConfig.fonts && typeof templateConfig.fonts === "object"
    ? templateConfig.fonts
    : {};
  const fontPaths = {
    regular: templateFonts.regular || defaultRegularFontPath,
    bold: templateFonts.bold || defaultBoldFontPath,
    ...templateFonts
  };
  const resolvedEntries = Object.entries(fontPaths).filter(([, value]) => normalizeCellValue(value));
  const cache = new Map();
  const fonts = {};

  for (const [key, fontPath] of resolvedEntries) {
    const resolvedPath = path.isAbsolute(fontPath) ? fontPath : path.join(__dirname, fontPath);

    if (!cache.has(resolvedPath)) {
      const fontBytes = await fs.readFile(resolvedPath);
      cache.set(resolvedPath, await pdfDoc.embedFont(fontBytes));
    }

    fonts[key] = cache.get(resolvedPath);
  }

  return fonts;
}

async function buildBusinessCardPdf(templateConfig, employee) {
  const pdfPath = path.join(templatesDir, templateConfig.pdfFile);
  const templateBytes = await fs.readFile(pdfPath);
  const pdfDoc = await PDFDocument.load(templateBytes);
  const page = pdfDoc.getPage(templateConfig.pageIndex);
  const fonts = await loadTemplateFonts(pdfDoc, templateConfig);
  const row = employee.row;
  const fieldPositions = employee.fieldPositions ?? {};

  for (const mask of templateConfig.masks || []) {
    page.drawRectangle({
      x: mask.x,
      y: mask.y,
      width: mask.width,
      height: mask.height,
      color: cmyk(...mask.colorCmyk)
    });
  }

  for (const [index, field] of (templateConfig.fields || []).entries()) {
    if (field.kind === "contact-block") {
      const labelFont = field.labelFontWeight === "bold" ? fonts.bold : fonts.regular;
      const valueFont = pickFont(fonts, field);
      let currentY = field.y;
      const blockPositionKey = getFieldPositionKey(field, index);

      for (const entry of buildContactEntries(field, row)) {
        const positionOverride = fieldPositions[entry.fieldKey];
        const labelPositionOverride = fieldPositions[`${blockPositionKey}:${entry.fieldKey}:label`];
        const entryX = Number.isFinite(positionOverride?.x) ? positionOverride.x : field.x;
        const entryY = Number.isFinite(positionOverride?.y) ? positionOverride.y : currentY;
        const labelX = Number.isFinite(labelPositionOverride?.x)
          ? labelPositionOverride.x
          : field.labelX ?? field.x;
        const labelY = Number.isFinite(labelPositionOverride?.y)
          ? labelPositionOverride.y
          : entryY;
        const textLines = splitRenderedTextLines(entry.rawValue);

        page.drawText(entry.label, {
          x: labelX,
          y: labelY,
          size: field.labelSize ?? field.size,
          font: labelFont,
          color: cmyk(...field.colorCmyk)
        });

        textLines.forEach((line, index) => {
          page.drawText(line, {
            x: entryX,
            y: entryY - index * (field.lineGap ?? field.size * 1.05),
            size: field.size,
            font: valueFont,
            color: cmyk(...field.colorCmyk)
          });
        });

        currentY = entryY - textLines.length * (field.lineGap ?? field.size * 1.05);
      }

      continue;
    }

    const rawValue = getFieldText(field, row);
    if (!rawValue) {
      continue;
    }

    if (field.kind === "multiline-address") {
      const font = pickFont(fonts, field);
      const layout = resolveMultilineTextLayout(rawValue, font, field);
      const position = resolveFieldPosition(
        field,
        fieldPositions[getFieldPositionKey(field, index)],
        font,
        layout.size
      );

      layout.lines.forEach((line, index) => {
        drawJustifiedText(page, line, {
          x: position.x,
          y: position.y - index * layout.lineGap,
          size: layout.size,
          font,
          color: cmyk(...field.colorCmyk),
          maxWidth: field.maxWidth ?? 70,
          justify: field.justify === true && index < layout.lines.length - 1
        });
      });

      continue;
    }

    const font = pickFont(fonts, field);
    const position = resolveFieldPosition(
      field,
      fieldPositions[getFieldPositionKey(field, index)],
      font,
      field.size
    );
    const textLines = splitRenderedTextLines(rawValue);

    if (textLines.length > 1) {
      textLines.forEach((line, index) => {
        page.drawText(line, {
          x: position.x,
          y: position.y - index * (field.lineGap ?? field.size * 1.05),
          size: field.size,
          font,
          color: cmyk(...field.colorCmyk)
        });
      });
      continue;
    }

    page.drawText(rawValue, {
      x: position.x,
      y: position.y,
      size: field.size,
      font,
      color: cmyk(...field.colorCmyk)
    });
  }

  return pdfDoc.save();
}

async function buildPreviewPdf(templateConfig, employee) {
  const fullPdfBytes = await buildBusinessCardPdf(templateConfig, employee);
  const sourcePdf = await PDFDocument.load(fullPdfBytes);
  const previewPdf = await PDFDocument.create();
  const [previewPage] = await previewPdf.copyPages(sourcePdf, [templateConfig.pageIndex]);
  previewPdf.addPage(previewPage);
  return previewPdf.save();
}

function ensureFeishuConfigured() {
  if (!FEISHU_APP_ID || !FEISHU_APP_SECRET) {
    throw new Error("请先配置 FEISHU_APP_ID 和 FEISHU_APP_SECRET。");
  }
}

function getEmployeeContactValue(row, columns) {
  for (const column of columns) {
    const value = normalizeCellValue(row[column]);
    if (value) {
      return value;
    }
  }

  return "";
}

function normalizeFeishuMobile(value) {
  const rawValue = normalizeCellValue(value);
  if (!rawValue) {
    return "";
  }

  if (rawValue.startsWith("+")) {
    return rawValue.replace(/[^\d+]/g, "");
  }

  const digits = rawValue.replace(/\D/g, "");
  if (!digits) {
    return "";
  }

  if (digits.startsWith("00")) {
    return `+${digits.slice(2)}`;
  }

  if (digits.length === 11 && digits.startsWith("1")) {
    return `+86${digits}`;
  }

  if (digits.startsWith("86") && digits.length >= 13) {
    return `+${digits}`;
  }

  if (digits.startsWith("852") || digits.startsWith("81")) {
    return `+${digits}`;
  }

  return digits;
}

async function callFeishuApi(endpoint, options = {}) {
  const response = await fetch(`${FEISHU_BASE_URL}${endpoint}`, options);
  const data = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(data.msg || data.message || `飞书接口请求失败：${response.status}`);
  }

  if (Number(data.code) !== 0) {
    throw new Error(data.msg || data.message || "飞书接口返回失败。");
  }

  return data;
}

async function getFeishuTenantAccessToken() {
  const now = Date.now();
  if (feishuTokenCache.token && feishuTokenCache.expiresAt > now + 60 * 1000) {
    return feishuTokenCache.token;
  }

  ensureFeishuConfigured();
  const data = await callFeishuApi("/open-apis/auth/v3/tenant_access_token/internal", {
    method: "POST",
    headers: {
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify({
      app_id: FEISHU_APP_ID,
      app_secret: FEISHU_APP_SECRET
    })
  });

  feishuTokenCache = {
    token: normalizeCellValue(data.tenant_access_token),
    expiresAt: now + (Number(data.expire) || 7200) * 1000
  };

  return feishuTokenCache.token;
}

async function resolveFeishuUser(token, employee) {
  const email = getEmployeeContactValue(employee.row, FEISHU_EMAIL_COLUMNS);
  const mobile = normalizeFeishuMobile(getEmployeeContactValue(employee.row, FEISHU_MOBILE_COLUMNS));

  if (!email && !mobile) {
    throw new Error("当前员工缺少邮箱或手机号，无法匹配飞书联系人。");
  }

  const payload = {};
  if (email) {
    payload.emails = [email.toLowerCase()];
  }
  if (mobile) {
    payload.mobiles = [mobile];
  }

  const data = await callFeishuApi("/open-apis/contact/v3/users/batch_get_id?user_id_type=user_id", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify(payload)
  });

  const userList = Array.isArray(data.data?.user_list)
    ? data.data.user_list
    : Array.isArray(data.user_list)
      ? data.user_list
      : [];
  const matchedUser = userList[0];

  if (!matchedUser?.user_id) {
    throw new Error("未在飞书中匹配到当前员工，请确认邮箱或手机号与飞书通讯录一致。");
  }

  return {
    userId: matchedUser.user_id,
    email,
    mobile
  };
}

async function uploadFeishuFile(token, fileName, pdfBytes) {
  const formData = new FormData();
  formData.append("file_type", "stream");
  formData.append("file_name", fileName);
  formData.append(
    "file",
    new Blob([Buffer.from(pdfBytes)], { type: "application/pdf" }),
    fileName
  );

  const data = await callFeishuApi("/open-apis/im/v1/files", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`
    },
    body: formData
  });

  const fileKey = normalizeCellValue(data.data?.file_key || data.file_key);
  if (!fileKey) {
    throw new Error("飞书文件上传成功，但未返回 file_key。");
  }

  return fileKey;
}

async function sendFeishuMessage(token, receiveId, msgType, content) {
  await callFeishuApi("/open-apis/im/v1/messages?receive_id_type=user_id", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify({
      receive_id: receiveId,
      msg_type: msgType,
      content: JSON.stringify(content)
    })
  });
}

function getEditableColumns(templateConfig) {
  return [
    ...new Set(templateConfig.fields.flatMap((field) => getFieldExcelColumns(field)))
  ].filter(Boolean);
}

function isEnglishNameColumn(columnName) {
  const normalized = normalizeCellValue(columnName);
  if (!normalized) {
    return false;
  }

  const lowerCased = normalized.toLowerCase();
  if (lowerCased.includes("english") && lowerCased.includes("name")) {
    return true;
  }

  return /[\u82f1][\u6587]/.test(normalized) && /[\u59d3\u540d]/.test(normalized);
}

function getFieldConfigByColumn(templateConfig, columnName) {
  return templateConfig.fields.find((field) => field.excelColumn === columnName) ?? null;
}

function formatRowValueForFileName(columnName, row, templateConfig) {
  const fieldConfig = getFieldConfigByColumn(templateConfig, columnName);
  return formatDynamicFieldValue(columnName, row[columnName], fieldConfig);
}

function looksLikeLatinNameValue(value) {
  const normalized = normalizeCellValue(value);
  if (!normalized || /[@\d]/.test(normalized)) {
    return false;
  }

  const compact = normalized.replace(/\s+/g, " ").trim();
  return /^[A-Z][A-Z .'-]{1,}$/i.test(compact) && /[A-Z]{2,}/i.test(compact);
}

function resolvePdfBaseFileName(templateConfig, row, displayName, fallbackIndex) {
  const preferredColumns = [];

  if (templateConfig.exportFileNameField) {
    preferredColumns.push(templateConfig.exportFileNameField);
  }

  preferredColumns.push(
    ...templateConfig.fields
      .flatMap((field) => getFieldExcelColumns(field))
      .filter((columnName) => isEnglishNameColumn(columnName))
  );

  preferredColumns.push("English Name");

  const seenColumns = new Set();
  for (const columnName of preferredColumns) {
    const normalizedColumnName = normalizeCellValue(columnName);
    if (!normalizedColumnName || seenColumns.has(normalizedColumnName)) {
      continue;
    }

    seenColumns.add(normalizedColumnName);
    const formattedValue = formatRowValueForFileName(normalizedColumnName, row, templateConfig);
    const baseFileName = sanitizeFileName(formattedValue);
    if (baseFileName) {
      return baseFileName;
    }
  }

  for (const field of templateConfig.fields) {
    if (field.kind !== "text" || !field.excelColumn || field.fontWeight !== "bold") {
      continue;
    }

    const formattedValue = formatRowValueForFileName(field.excelColumn, row, templateConfig);
    if (!looksLikeLatinNameValue(formattedValue)) {
      continue;
    }

    const baseFileName = sanitizeFileName(formattedValue);
    if (baseFileName) {
      return baseFileName;
    }
  }

  return sanitizeFileName(displayName) || `employee-${fallbackIndex}`;
}

function sanitizeFileName(value) {
  const normalized = normalizeCellValue(value)
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  return normalized;
}

function withPdfExtension(fileName) {
  return fileName.toLowerCase().endsWith(".pdf") ? fileName : `${fileName}.pdf`;
}

function toAsciiFileName(fileName, fallbackBaseName = "file") {
  const parsed = path.parse(normalizeCellValue(fileName));
  const asciiName = sanitizeFileName(parsed.name).replace(/[^\x20-\x7E]/g, "").trim();
  const asciiExt = (parsed.ext || "").replace(/[^\x20-\x7E]/g, "");
  const safeName = asciiName || fallbackBaseName;
  const safeExt = asciiExt || path.extname(fileName) || "";
  return `${safeName}${safeExt}`;
}

function encodeContentDispositionFileName(fileName) {
  return encodeURIComponent(normalizeCellValue(fileName))
    .replace(/['()]/g, (char) => `%${char.charCodeAt(0).toString(16).toUpperCase()}`)
    .replace(/\*/g, "%2A");
}

function buildContentDisposition(dispositionType, fileName, fallbackBaseName = "file") {
  const fallbackName = toAsciiFileName(fileName, fallbackBaseName);
  const encodedName = encodeContentDispositionFileName(fileName);
  return `${dispositionType}; filename="${fallbackName}"; filename*=UTF-8''${encodedName}`;
}

function normalizeFieldPositions(rawPositions, templateConfig) {
  if (!rawPositions || typeof rawPositions !== "object") {
    return {};
  }

  const validKeys = getTemplatePositionKeys(templateConfig);
  const normalizedEntries = Object.entries(rawPositions)
    .filter(([fieldKey, value]) => validKeys.has(fieldKey) && value && typeof value === "object")
    .map(([fieldKey, value]) => [
      fieldKey,
      {
        ...(Number.isFinite(value.x) ? { x: Number(value.x) } : {}),
        ...(Number.isFinite(value.y) ? { y: Number(value.y) } : {})
      }
    ])
    .filter(([, value]) => Object.keys(value).length > 0);

  return Object.fromEntries(normalizedEntries);
}

function buildEmployeePayload(employeeInput, templateConfig, fallbackIndex) {
  const editableColumns = getEditableColumns(templateConfig);
  const inputRow = employeeInput?.row && typeof employeeInput.row === "object"
    ? employeeInput.row
    : employeeInput ?? {};
  const row = Object.fromEntries(
    editableColumns.map((column) => [column, normalizeCellValue(inputRow[column])])
  );
  const fileNameFieldConfig =
    templateConfig.fields.find((field) => field.excelColumn === templateConfig.fileNameField) ??
    null;
  const displayName =
    formatDynamicFieldValue(
      templateConfig.fileNameField,
      row[templateConfig.fileNameField],
      fileNameFieldConfig
    ) ||
    `EMPLOYEE ${fallbackIndex}`;
  const baseFileName = resolvePdfBaseFileName(templateConfig, row, displayName, fallbackIndex);

  return {
    id: normalizeCellValue(employeeInput?.id) || `employee-${fallbackIndex}`,
    sourceFileName: normalizeCellValue(employeeInput?.sourceFileName),
    displayName,
    pdfFileName: withPdfExtension(baseFileName),
    row,
    fieldPositions: normalizeFieldPositions(employeeInput?.fieldPositions, templateConfig)
  };
}

function buildEmployeePayloadsFromFiles(files, templateConfig) {
  const employees = [];
  let employeeIndex = 1;

  for (const file of files) {
    const rawRows = rowsFromWorkbook(file.buffer);
    const missingColumns = findMissingColumns(rawRows, templateConfig);
    if (missingColumns.length) {
      throw new Error(
        `${file.originalname} 缺少这些列：${missingColumns.join("、")}`
      );
    }

    const rows = filterEmployeeRows(rawRows, templateConfig);
    if (!rows.length) {
      continue;
    }

    for (const row of rows) {
      employees.push(
        buildEmployeePayload(
          {
            id: `${sanitizeFileName(file.originalname) || "file"}-${employeeIndex}`,
            sourceFileName: file.originalname,
            row
          },
          templateConfig,
          employeeIndex
        )
      );
      employeeIndex += 1;
    }
  }

  return employees;
}

function ensureUniqueFileName(fileName, seenNames) {
  const parsed = path.parse(fileName);
  const key = fileName.toLowerCase();
  const count = seenNames.get(key) ?? 0;
  seenNames.set(key, count + 1);

  if (count === 0) {
    return fileName;
  }

  return `${parsed.name} (${count + 1})${parsed.ext || ".pdf"}`;
}

function cmykToCss(colorCmyk = [0, 0, 0, 1]) {
  const [c, m, y, k] = colorCmyk.map((value) => Math.min(Math.max(Number(value) || 0, 0), 1));
  const r = Math.round(255 * (1 - c) * (1 - k));
  const g = Math.round(255 * (1 - m) * (1 - k));
  const b = Math.round(255 * (1 - y) * (1 - k));
  return `rgb(${r}, ${g}, ${b})`;
}

async function getTemplatePageSize(templateConfig) {
  if (Number.isFinite(templateConfig.editor?.pageWidth) && Number.isFinite(templateConfig.editor?.pageHeight)) {
    return {
      width: templateConfig.editor.pageWidth,
      height: templateConfig.editor.pageHeight
    };
  }

  const pdfPath = path.join(templatesDir, templateConfig.pdfFile);
  const templateBytes = await fs.readFile(pdfPath);
  const pdfDoc = await PDFDocument.load(templateBytes);
  const page = pdfDoc.getPage(templateConfig.pageIndex);
  const { width, height } = page.getSize();
  return { width, height };
}

async function buildBusinessCardRenderModel(templateConfig, employee) {
  const page = await getTemplatePageSize(templateConfig);
  const tempPdf = await PDFDocument.create();
  const fonts = await loadTemplateFonts(tempPdf, templateConfig);
  const elements = [];
  const row = employee.row;
  const fieldPositions = employee.fieldPositions ?? {};

  templateConfig.fields.forEach((field, index) => {
    if (field.kind === "contact-block") {
      let currentY = field.y;
      const blockPositionKey = getFieldPositionKey(field, index);

      buildContactEntries(field, row).forEach((entry, entryIndex) => {
        const positionOverride = fieldPositions[entry.fieldKey];
        const labelPositionOverride = fieldPositions[`${blockPositionKey}:${entry.fieldKey}:label`];
        const entryX = Number.isFinite(positionOverride?.x) ? positionOverride.x : field.x;
        const entryY = Number.isFinite(positionOverride?.y) ? positionOverride.y : currentY;
        const labelX = Number.isFinite(labelPositionOverride?.x)
          ? labelPositionOverride.x
          : field.labelX ?? field.x;
        const labelY = Number.isFinite(labelPositionOverride?.y)
          ? labelPositionOverride.y
          : entryY;
        const textLines = splitRenderedTextLines(entry.rawValue);
        const lineGap = field.lineGap ?? field.size * 1.05;

        elements.push({
          id: `${entry.fieldKey}-label-${entryIndex}`,
          editable: true,
          textEditable: false,
          positionKey: `${blockPositionKey}:${entry.fieldKey}:label`,
          fieldKey: "",
          label: entry.label,
          fontFamily: field.labelFontCssFamily || field.fontCssFamily || "",
          fontWeight: field.labelFontWeight ?? "normal",
          color: cmykToCss(field.colorCmyk),
          kind: "static-text",
          type: "text",
          rawValue: entry.label,
          size: field.labelSize ?? field.size,
          x: labelX,
          y: labelY
        });

        if (textLines.length > 1) {
          elements.push({
            id: entry.fieldKey,
            editable: true,
            textEditable: true,
            positionKey: entry.fieldKey,
            fieldKey: entry.fieldKey,
            label: entry.fieldKey,
            fontFamily: field.fontCssFamily || "",
            fontWeight: field.fontWeight ?? "normal",
            color: cmykToCss(field.colorCmyk),
            kind: "text",
            type: "multiline-text",
            rawValue: entry.rawValue,
            size: field.size,
            lineGap,
            x: entryX,
            y: entryY,
            lines: textLines.map((line, lineIndex) => ({
              text: line,
              x: entryX,
              y: entryY - lineIndex * lineGap,
              justify: false
            }))
          });
        } else {
          elements.push({
            id: entry.fieldKey,
            editable: true,
            textEditable: true,
            positionKey: entry.fieldKey,
            fieldKey: entry.fieldKey,
            label: entry.fieldKey,
            fontFamily: field.fontCssFamily || "",
            fontWeight: field.fontWeight ?? "normal",
            color: cmykToCss(field.colorCmyk),
            kind: "text",
            type: "text",
            rawValue: entry.rawValue,
            size: field.size,
            x: entryX,
            y: entryY
          });
        }

        currentY = entryY - textLines.length * lineGap;
      });

      return;
    }

    const rawValue = getFieldText(field, row);
    if (!rawValue) {
      return;
    }

    const font = pickFont(fonts, field);
    const positionKey = getFieldPositionKey(field, index);
    const baseElement = {
      id: positionKey,
      editable: true,
      textEditable: Boolean(field.excelColumn),
      positionKey,
      fieldKey: field.excelColumn || "",
      label: field.excelColumn || field.text || `field-${index}`,
      fontFamily: field.fontCssFamily || "",
      fontWeight: field.fontWeight ?? "normal",
      color: cmykToCss(field.colorCmyk),
      kind: field.kind
    };

    if (field.kind === "multiline-address") {
      const layout = resolveMultilineTextLayout(rawValue, font, field);
      const position = resolveFieldPosition(field, fieldPositions[positionKey], font, layout.size);
      const lines = layout.lines.map((line, lineIndex) => ({
        text: line,
        x: position.x,
        y: position.y - lineIndex * layout.lineGap,
        justify: field.justify === true && lineIndex < layout.lines.length - 1
      }));

      elements.push({
        ...baseElement,
        type: "multiline-text",
        rawValue,
        size: layout.size,
        lineGap: layout.lineGap,
        x: position.x,
        y: position.y,
        maxWidth: field.maxWidth ?? 70,
        maxHeight: field.maxHeight ?? null,
        lines
      });
      return;
    }

    const position = resolveFieldPosition(field, fieldPositions[positionKey], font, field.size);
    const textLines = splitRenderedTextLines(rawValue);

    if (textLines.length > 1) {
      elements.push({
        ...baseElement,
        type: "multiline-text",
        rawValue,
        size: field.size,
        lineGap: field.lineGap ?? field.size * 1.05,
        x: position.x,
        y: position.y,
        lines: textLines.map((line, lineIndex) => ({
          text: line,
          x: position.x,
          y: position.y - lineIndex * (field.lineGap ?? field.size * 1.05),
          justify: false
        }))
      });
      return;
    }

    elements.push({
      ...baseElement,
      type: "text",
      rawValue,
      size: field.size,
      x: position.x,
      y: position.y
    });
  });

  return {
    page,
    decorations: templateConfig.editor?.decorations ?? [],
    elements
  };
}

app.get("/api/templates", async (_req, res) => {
  try {
    const configs = await loadTemplateConfigs();
    res.json({
      templates: configs.map((config) => ({
        id: config.id,
        name: config.name,
        columns: getEditableColumns(config),
        fileNameField: config.fileNameField
      }))
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/parse", upload.array("excel"), async (req, res) => {
  try {
    if (!req.files?.length) {
      return res.status(400).json({ error: "请先上传至少一个 Excel 文件。" });
    }

    const templateId = req.body.templateId;
    if (!templateId) {
      return res.status(400).json({ error: "请选择名片模板。" });
    }

    const templateConfig = await getTemplateConfig(templateId);
    if (!templateConfig) {
      return res.status(404).json({ error: "未找到所选模板。" });
    }

    const employees = buildEmployeePayloadsFromFiles(req.files, templateConfig);
    if (!employees.length) {
      return res.status(400).json({ error: "上传的 Excel 中没有可生成的员工数据。" });
    }

    res.json({
      employees,
      columns: getEditableColumns(templateConfig),
      fileNameField: templateConfig.fileNameField
    });
  } catch (error) {
    res.status(500).json({ error: error.message || "解析 Excel 失败。" });
  }
});

app.post("/api/preview", async (req, res) => {
  try {
    const { templateId, employee } = req.body ?? {};
    if (!templateId) {
      return res.status(400).json({ error: "请选择名片模板。" });
    }

    const templateConfig = await getTemplateConfig(templateId);
    if (!templateConfig) {
      return res.status(404).json({ error: "未找到所选模板。" });
    }

    const resolvedEmployee = buildEmployeePayload(employee, templateConfig, 1);
    const pdfBytes = await buildPreviewPdf(templateConfig, resolvedEmployee);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      buildContentDisposition("inline", resolvedEmployee.pdfFileName, "business-card")
    );
    res.send(Buffer.from(pdfBytes));
  } catch (error) {
    res.status(500).json({ error: error.message || "生成预览失败。" });
  }
});

app.post("/api/feishu/send-current", async (req, res) => {
  try {
    const { templateId, employee } = req.body ?? {};
    if (!templateId) {
      return res.status(400).json({ error: "请选择名片模板。" });
    }

    if (!employee || typeof employee !== "object") {
      return res.status(400).json({ error: "请先选择要发送的员工。" });
    }

    const templateConfig = await getTemplateConfig(templateId);
    if (!templateConfig) {
      return res.status(404).json({ error: "未找到所选模板。" });
    }

    const resolvedEmployee = buildEmployeePayload(employee, templateConfig, 1);
    const token = await getFeishuTenantAccessToken();
    const matchedUser = await resolveFeishuUser(token, resolvedEmployee);
    const pdfBytes = await buildBusinessCardPdf(templateConfig, resolvedEmployee);
    const fileKey = await uploadFeishuFile(token, resolvedEmployee.pdfFileName, pdfBytes);

    await sendFeishuMessage(token, matchedUser.userId, "file", { file_key: fileKey });
    await sendFeishuMessage(token, matchedUser.userId, "text", { text: FEISHU_CONFIRMATION_TEXT });

    res.json({
      ok: true,
      employeeName: resolvedEmployee.displayName,
      receiveId: matchedUser.userId,
      matchedBy: matchedUser.email ? "email" : "mobile"
    });
  } catch (error) {
    res.status(500).json({ error: error.message || "发送飞书确认消息失败。" });
  }
});

app.post("/api/render-model", async (req, res) => {
  try {
    const { templateId, employee } = req.body ?? {};
    if (!templateId) {
      return res.status(400).json({ error: "请选择名片模板。" });
    }

    const templateConfig = await getTemplateConfig(templateId);
    if (!templateConfig) {
      return res.status(404).json({ error: "未找到所选模板。" });
    }

    const resolvedEmployee = buildEmployeePayload(employee, templateConfig, 1);
    const model = await buildBusinessCardRenderModel(templateConfig, resolvedEmployee);

    res.json(model);
  } catch (error) {
    res.status(500).json({ error: error.message || "生成画布模型失败。" });
  }
});

app.post("/api/download-merged", async (req, res) => {
  try {
    const { templateId, employees } = req.body ?? {};
    if (!templateId) {
      return res.status(400).json({ error: "请选择名片模板。" });
    }

    if (!Array.isArray(employees) || !employees.length) {
      return res.status(400).json({ error: "请先加载员工数据。" });
    }

    const templateConfig = await getTemplateConfig(templateId);
    if (!templateConfig) {
      return res.status(404).json({ error: "未找到所选模板。" });
    }

    const mergedPdf = await PDFDocument.create();
    for (const [index, employeeInput] of employees.entries()) {
      const employee = buildEmployeePayload(employeeInput, templateConfig, index + 1);
      const pdfBytes = await buildBusinessCardPdf(templateConfig, employee);
      const sourcePdf = await PDFDocument.load(pdfBytes);
      const copiedPages = await mergedPdf.copyPages(sourcePdf, sourcePdf.getPageIndices());
      copiedPages.forEach((page) => mergedPdf.addPage(page));
    }

    const mergedBytes = await mergedPdf.save();
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      buildContentDisposition("attachment", `${templateConfig.id}-cards.pdf`, "business-cards")
    );
    res.send(Buffer.from(mergedBytes));
  } catch (error) {
    res.status(500).json({ error: error.message || "合并导出 PDF 失败。" });
  }
});

app.post("/api/download-all", async (req, res) => {
  try {
    const { templateId, employees } = req.body ?? {};
    if (!templateId) {
      return res.status(400).json({ error: "请选择名片模板。" });
    }

    if (!Array.isArray(employees) || !employees.length) {
      return res.status(400).json({ error: "请先加载员工数据。" });
    }

    const templateConfig = await getTemplateConfig(templateId);
    if (!templateConfig) {
      return res.status(404).json({ error: "未找到所选模板。" });
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      buildContentDisposition("attachment", `${templateConfig.id}-cards.zip`, "business-cards")
    );

    const archive = archiver("zip", { zlib: { level: 9 } });
    archive.on("error", (error) => {
      if (!res.headersSent) {
        res.status(500).json({ error: error.message || "打包 PDF 失败。" });
        return;
      }

      res.destroy(error);
    });

    archive.pipe(res);

    const usedNames = new Map();
    for (const [index, employeeInput] of employees.entries()) {
      const employee = buildEmployeePayload(employeeInput, templateConfig, index + 1);
      const pdfBytes = await buildBusinessCardPdf(templateConfig, employee);
      const entryName = ensureUniqueFileName(employee.pdfFileName, usedNames);
      archive.append(Buffer.from(pdfBytes), { name: entryName });
    }

    await archive.finalize();
  } catch (error) {
    if (!res.headersSent) {
      res.status(500).json({ error: error.message || "批量导出失败。" });
      return;
    }

    res.end();
  }
});

app.post("/api/generate", upload.single("excel"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "请先上传员工 Excel 表格。" });
    }

    const templateId = req.body.templateId;
    if (!templateId) {
      return res.status(400).json({ error: "请选择名片模板。" });
    }

    const templateConfig = await getTemplateConfig(templateId);
    if (!templateConfig) {
      return res.status(404).json({ error: "未找到所选模板。" });
    }

    const rawRows = rowsFromWorkbook(req.file.buffer);
    const missingColumns = findMissingColumns(rawRows, templateConfig);
    if (missingColumns.length) {
      return res.status(400).json({
        error: `Excel 缺少这些列：${missingColumns.join("、")}`
      });
    }

    const rows = filterEmployeeRows(rawRows, templateConfig);
    if (!rows.length) {
      return res.status(400).json({ error: "Excel 表格中没有可生成的员工数据。" });
    }

    const mergedPdf = await PDFDocument.create();

    for (const [index, row] of rows.entries()) {
      const employee = buildEmployeePayload({ row }, templateConfig, index + 1);
      const pdfBytes = await buildBusinessCardPdf(templateConfig, employee);
      const sourcePdf = await PDFDocument.load(pdfBytes);
      const copiedPages = await mergedPdf.copyPages(sourcePdf, sourcePdf.getPageIndices());
      copiedPages.forEach((page) => mergedPdf.addPage(page));
    }

    const mergedBytes = await mergedPdf.save();
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      buildContentDisposition("attachment", `${templateConfig.id}-cards.pdf`, "business-cards")
    );
    res.send(Buffer.from(mergedBytes));
  } catch (error) {
    if (!res.headersSent) {
      res.status(500).json({ error: error.message || "生成名片失败。" });
      return;
    }
    res.end();
  }
});

app.use((error, _req, res, next) => {
  if (error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE") {
    res.status(400).json({
      error: `上传文件不能超过 ${uploadFileSizeMb} MB。`
    });
    return;
  }

  next(error);
});

app.listen(port, host, () => {
  console.log(`Business card web app running at http://${host}:${port}`);
});
