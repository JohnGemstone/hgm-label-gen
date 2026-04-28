import fontkit from "@pdf-lib/fontkit";
import {
  PDFDocument,
  PDFPage,
  StandardFonts,
  type PDFFont,
  rgb,
} from "pdf-lib";

const LABEL_REQUIRED_CSV_HEADERS = [
  "firstname",
  "lastname",
  "address1",
  "address2",
  "address3",
  "town",
  "county",
  "postcode",
] as const;

const FORM_REQUIRED_CSV_HEADERS = [
  "enquiryid",
  "price",
  "9ct",
  "14ct",
  "18ct",
  "22ct",
  "24ct",
  "sovereign",
  "1ozbritiannia",
  "goldkrugerrand",
  "returnbydate",
] as const;

export const REQUIRED_CSV_HEADERS = [
  ...LABEL_REQUIRED_CSV_HEADERS,
  ...FORM_REQUIRED_CSV_HEADERS,
] as const;

export const REQUIRED_CSV_HEADER_DISPLAY_NAMES = [
  "EnquiryID",
  "FirstName",
  "LastName",
  "Address1",
  "Address2",
  "Address3",
  "Town",
  "County",
  "Postcode",
  "Price",
  "9ct",
  "14ct",
  "18ct",
  "22ct",
  "24ct",
  "Sovereign",
  "1oz Britiannia",
  "Gold Krugerrand",
  "Return By Date",
] as const;

export type RequiredCsvHeader = (typeof REQUIRED_CSV_HEADERS)[number];

export interface AddressLabel {
  name: string;
  lines: string[];
  postcode: string;
  displayLines: string[];
  warnings: string[];
}

export interface PreparedFormPrice {
  label: string;
  value: string;
}

export interface PreparedFormCustomerDetails {
  nameLine: string;
  addressLines: string[];
  countyLine: string;
  postcode: string;
}

export interface PreparedFormRow {
  customerDetails: PreparedFormCustomerDetails;
  enquiryId: string;
  returnByDate: string;
  price: PreparedFormPrice;
  nineCt: string;
  fourteenCt: string;
  eighteenCt: string;
  twentyTwoCt: string;
  twentyFourCt: string;
  sovereign: string;
  britiannia: string;
  krugerrand: string;
}

export interface SkippedCsvRow {
  rowNumber: number;
  reason: string;
}

export interface CsvPreparationInput {
  headers: string[];
  rows: Array<Record<string, string | undefined>>;
}

export interface PreparedLabelResult {
  labels: AddressLabel[];
  pages: AddressLabel[][];
  forms: PreparedFormRow[];
  formPageCount: number;
  missingHeaders: RequiredCsvHeader[];
  skippedRows: SkippedCsvRow[];
  emptyRowCount: number;
  totalRowCount: number;
  warningCount: number;
}

export interface LabelLayoutConfig {
  pageWidth: number;
  pageHeight: number;
  columns: number;
  rows: number;
  marginX: number;
  marginY: number;
  gutterX: number;
  gutterY: number;
  cellWidth: number;
  cellHeight: number;
  innerPadding: number;
  fontSize: number;
  lineHeight: number;
  maxLines: number;
  borderWidth: number;
}

type FormFontKey = "regular" | "medium" | "bold";

type FormFontData = Record<FormFontKey, Uint8Array>;

interface GenerateFormPdfOptions {
  fontData?: Partial<Record<FormFontKey, ArrayBuffer | Uint8Array>>;
  globalOffsetMm?: FormCoordinateOffset;
}

interface GenerateLabelPdfOptions {
  fontData?: Partial<Record<FormFontKey, ArrayBuffer | Uint8Array>>;
}

interface LabelFonts {
  name: PDFFont;
  address: PDFFont;
}

interface FormTextStyle {
  fontKey: FormFontKey;
  fontSize: number;
  lineHeight: number;
  uppercase?: boolean;
}

interface FormFieldLayout {
  left: number;
  top: number;
  width: number;
  height: number;
  style: FormTextStyle;
}

interface FormLayoutConfig {
  pageWidth: number;
  pageHeight: number;
  fields: {
    customerDetails: FormFieldLayout;
    enquiryId: FormFieldLayout;
    returnByDate: FormFieldLayout;
    price: FormFieldLayout;
    nineCt: FormFieldLayout;
    fourteenCt: FormFieldLayout;
    eighteenCt: FormFieldLayout;
    twentyTwoCt: FormFieldLayout;
    twentyFourCt: FormFieldLayout;
    sovereign: FormFieldLayout;
    britiannia: FormFieldLayout;
    krugerrand: FormFieldLayout;
  };
}

type FormFieldName = keyof FormLayoutConfig["fields"];

export interface FormCoordinateOffset {
  x: number;
  y: number;
}

const PAGE_WIDTH_PT = 595.28;
const PAGE_HEIGHT_PT = 841.89;
const GRID_COLUMNS = 3;
const GRID_ROWS = 7;
const CELL_WIDTH_PT = 180;
const LABEL_AREA_HEIGHT_PT = 761;
const TOP_OFFSET_PT = 43.5;
const POINTS_PER_MM = 72 / 25.4;
const COLUMN_GAPS_PT = [2.5 * POINTS_PER_MM, 8] as const;
const GUTTER_X = COLUMN_GAPS_PT[0];
const GUTTER_Y = 0;
const MARGIN_X =
  (PAGE_WIDTH_PT -
    (GRID_COLUMNS * CELL_WIDTH_PT +
      COLUMN_GAPS_PT.reduce((total, gap) => total + gap, 0))) /
  2;
const INNER_PADDING = 12;
const FONT_SIZE = 10;
const LINE_HEIGHT = 12;
const SHOW_LABEL_OUTLINES = false;
export const FORM_GLOBAL_OFFSET_MM: FormCoordinateOffset = {
  x: 4,
  y: -1.8,
};
const FORM_FIELD_OFFSETS_MM: Partial<
  Record<FormFieldName, FormCoordinateOffset>
> = {
  returnByDate: {
    x: -6,
    y: 0,
  },
};

const FORM_FONT_PATHS: Record<FormFontKey, string> = {
  regular: "/fonts/IBMPlexSans-Regular.ttf",
  medium: "/fonts/IBMPlexSans-Medium.ttf",
  bold: "/fonts/IBMPlexSans-Bold.ttf",
};

const GBP_FORMATTER = new Intl.NumberFormat("en-GB", {
  style: "currency",
  currency: "GBP",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

let cachedFormFontDataPromise: Promise<FormFontData> | null = null;

export const LABEL_LAYOUT_CONFIG: LabelLayoutConfig = {
  pageWidth: PAGE_WIDTH_PT,
  pageHeight: PAGE_HEIGHT_PT,
  columns: GRID_COLUMNS,
  rows: GRID_ROWS,
  marginX: MARGIN_X,
  marginY: TOP_OFFSET_PT,
  gutterX: GUTTER_X,
  gutterY: GUTTER_Y,
  cellWidth: CELL_WIDTH_PT,
  cellHeight: LABEL_AREA_HEIGHT_PT / GRID_ROWS,
  innerPadding: INNER_PADDING,
  fontSize: FONT_SIZE,
  lineHeight: LINE_HEIGHT,
  maxLines: Math.floor(
    (LABEL_AREA_HEIGHT_PT / GRID_ROWS - INNER_PADDING * 2) / LINE_HEIGHT,
  ),
  borderWidth: 0.75,
};

export const FORM_LAYOUT_CONFIG: FormLayoutConfig = {
  pageWidth: PAGE_WIDTH_PT,
  pageHeight: PAGE_HEIGHT_PT,
  fields: {
    customerDetails: {
      left: 33.732,
      top: 226.772,
      width: 143.646,
      height: 57.543,
      style: {
        fontKey: "medium",
        fontSize: 9,
        lineHeight: 12,
      },
    },
    enquiryId: {
      left: 33.732,
      top: 289.134,
      width: 143.646,
      height: 12.047,
      style: {
        fontKey: "bold",
        fontSize: 9,
        lineHeight: 13,
      },
    },
    returnByDate: {
      left: 355.039,
      top: 293.669,
      width: 92.835,
      height: 11.906,
      style: {
        fontKey: "medium",
        fontSize: 7.2,
        lineHeight: 12,
        uppercase: true,
      },
    },
    price: {
      left: 440.433,
      top: 293.669,
      width: 113.953,
      height: 11.906,
      style: {
        fontKey: "medium",
        fontSize: 7.2,
        lineHeight: 12,
        uppercase: true,
      },
    },
    nineCt: {
      left: 200.976,
      top: 265.89,
      width: 31.748,
      height: 14.173,
      style: {
        fontKey: "medium",
        fontSize: 8,
        lineHeight: 12,
      },
    },
    fourteenCt: {
      left: 235.843,
      top: 265.89,
      width: 31.748,
      height: 14.173,
      style: {
        fontKey: "medium",
        fontSize: 8,
        lineHeight: 12,
      },
    },
    eighteenCt: {
      left: 270.709,
      top: 265.89,
      width: 31.748,
      height: 14.173,
      style: {
        fontKey: "medium",
        fontSize: 8,
        lineHeight: 12,
      },
    },
    twentyTwoCt: {
      left: 305.575,
      top: 265.89,
      width: 31.748,
      height: 14.173,
      style: {
        fontKey: "medium",
        fontSize: 8,
        lineHeight: 12,
      },
    },
    twentyFourCt: {
      left: 340.441,
      top: 265.89,
      width: 31.748,
      height: 14.173,
      style: {
        fontKey: "medium",
        fontSize: 8,
        lineHeight: 12,
      },
    },
    sovereign: {
      left: 381.26,
      top: 265.89,
      width: 50.173,
      height: 14.173,
      style: {
        fontKey: "medium",
        fontSize: 8,
        lineHeight: 12,
      },
    },
    britiannia: {
      left: 440.504,
      top: 265.89,
      width: 50.173,
      height: 14.173,
      style: {
        fontKey: "medium",
        fontSize: 8,
        lineHeight: 12,
      },
    },
    krugerrand: {
      left: 509.102,
      top: 265.89,
      width: 50.173,
      height: 14.173,
      style: {
        fontKey: "medium",
        fontSize: 8,
        lineHeight: 12,
      },
    },
  },
};

export const LABELS_PER_PAGE =
  LABEL_LAYOUT_CONFIG.columns * LABEL_LAYOUT_CONFIG.rows;

export function normalizeCsvHeader(header: string): string {
  return header
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, "");
}

export function normalizeCsvValue(
  value: string | number | boolean | undefined | null,
): string {
  return String(value ?? "")
    .replace(/\r?\n/g, " ")
    .replace(/\u00a0/g, " ")
    .replace(/\u2028/g, " ")
    .replace(/\s*,\s*/g, ", ")
    .replace(/\s+/g, " ")
    .trim();
}

export function formatUkPostcode(value: string): string {
  const compact = value.toUpperCase().replace(/\s+/g, "").trim();

  if (!compact) {
    return "";
  }

  if (compact.length >= 5 && compact.length <= 7) {
    return `${compact.slice(0, -3)} ${compact.slice(-3)}`;
  }

  return value.toUpperCase().replace(/\s+/g, " ").trim();
}

export function formatLabelText(value: string): string {
  return value.toLowerCase().replace(/[a-z]+(?:'[a-z]+)*/g, (segment) => {
    return `${segment.charAt(0).toUpperCase()}${segment.slice(1)}`;
  });
}

export function findMissingRequiredHeaders(
  headers: string[],
): RequiredCsvHeader[] {
  const normalizedHeaders = new Set(headers.map(normalizeCsvHeader));

  return REQUIRED_CSV_HEADERS.filter(
    (header) => !normalizedHeaders.has(header),
  );
}

export function paginateLabels(labels: AddressLabel[]): AddressLabel[][] {
  const pages: AddressLabel[][] = [];

  for (let index = 0; index < labels.length; index += LABELS_PER_PAGE) {
    pages.push(labels.slice(index, index + LABELS_PER_PAGE));
  }

  return pages;
}

export async function prepareLabelsForOutput(
  input: CsvPreparationInput,
): Promise<PreparedLabelResult> {
  const missingHeaders = findMissingRequiredHeaders(input.headers);

  if (missingHeaders.length > 0) {
    return {
      labels: [],
      pages: [],
      forms: [],
      formPageCount: 0,
      missingHeaders,
      skippedRows: [],
      emptyRowCount: 0,
      totalRowCount: input.rows.length,
      warningCount: 0,
    };
  }

  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const labels: AddressLabel[] = [];
  const forms: PreparedFormRow[] = [];
  const skippedRows: SkippedCsvRow[] = [];
  let emptyRowCount = 0;

  for (const [index, row] of input.rows.entries()) {
    const normalizedRow = normalizeCsvRow(row);

    if (isRowCompletelyEmpty(normalizedRow)) {
      emptyRowCount += 1;
      continue;
    }

    const label = createAddressLabel(normalizedRow, font);

    if (!label) {
      skippedRows.push({
        rowNumber: index + 2,
        reason: "Missing delivery lines after normalization",
      });
      continue;
    }

    labels.push(label);
    forms.push(createPreparedFormRow(normalizedRow));
  }

  const pages = paginateLabels(labels);
  const warningCount = labels.reduce(
    (count, label) => count + label.warnings.length,
    0,
  );

  return {
    labels,
    pages,
    forms,
    formPageCount: forms.length,
    missingHeaders: [],
    skippedRows,
    emptyRowCount,
    totalRowCount: input.rows.length,
    warningCount,
  };
}

export async function generateLabelPdf(
  labels: AddressLabel[],
  options: GenerateLabelPdfOptions = {},
): Promise<ArrayBuffer> {
  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);

  const fonts = await embedLabelFonts(pdf, options.fontData);
  const pages = paginateLabels(labels);

  if (pages.length === 0) {
    pages.push([]);
  }

  for (const pageLabels of pages) {
    const page = pdf.addPage([
      LABEL_LAYOUT_CONFIG.pageWidth,
      LABEL_LAYOUT_CONFIG.pageHeight,
    ]);

    drawGrid(page);

    for (const [index, label] of pageLabels.entries()) {
      const position = getLabelCellPosition(index);
      const baseY =
        position.top -
        LABEL_LAYOUT_CONFIG.innerPadding -
        LABEL_LAYOUT_CONFIG.fontSize;

      for (const [lineIndex, line] of label.displayLines.entries()) {
        page.drawText(line, {
          x: position.left + LABEL_LAYOUT_CONFIG.innerPadding,
          y: baseY - lineIndex * LABEL_LAYOUT_CONFIG.lineHeight,
          size: LABEL_LAYOUT_CONFIG.fontSize,
          font: lineIndex === 0 ? fonts.name : fonts.address,
          color: rgb(0.08, 0.08, 0.08),
        });
      }
    }
  }

  const pdfBytes = await pdf.save();

  return new Uint8Array(pdfBytes).buffer as ArrayBuffer;
}

export async function generateFormPdf(
  forms: PreparedFormRow[],
  options: GenerateFormPdfOptions = {},
): Promise<ArrayBuffer> {
  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);

  const formFonts = await embedFormFonts(pdf, options.fontData);
  const pagesToRender = forms.length === 0 ? [null] : forms;

  for (const form of pagesToRender) {
    const page = pdf.addPage([
      FORM_LAYOUT_CONFIG.pageWidth,
      FORM_LAYOUT_CONFIG.pageHeight,
    ]);

    if (!form) {
      continue;
    }

    drawFormPage(page, form, formFonts, options.globalOffsetMm);
  }

  const pdfBytes = await pdf.save();

  return new Uint8Array(pdfBytes).buffer as ArrayBuffer;
}

function normalizeCsvRow(
  row: Record<string, string | undefined>,
): Record<string, string> {
  const normalized: Record<string, string> = {};

  for (const [header, value] of Object.entries(row)) {
    normalized[normalizeCsvHeader(header)] = normalizeCsvValue(value);
  }

  return normalized;
}

function isRowCompletelyEmpty(row: Record<string, string>): boolean {
  return Object.values(row).every((value) => value.length === 0);
}

function createAddressLabel(
  row: Record<string, string>,
  font: PDFFont,
): AddressLabel | null {
  const name = formatLabelText(
    [row.firstname ?? "", row.lastname ?? ""].filter(Boolean).join(" ").trim(),
  );
  const addressLines = [
    row.address1 ?? "",
    row.address2 ?? "",
    row.address3 ?? "",
    row.town ?? "",
  ]
    .filter(Boolean)
    .map(formatLabelText);
  const postcode = formatUkPostcode(row.postcode ?? "");
  const destinationLineCount = addressLines.length + (postcode ? 1 : 0);

  if (name.length === 0 || destinationLineCount === 0) {
    return null;
  }

  const rawDisplayLines = [name, ...addressLines, postcode].filter(Boolean);
  const wrappedLines = rawDisplayLines.flatMap((line) =>
    wrapTextToWidth(
      line,
      font,
      LABEL_LAYOUT_CONFIG.fontSize,
      getMaxTextWidth(),
    ),
  );

  const warnings: string[] = [];
  let displayLines = wrappedLines;

  if (wrappedLines.length > LABEL_LAYOUT_CONFIG.maxLines) {
    warnings.push("Label text was truncated to fit the sheet");
    displayLines = wrappedLines.slice(0, LABEL_LAYOUT_CONFIG.maxLines);
    displayLines[displayLines.length - 1] = ellipsizeTextToWidth(
      displayLines[displayLines.length - 1],
      font,
      LABEL_LAYOUT_CONFIG.fontSize,
      getMaxTextWidth(),
    );
  }

  return {
    name,
    lines: addressLines,
    postcode,
    displayLines,
    warnings,
  };
}

function createPreparedFormRow(row: Record<string, string>): PreparedFormRow {
  const customerName = formatLabelText(
    [row.firstname ?? "", row.lastname ?? ""].filter(Boolean).join(" ").trim(),
  );
  const countyLine = formatLabelText(row.county ?? "");
  const addressLines = [
    row.address1 ?? "",
    row.address2 ?? "",
    row.address3 ?? "",
    row.town ?? "",
    countyLine,
  ]
    .filter(Boolean)
    .map(formatLabelText);

  return {
    customerDetails: {
      nameLine: customerName ? `${customerName},` : "",
      addressLines,
      countyLine,
      postcode: formatUkPostcode(row.postcode ?? ""),
    },
    enquiryId: formatEnquiryId(row.enquiryid ?? ""),
    returnByDate: formatReturnByDate(row.returnbydate ?? ""),
    price: {
      label: "Insurance Value",
      value: formatCurrency(row.price ?? ""),
    },
    nineCt: formatCurrency(row["9ct"] ?? ""),
    fourteenCt: formatCurrency(row["14ct"] ?? ""),
    eighteenCt: formatCurrency(row["18ct"] ?? ""),
    twentyTwoCt: formatCurrency(row["22ct"] ?? ""),
    twentyFourCt: formatCurrency(row["24ct"] ?? ""),
    sovereign: formatCurrency(row.sovereign ?? ""),
    britiannia: formatCurrency(row["1ozbritiannia"] ?? ""),
    krugerrand: formatCurrency(row.goldkrugerrand ?? ""),
  };
}

function formatEnquiryId(value: string): string {
  return value ? `Enquiry ID: ${value}` : "";
}

function formatCurrency(value: string): string {
  const normalized = value.replace(/[^\d,.-]/g, "").trim();

  if (!normalized) {
    return "";
  }

  const numericValue = Number.parseFloat(normalized.replace(/,/g, ""));

  if (Number.isNaN(numericValue)) {
    return value.trim();
  }

  return GBP_FORMATTER.format(numericValue);
}

function formatReturnByDate(value: string): string {
  const trimmedValue = value.trim();

  if (!trimmedValue) {
    return "";
  }

  const date = parseDateValue(trimmedValue);

  if (!date) {
    return trimmedValue;
  }

  const weekday = ["Sun", "Mon", "Tues", "Wed", "Thurs", "Fri", "Sat"][
    date.getUTCDay()
  ];
  const month = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ][date.getUTCMonth()];
  const day = date.getUTCDate();
  const year = String(date.getUTCFullYear()).slice(-2);

  return `${weekday} ${day}${getOrdinalSuffix(day)} ${month} ${year}`;
}

function parseDateValue(value: string): Date | null {
  const ukDateMatch = value.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);

  if (ukDateMatch) {
    const [, dayPart, monthPart, yearPart] = ukDateMatch;
    const year =
      yearPart.length === 2 ? Number(`20${yearPart}`) : Number(yearPart);
    const month = Number(monthPart);
    const day = Number(dayPart);

    return new Date(Date.UTC(year, month - 1, day));
  }

  const parsedValue = Date.parse(value);

  if (Number.isNaN(parsedValue)) {
    return null;
  }

  return new Date(parsedValue);
}

function getOrdinalSuffix(day: number): string {
  if (day >= 11 && day <= 13) {
    return "th";
  }

  switch (day % 10) {
    case 1:
      return "st";
    case 2:
      return "nd";
    case 3:
      return "rd";
    default:
      return "th";
  }
}

async function embedFormFonts(
  pdf: PDFDocument,
  fontDataOverride?: Partial<Record<FormFontKey, ArrayBuffer | Uint8Array>>,
): Promise<Record<FormFontKey, PDFFont>> {
  const fontData = await resolveFormFontData(fontDataOverride);

  return {
    regular: await pdf.embedFont(fontData.regular),
    medium: await pdf.embedFont(fontData.medium),
    bold: await pdf.embedFont(fontData.bold),
  };
}

async function embedLabelFonts(
  pdf: PDFDocument,
  fontDataOverride?: Partial<Record<FormFontKey, ArrayBuffer | Uint8Array>>,
): Promise<LabelFonts> {
  const fonts = await embedFormFonts(pdf, fontDataOverride);

  return {
    name: fonts.medium,
    address: fonts.regular,
  };
}

async function resolveFormFontData(
  fontDataOverride?: Partial<Record<FormFontKey, ArrayBuffer | Uint8Array>>,
): Promise<FormFontData> {
  if (
    fontDataOverride?.regular &&
    fontDataOverride.medium &&
    fontDataOverride.bold
  ) {
    return {
      regular: toUint8Array(fontDataOverride.regular),
      medium: toUint8Array(fontDataOverride.medium),
      bold: toUint8Array(fontDataOverride.bold),
    };
  }

  if (!cachedFormFontDataPromise) {
    cachedFormFontDataPromise = loadFormFontData();
  }

  return cachedFormFontDataPromise;
}

async function loadFormFontData(): Promise<FormFontData> {
  const entries = await Promise.all(
    Object.entries(FORM_FONT_PATHS).map(async ([fontKey, fontPath]) => {
      const response = await fetch(fontPath);

      if (!response.ok) {
        throw new Error(`Unable to load form font: ${fontPath}`);
      }

      return [fontKey, new Uint8Array(await response.arrayBuffer())] as const;
    }),
  );

  return Object.fromEntries(entries) as FormFontData;
}

function toUint8Array(value: ArrayBuffer | Uint8Array): Uint8Array {
  return value instanceof Uint8Array ? value : new Uint8Array(value);
}

function drawFormPage(
  page: PDFPage,
  form: PreparedFormRow,
  fonts: Record<FormFontKey, PDFFont>,
  globalOffsetMm: FormCoordinateOffset = FORM_GLOBAL_OFFSET_MM,
) {
  const fields = getResolvedFormFieldLayouts(globalOffsetMm);

  drawCustomerDetails(
    page,
    form.customerDetails,
    fields.customerDetails,
    fonts,
  );
  drawSingleLineField(page, form.enquiryId, fields.enquiryId, fonts);
  drawSingleLineField(page, form.returnByDate, fields.returnByDate, fonts);
  drawPriceField(page, form.price, fields.price, fonts);
  drawSingleLineField(page, form.nineCt, fields.nineCt, fonts);
  drawSingleLineField(page, form.fourteenCt, fields.fourteenCt, fonts);
  drawSingleLineField(page, form.eighteenCt, fields.eighteenCt, fonts);
  drawSingleLineField(page, form.twentyTwoCt, fields.twentyTwoCt, fonts);
  drawSingleLineField(page, form.twentyFourCt, fields.twentyFourCt, fonts);
  drawSingleLineField(page, form.sovereign, fields.sovereign, fonts);
  drawSingleLineField(page, form.britiannia, fields.britiannia, fonts);
  drawSingleLineField(page, form.krugerrand, fields.krugerrand, fonts);
}

function getResolvedFormFieldLayouts(
  globalOffsetMm: FormCoordinateOffset = FORM_GLOBAL_OFFSET_MM,
): FormLayoutConfig["fields"] {
  return {
    customerDetails: getResolvedFormFieldLayout("customerDetails", globalOffsetMm),
    enquiryId: getResolvedFormFieldLayout("enquiryId", globalOffsetMm),
    returnByDate: getResolvedFormFieldLayout("returnByDate", globalOffsetMm),
    price: getResolvedFormFieldLayout("price", globalOffsetMm),
    nineCt: getResolvedFormFieldLayout("nineCt", globalOffsetMm),
    fourteenCt: getResolvedFormFieldLayout("fourteenCt", globalOffsetMm),
    eighteenCt: getResolvedFormFieldLayout("eighteenCt", globalOffsetMm),
    twentyTwoCt: getResolvedFormFieldLayout("twentyTwoCt", globalOffsetMm),
    twentyFourCt: getResolvedFormFieldLayout("twentyFourCt", globalOffsetMm),
    sovereign: getResolvedFormFieldLayout("sovereign", globalOffsetMm),
    britiannia: getResolvedFormFieldLayout("britiannia", globalOffsetMm),
    krugerrand: getResolvedFormFieldLayout("krugerrand", globalOffsetMm),
  };
}

export function getResolvedFormFieldLayout(
  fieldName: FormFieldName,
  globalOffsetMm: FormCoordinateOffset = FORM_GLOBAL_OFFSET_MM,
): FormFieldLayout {
  const layout = FORM_LAYOUT_CONFIG.fields[fieldName];
  const fieldOffset = FORM_FIELD_OFFSETS_MM[fieldName] ?? { x: 0, y: 0 };

  return offsetFormFieldLayout(layout, {
    x: globalOffsetMm.x + fieldOffset.x,
    y: globalOffsetMm.y + fieldOffset.y,
  });
}

function offsetFormFieldLayout(
  layout: FormFieldLayout,
  offset: FormCoordinateOffset,
): FormFieldLayout {
  return {
    ...layout,
    left: layout.left + mmToPoints(offset.x),
    top: layout.top + mmToPoints(offset.y),
  };
}

function drawCustomerDetails(
  page: PDFPage,
  details: PreparedFormCustomerDetails,
  layout: FormFieldLayout,
  fonts: Record<FormFontKey, PDFFont>,
) {
  const font = fonts[layout.style.fontKey];
  const lines = getFormCustomerDetailsLines(details, font, layout);

  for (const [index, line] of lines.entries()) {
    drawText(page, line, layout.left, layout.top, layout.style, font, index);
  }
}

export function getFormCustomerDetailsLines(
  details: PreparedFormCustomerDetails,
  font: PDFFont,
  layout: FormFieldLayout = FORM_LAYOUT_CONFIG.fields.customerDetails,
): string[] {
  const maxLineCount = getMaxFormLineCount(layout);
  const baseAddressLines = details.addressLines;
  const compactAddressLines = removeCountyLine(
    details.addressLines,
    details.countyLine,
  );
  const candidateAddressLines = dedupeAddressLineCandidates([
    baseAddressLines,
    compactAddressLines,
    ...buildMergedAddressLineCandidates(compactAddressLines),
    ...buildMergedAddressLineCandidates(baseAddressLines),
  ]);

  for (const addressLines of candidateAddressLines) {
    const renderedLines = buildRenderedCustomerDetailLines(
      details.nameLine,
      addressLines,
      details.postcode,
      font,
      layout,
    );

    if (renderedLines.length <= maxLineCount) {
      return renderedLines;
    }
  }

  const fallbackAddressLines =
    candidateAddressLines[candidateAddressLines.length - 1] ?? [];

  return truncateCustomerDetailLines(
    buildRenderedCustomerDetailLines(
      details.nameLine,
      fallbackAddressLines,
      details.postcode,
      font,
      layout,
    ),
    details.postcode,
    font,
    layout,
    maxLineCount,
  );
}

function drawPriceField(
  page: PDFPage,
  price: PreparedFormPrice,
  layout: FormFieldLayout,
  fonts: Record<FormFontKey, PDFFont>,
) {
  const labelFont = fonts.bold;
  const valueFont = fonts[layout.style.fontKey];
  const transformedLabel = applyTextTransform(price.label, layout.style);
  const transformedValue = applyTextTransform(price.value, layout.style);
  const y = toPdfY(layout.top, layout.style.fontSize);

  page.drawText(transformedLabel, {
    x: layout.left,
    y,
    size: layout.style.fontSize,
    font: labelFont,
    color: rgb(0, 0, 0),
  });

  page.drawText(transformedValue, {
    x:
      layout.left +
      measureTextWidth(transformedLabel, labelFont, layout.style.fontSize) +
      3,
    y,
    size: layout.style.fontSize,
    font: valueFont,
    color: rgb(0, 0, 0),
  });
}

function drawSingleLineField(
  page: PDFPage,
  value: string,
  layout: FormFieldLayout,
  fonts: Record<FormFontKey, PDFFont>,
) {
  const font = fonts[layout.style.fontKey];

  drawText(page, value, layout.left, layout.top, layout.style, font);
}

function drawText(
  page: PDFPage,
  value: string,
  left: number,
  top: number,
  style: FormTextStyle,
  font: PDFFont,
  lineIndex = 0,
) {
  page.drawText(applyTextTransform(value, style), {
    x: left,
    y: toPdfY(top, style.fontSize) - lineIndex * style.lineHeight,
    size: style.fontSize,
    font,
    color: rgb(0, 0, 0),
  });
}

function applyTextTransform(value: string, style: FormTextStyle): string {
  return style.uppercase ? value.toUpperCase() : value;
}

function toPdfY(top: number, fontSize: number): number {
  return FORM_LAYOUT_CONFIG.pageHeight - top - fontSize;
}

function mmToPoints(value: number): number {
  return value * POINTS_PER_MM;
}

function buildRenderedCustomerDetailLines(
  nameLine: string,
  addressLines: string[],
  postcode: string,
  font: PDFFont,
  layout: FormFieldLayout,
): string[] {
  return [
    nameLine,
    ...addressLines.flatMap((line) =>
      wrapTextToWidth(line, font, layout.style.fontSize, layout.width),
    ),
    postcode,
  ].filter(Boolean);
}

function getMaxFormLineCount(layout: FormFieldLayout): number {
  return Math.max(
    1,
    Math.floor((layout.height - layout.style.fontSize) / layout.style.lineHeight) + 1,
  );
}

function removeCountyLine(addressLines: string[], countyLine: string): string[] {
  if (!countyLine) {
    return [...addressLines];
  }

  const countyIndex = addressLines.lastIndexOf(countyLine);

  if (countyIndex === -1) {
    return [...addressLines];
  }

  return [
    ...addressLines.slice(0, countyIndex),
    ...addressLines.slice(countyIndex + 1),
  ];
}

function buildMergedAddressLineCandidates(addressLines: string[]): string[][] {
  const candidates: string[][] = [];
  let mergedLines = [...addressLines];

  while (mergedLines.length > 1) {
    mergedLines = [
      ...mergedLines.slice(0, -2),
      joinAddressLines(
        mergedLines[mergedLines.length - 2] ?? "",
        mergedLines[mergedLines.length - 1] ?? "",
      ),
    ];
    candidates.push(mergedLines);
  }

  return candidates;
}

function joinAddressLines(left: string, right: string): string {
  if (!left) {
    return right;
  }

  if (!right) {
    return left;
  }

  if (left.endsWith(",") || right.startsWith(",")) {
    return `${left} ${right}`;
  }

  return `${left}, ${right}`;
}

function dedupeAddressLineCandidates(candidates: string[][]): string[][] {
  const seen = new Set<string>();

  return candidates.filter((lines) => {
    const key = lines.join("\u001f");

    if (seen.has(key)) {
      return false;
    }

    seen.add(key);
    return true;
  });
}

function truncateCustomerDetailLines(
  lines: string[],
  postcode: string,
  font: PDFFont,
  layout: FormFieldLayout,
  maxLineCount: number,
): string[] {
  if (lines.length <= maxLineCount) {
    return lines;
  }

  if (postcode && maxLineCount > 1 && lines.at(-1) === postcode) {
    const visibleLines = lines.slice(0, maxLineCount - 1);

    if (visibleLines.length > 0) {
      visibleLines[visibleLines.length - 1] = ellipsizeTextToWidth(
        visibleLines[visibleLines.length - 1] ?? "",
        font,
        layout.style.fontSize,
        layout.width,
      );
    }

    return [...visibleLines, postcode];
  }

  const truncatedLines = lines.slice(0, maxLineCount);

  truncatedLines[truncatedLines.length - 1] = ellipsizeTextToWidth(
    truncatedLines[truncatedLines.length - 1] ?? "",
    font,
    layout.style.fontSize,
    layout.width,
  );

  return truncatedLines;
}

function wrapTextToWidth(
  text: string,
  font: PDFFont,
  fontSize: number,
  maxWidth: number,
): string[] {
  const normalizedText = text.replace(/\s+/g, " ").trim();

  if (!normalizedText) {
    return [];
  }

  const tokens = normalizedText.split(" ");
  const lines: string[] = [];
  let currentLine = "";

  for (const token of tokens) {
    const parts =
      measureTextWidth(token, font, fontSize) <= maxWidth
        ? [token]
        : splitLongToken(token, font, fontSize, maxWidth);

    for (const [partIndex, part] of parts.entries()) {
      const candidate =
        currentLine.length === 0
          ? part
          : partIndex === 0
            ? `${currentLine} ${part}`
            : `${currentLine}${part}`;

      if (measureTextWidth(candidate, font, fontSize) <= maxWidth) {
        currentLine = candidate;
        continue;
      }

      if (currentLine.length > 0) {
        lines.push(currentLine);
      }

      currentLine = part;
    }
  }

  if (currentLine.length > 0) {
    lines.push(currentLine);
  }

  return lines;
}

function splitLongToken(
  token: string,
  font: PDFFont,
  fontSize: number,
  maxWidth: number,
): string[] {
  const segments: string[] = [];
  let current = "";

  for (const character of token) {
    const candidate = `${current}${character}`;

    if (
      current.length === 0 ||
      measureTextWidth(candidate, font, fontSize) <= maxWidth
    ) {
      current = candidate;
      continue;
    }

    segments.push(current);
    current = character;
  }

  if (current.length > 0) {
    segments.push(current);
  }

  return segments;
}

function ellipsizeTextToWidth(
  text: string,
  font: PDFFont,
  fontSize: number,
  maxWidth: number,
): string {
  let truncated = text.trimEnd();

  while (truncated.length > 0) {
    const candidate = `${truncated}…`;

    if (measureTextWidth(candidate, font, fontSize) <= maxWidth) {
      return candidate;
    }

    truncated = truncated.slice(0, -1);
  }

  return "…";
}

function drawGrid(page: PDFPage) {
  if (!SHOW_LABEL_OUTLINES) {
    return;
  }

  for (let row = 0; row < LABEL_LAYOUT_CONFIG.rows; row += 1) {
    for (let column = 0; column < LABEL_LAYOUT_CONFIG.columns; column += 1) {
      const index = row * LABEL_LAYOUT_CONFIG.columns + column;
      const position = getLabelCellPosition(index);

      page.drawRectangle({
        x: position.left,
        y: position.bottom,
        width: LABEL_LAYOUT_CONFIG.cellWidth,
        height: LABEL_LAYOUT_CONFIG.cellHeight,
        borderColor: rgb(0.82, 0.82, 0.82),
        borderWidth: LABEL_LAYOUT_CONFIG.borderWidth,
      });
    }
  }
}

function getColumnOffset(column: number) {
  let offset = 0;

  for (let currentColumn = 0; currentColumn < column; currentColumn += 1) {
    offset += LABEL_LAYOUT_CONFIG.cellWidth;

    if (currentColumn < COLUMN_GAPS_PT.length) {
      offset += COLUMN_GAPS_PT[currentColumn];
    }
  }

  return offset;
}

export function getLabelCellPosition(index: number) {
  const column = index % LABEL_LAYOUT_CONFIG.columns;
  const row = Math.floor(index / LABEL_LAYOUT_CONFIG.columns);
  const left = LABEL_LAYOUT_CONFIG.marginX + getColumnOffset(column);
  const top =
    LABEL_LAYOUT_CONFIG.pageHeight -
    TOP_OFFSET_PT -
    row * (LABEL_LAYOUT_CONFIG.cellHeight + LABEL_LAYOUT_CONFIG.gutterY);

  return {
    left,
    top,
    bottom: top - LABEL_LAYOUT_CONFIG.cellHeight,
  };
}

function getMaxTextWidth() {
  return LABEL_LAYOUT_CONFIG.cellWidth - LABEL_LAYOUT_CONFIG.innerPadding * 2;
}

function measureTextWidth(text: string, font: PDFFont, fontSize: number) {
  return font.widthOfTextAtSize(text, fontSize);
}
