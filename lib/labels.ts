import {
  PDFDocument,
  PDFPage,
  StandardFonts,
  type PDFFont,
  rgb,
} from "pdf-lib";

export const REQUIRED_CSV_HEADERS = [
  "name",
  "address line 1",
  "address line 2",
  "address line 3",
  "town",
  "postcode",
] as const;

export type RequiredCsvHeader = (typeof REQUIRED_CSV_HEADERS)[number];

export interface AddressLabel {
  name: string;
  lines: string[];
  postcode: string;
  displayLines: string[];
  warnings: string[];
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

const PAGE_WIDTH_PT = 595.28;
const PAGE_HEIGHT_PT = 841.89;
const GRID_COLUMNS = 3;
const GRID_ROWS = 8;
const MARGIN_X = 24;
const MARGIN_Y = 24;
const GUTTER_X = 6;
const GUTTER_Y = 6;
const INNER_PADDING = 8;
const FONT_SIZE = 10;
const LINE_HEIGHT = 12;

export const LABEL_LAYOUT_CONFIG: LabelLayoutConfig = {
  pageWidth: PAGE_WIDTH_PT,
  pageHeight: PAGE_HEIGHT_PT,
  columns: GRID_COLUMNS,
  rows: GRID_ROWS,
  marginX: MARGIN_X,
  marginY: MARGIN_Y,
  gutterX: GUTTER_X,
  gutterY: GUTTER_Y,
  cellWidth:
    (PAGE_WIDTH_PT - MARGIN_X * 2 - GUTTER_X * (GRID_COLUMNS - 1)) /
    GRID_COLUMNS,
  cellHeight:
    (PAGE_HEIGHT_PT - MARGIN_Y * 2 - GUTTER_Y * (GRID_ROWS - 1)) / GRID_ROWS,
  innerPadding: INNER_PADDING,
  fontSize: FONT_SIZE,
  lineHeight: LINE_HEIGHT,
  maxLines: Math.floor(
    (
      (PAGE_HEIGHT_PT - MARGIN_Y * 2 - GUTTER_Y * (GRID_ROWS - 1)) /
        GRID_ROWS -
      INNER_PADDING * 2
    ) / LINE_HEIGHT,
  ),
  borderWidth: 0.75,
};

export const LABELS_PER_PAGE =
  LABEL_LAYOUT_CONFIG.columns * LABEL_LAYOUT_CONFIG.rows;

export function normalizeCsvHeader(header: string): string {
  return header.trim().toLowerCase().replace(/[_\s]+/g, " ");
}

export function normalizeCsvValue(value: string | undefined): string {
  return (value ?? "").replace(/\r?\n/g, " ").trim();
}

export function findMissingRequiredHeaders(headers: string[]): RequiredCsvHeader[] {
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
  }

  const pages = paginateLabels(labels);
  const warningCount = labels.reduce(
    (count, label) => count + label.warnings.length,
    0,
  );

  return {
    labels,
    pages,
    missingHeaders: [],
    skippedRows,
    emptyRowCount,
    totalRowCount: input.rows.length,
    warningCount,
  };
}

export async function generateLabelPdf(
  labels: AddressLabel[],
): Promise<ArrayBuffer> {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
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
      const position = getCellPosition(index);
      const baseY =
        position.top -
        LABEL_LAYOUT_CONFIG.innerPadding -
        LABEL_LAYOUT_CONFIG.fontSize;

      for (const [lineIndex, line] of label.displayLines.entries()) {
        page.drawText(line, {
          x: position.left + LABEL_LAYOUT_CONFIG.innerPadding,
          y: baseY - lineIndex * LABEL_LAYOUT_CONFIG.lineHeight,
          size: LABEL_LAYOUT_CONFIG.fontSize,
          font,
          color: rgb(0.08, 0.08, 0.08),
        });
      }
    }
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
  const name = row.name ?? "";
  const addressLines = [
    row["address line 1"] ?? "",
    row["address line 2"] ?? "",
    row["address line 3"] ?? "",
    row.town ?? "",
  ].filter(Boolean);
  const postcode = row.postcode ?? "";
  const destinationLineCount = addressLines.length + (postcode ? 1 : 0);

  if (destinationLineCount === 0) {
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
  for (let row = 0; row < LABEL_LAYOUT_CONFIG.rows; row += 1) {
    for (let column = 0; column < LABEL_LAYOUT_CONFIG.columns; column += 1) {
      const index = row * LABEL_LAYOUT_CONFIG.columns + column;
      const position = getCellPosition(index);

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

function getCellPosition(index: number) {
  const column = index % LABEL_LAYOUT_CONFIG.columns;
  const row = Math.floor(index / LABEL_LAYOUT_CONFIG.columns);
  const left =
    LABEL_LAYOUT_CONFIG.marginX +
    column * (LABEL_LAYOUT_CONFIG.cellWidth + LABEL_LAYOUT_CONFIG.gutterX);
  const top =
    LABEL_LAYOUT_CONFIG.pageHeight -
    LABEL_LAYOUT_CONFIG.marginY -
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
