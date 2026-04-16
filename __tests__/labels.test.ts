import { readFileSync } from "node:fs";
import { join } from "node:path";
import { PDFDocument } from "pdf-lib";
import { describe, expect, test } from "vitest";
import {
  findMissingRequiredHeaders,
  formatLabelText,
  formatUkPostcode,
  generateFormPdf,
  LABEL_LAYOUT_CONFIG,
  normalizeCsvHeader,
  paginateLabels,
  prepareLabelsForOutput,
  REQUIRED_CSV_HEADERS,
  type AddressLabel,
} from "@/lib/labels";

const COMPLETE_ROW = {
  enquiryid: "449587",
  firstname: "alice",
  lastname: "example",
  address1: "10 downing street",
  address2: "",
  address3: "",
  town: "Rawcliffe,  york ",
  county: "",
  postcode: "sw1a2aa",
  price: "£2,500.00",
  "9ct": "£35.70",
  "14ct": "£55.69",
  "18ct": "£71.40",
  "22ct": "£87.20",
  "24ct": "£95.20",
  sovereign: "£803.56",
  "1ozbritiannia": "£3,418.84",
  goldkrugerrand: "£3,381.40",
  returnbydate: "20/04/2026",
} satisfies Record<string, string>;

describe("label utilities", () => {
  test("normalizes headers and reports missing required columns", () => {
    expect(normalizeCsvHeader(" Address_1 ")).toBe("address1");

    const missingHeaders = findMissingRequiredHeaders([
      "EnquiryID",
      "FirstName",
      "LastName",
      "Address1",
      "Address2",
      "Address3",
      "Town",
      "County",
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
    ]);

    expect(missingHeaders).toEqual(["postcode"]);
  });

  test("filters empty rows, creates forms, and collapses blank address lines", async () => {
    const result = await prepareLabelsForOutput({
      headers: [...REQUIRED_CSV_HEADERS],
      rows: [
        Object.fromEntries(
          REQUIRED_CSV_HEADERS.map((header) => [header, ""]),
        ) as Record<string, string>,
        COMPLETE_ROW,
      ],
    });

    expect(result.emptyRowCount).toBe(1);
    expect(result.labels).toHaveLength(1);
    expect(result.forms).toHaveLength(1);
    expect(result.formPageCount).toBe(1);
    expect(result.labels[0]?.name).toBe("Alice Example");
    expect(result.labels[0]?.lines).toEqual([
      "10 Downing Street",
      "Rawcliffe, York",
    ]);
    expect(result.labels[0]?.displayLines).toEqual([
      "Alice Example",
      "10 Downing Street",
      "Rawcliffe, York",
      "SW1A 2AA",
    ]);
    expect(result.forms[0]).toMatchObject({
      enquiryId: "Customer ID: 449587",
      returnByDate: "Mon 20th Apr 26",
      price: {
        label: "Insurance Value",
        value: "£2,500.00",
      },
      britiannia: "£3,418.84",
      krugerrand: "£3,381.40",
    });
    expect(result.forms[0]?.customerDetails).toMatchObject({
      nameLine: "Alice Example,",
      addressLines: ["10 Downing Street", "Rawcliffe, York"],
      postcode: "SW1A 2AA",
    });
  });

  test("formats manually entered uk postcodes with no space", () => {
    expect(formatUkPostcode("ts57eg")).toBe("TS5 7EG");
    expect(formatUkPostcode("YO30 5WG")).toBe("YO30 5WG");
  });

  test("capitalizes manually entered names and address text", () => {
    expect(formatLabelText("karine richardson")).toBe("Karine Richardson");
    expect(formatLabelText("37 st peter's road")).toBe("37 St Peter's Road");
    expect(formatLabelText("rawcliffe, york")).toBe("Rawcliffe, York");
  });

  test("splits labels into pages of twenty four and forms into one page per row", async () => {
    const rows = Array.from({ length: 25 }, (_, index) => ({
      ...COMPLETE_ROW,
      enquiryid: `${449587 + index}`,
      firstname: "Customer",
      lastname: `${index + 1}`,
      address1: `${index + 1} Test Street`,
      town: "London",
      postcode: `SW1A ${index + 1}`,
    }));

    const result = await prepareLabelsForOutput({
      headers: [...REQUIRED_CSV_HEADERS],
      rows,
    });

    expect(result.labels).toHaveLength(25);
    expect(result.forms).toHaveLength(25);
    expect(result.pages).toHaveLength(2);
    expect(result.pages[0]).toHaveLength(24);
    expect(result.pages[1]).toHaveLength(1);
    expect(result.formPageCount).toBe(25);
  });

  test("flags and truncates labels that overflow vertically", async () => {
    const result = await prepareLabelsForOutput({
      headers: [...REQUIRED_CSV_HEADERS],
      rows: [
        {
          ...COMPLETE_ROW,
          firstname:
            "A very long recipient name that should wrap across lines",
          lastname: "For Testing",
          address1:
            "123 Extremely Long Street Name That Keeps Going Without Much Mercy",
          address2:
            "Industrial Estate Building Name With Additional Location Details",
          address3:
            "Distribution Hub Access Corridor West Entrance Loading Bay Three",
          town: "Hatton Garden Metropolitan Delivery District",
          county: "North Administrative Region",
          postcode: "EC1N 8XX",
        },
      ],
    });

    expect(result.warningCount).toBe(1);
    expect(result.labels[0]?.warnings).toEqual([
      "Label text was truncated to fit the sheet",
    ]);
    expect(result.labels[0]?.displayLines).toHaveLength(
      LABEL_LAYOUT_CONFIG.maxLines,
    );
    expect(result.labels[0]?.displayLines.at(-1)?.endsWith("…")).toBe(true);
  });

  test("paginates arbitrary label arrays with the shared page size", () => {
    const labels = Array.from({ length: 26 }, (_, index) => ({
      name: `Name ${index + 1}`,
      lines: ["10 Test Street"],
      postcode: "EC1N 8XX",
      displayLines: [`Name ${index + 1}`, "10 Test Street", "EC1N 8XX"],
      warnings: [],
    })) satisfies AddressLabel[];

    const pages = paginateLabels(labels);

    expect(pages).toHaveLength(2);
    expect(pages[0]).toHaveLength(24);
    expect(pages[1]).toHaveLength(2);
  });

  test("generates a multi-page form PDF with embedded IBM Plex fonts", async () => {
    const result = await prepareLabelsForOutput({
      headers: [...REQUIRED_CSV_HEADERS],
      rows: [
        COMPLETE_ROW,
        {
          ...COMPLETE_ROW,
          enquiryid: "449588",
          firstname: "Tim",
          lastname: "Cowley",
        },
      ],
    });

    const pdfBytes = await generateFormPdf(result.forms, {
      fontData: {
        regular: readFileSync(
          join(process.cwd(), "public/fonts/IBMPlexSans-Regular.ttf"),
        ),
        medium: readFileSync(
          join(process.cwd(), "public/fonts/IBMPlexSans-Medium.ttf"),
        ),
        bold: readFileSync(
          join(process.cwd(), "public/fonts/IBMPlexSans-Bold.ttf"),
        ),
      },
    });
    const pdf = await PDFDocument.load(pdfBytes);

    expect(pdf.getPageCount()).toBe(2);
  });
});
