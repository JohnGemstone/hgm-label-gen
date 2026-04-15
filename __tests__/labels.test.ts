import { describe, expect, test } from "vitest";
import {
  findMissingRequiredHeaders,
  LABEL_LAYOUT_CONFIG,
  normalizeCsvHeader,
  paginateLabels,
  prepareLabelsForOutput,
  REQUIRED_CSV_HEADERS,
  type AddressLabel,
} from "@/lib/labels";

describe("label utilities", () => {
  test("normalizes headers and reports missing required columns", () => {
    expect(normalizeCsvHeader(" Address_Line_1 ")).toBe("address line 1");

    expect(
      findMissingRequiredHeaders([
        "Name",
        "Address Line 1",
        "Address Line 2",
        "Address Line 3",
        "Town",
      ]),
    ).toEqual(["postcode"]);
  });

  test("filters empty rows and collapses blank address lines", async () => {
    const result = await prepareLabelsForOutput({
      headers: [...REQUIRED_CSV_HEADERS],
      rows: [
        {
          name: "",
          "address line 1": "",
          "address line 2": "",
          "address line 3": "",
          town: "",
          postcode: "",
        },
        {
          name: "Alice Example",
          "address line 1": "10 Downing Street",
          "address line 2": "",
          "address line 3": " ",
          town: "London",
          postcode: "SW1A 2AA",
        },
      ],
    });

    expect(result.emptyRowCount).toBe(1);
    expect(result.labels).toHaveLength(1);
    expect(result.labels[0]?.lines).toEqual(["10 Downing Street", "London"]);
    expect(result.labels[0]?.displayLines).toEqual([
      "Alice Example",
      "10 Downing Street",
      "London",
      "SW1A 2AA",
    ]);
  });

  test("splits labels into pages of twenty four", async () => {
    const rows = Array.from({ length: 25 }, (_, index) => ({
      name: `Customer ${index + 1}`,
      "address line 1": `${index + 1} Test Street`,
      "address line 2": "",
      "address line 3": "",
      town: "London",
      postcode: `SW1A ${index + 1}`,
    }));

    const result = await prepareLabelsForOutput({
      headers: [...REQUIRED_CSV_HEADERS],
      rows,
    });

    expect(result.labels).toHaveLength(25);
    expect(result.pages).toHaveLength(2);
    expect(result.pages[0]).toHaveLength(24);
    expect(result.pages[1]).toHaveLength(1);
  });

  test("flags and truncates labels that overflow vertically", async () => {
    const result = await prepareLabelsForOutput({
      headers: [...REQUIRED_CSV_HEADERS],
      rows: [
        {
          name: "A very long recipient name that should wrap across lines",
          "address line 1":
            "123 Extremely Long Street Name That Keeps Going Without Much Mercy",
          "address line 2":
            "Industrial Estate Building Name With Additional Location Details",
          "address line 3":
            "Distribution Hub Access Corridor West Entrance Loading Bay Three",
          town: "Hatton Garden Metropolitan Delivery District",
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
});
