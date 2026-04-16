import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, test, vi } from "vitest";
import * as labelsModule from "@/lib/labels";
import LabelGenerator from "@/app/components/label-generator";

const COMPLETE_HEADER = [
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
].join(",");

const COMPLETE_ROW = [
  "449587",
  "Alice",
  "Example",
  "10 Downing Street",
  "",
  "",
  "London",
  "",
  "SW1A 2AA",
  "£2,500.00",
  "£35.70",
  "£55.69",
  "£71.40",
  "£87.20",
  "£95.20",
  "£803.56",
  "£3,418.84",
  "£3,381.40",
  "20/04/2026",
].join(",");

afterEach(() => {
  vi.restoreAllMocks();
});

describe("LabelGenerator", () => {
  test("uploading a valid CSV enables data review and PDF generation", async () => {
    const user = userEvent.setup();

    render(<LabelGenerator />);

    await user.upload(
      screen.getByLabelText(/select a/i),
      new File(
        [[COMPLETE_HEADER, COMPLETE_ROW].join("\n")],
        "customers.csv",
        { type: "text/csv" },
      ),
    );

    const generateButton = screen.getByRole("button", {
      name: /generate pdfs/i,
    }) as HTMLButtonElement;

    await waitFor(() => {
      expect(generateButton.disabled).toBe(false);
    });

    expect(screen.getByTestId("review-data-table")).toBeDefined();
    expect(screen.getByRole("heading", { name: /review data/i })).toBeDefined();
    expect(screen.getByTestId("valid-label-count").textContent).toBe("1");
    expect(screen.getByTestId("page-count").textContent).toBe("1");
    expect(screen.getByTestId("form-page-count").textContent).toBe("1");
    expect(screen.getByText("Alice Example")).toBeDefined();
    expect(screen.getByText("10 Downing Street, London, SW1A 2AA")).toBeDefined();
  });

  test("shows a blocking error when required headers are missing", async () => {
    const user = userEvent.setup();

    render(<LabelGenerator />);

    await user.upload(
      screen.getByLabelText(/select a/i),
      new File(
        [
          [
            COMPLETE_HEADER.replace(",Postcode", ""),
            COMPLETE_ROW.replace(",SW1A 2AA", ""),
          ].join("\n"),
        ],
        "missing-postcode.csv",
        { type: "text/csv" },
      ),
    );

    await waitFor(() => {
      expect(
        screen.getByText(/missing required columns: postcode/i),
      ).toBeDefined();
    });

    const generateButton = screen.getByRole("button", {
      name: /generate pdfs/i,
    }) as HTMLButtonElement;

    expect(generateButton.disabled).toBe(true);
  });

  test("reports the correct page count for multi-page CSV uploads", async () => {
    const user = userEvent.setup();
    const rows = Array.from(
      { length: 25 },
      (_, index) =>
        [
          `${449587 + index}`,
          "Customer",
          `${index + 1}`,
          `${index + 1} Test Street`,
          "",
          "",
          "London",
          "",
          `EC1N ${index + 1}`,
          "£2,500.00",
          "£35.70",
          "£55.69",
          "£71.40",
          "£87.20",
          "£95.20",
          "£803.56",
          "£3,418.84",
          "£3,381.40",
          "20/04/2026",
        ].join(","),
    );

    render(<LabelGenerator />);

    await user.upload(
      screen.getByLabelText(/select a/i),
      new File([[COMPLETE_HEADER, ...rows].join("\n")], "bulk.csv", {
        type: "text/csv",
      }),
    );

    await waitFor(() => {
      expect(screen.getByTestId("valid-label-count").textContent).toBe("25");
    });

    expect(screen.getByTestId("page-count").textContent).toBe("2");
    expect(screen.getByTestId("form-page-count").textContent).toBe("25");
  });

  test("renders separate generated label and form previews after generation", async () => {
    const user = userEvent.setup();
    vi.spyOn(labelsModule, "generateLabelPdf").mockResolvedValue(
      new ArrayBuffer(8),
    );
    vi.spyOn(labelsModule, "generateFormPdf").mockResolvedValue(new ArrayBuffer(8));
    Object.defineProperty(URL, "createObjectURL", {
      configurable: true,
      writable: true,
      value: vi
        .fn()
        .mockReturnValueOnce("blob:labels")
        .mockReturnValueOnce("blob:forms"),
    });
    Object.defineProperty(URL, "revokeObjectURL", {
      configurable: true,
      writable: true,
      value: vi.fn(),
    });
    Object.defineProperty(HTMLElement.prototype, "scrollIntoView", {
      configurable: true,
      writable: true,
      value: vi.fn(),
    });

    render(<LabelGenerator />);

    await user.upload(
      screen.getByLabelText(/select a/i),
      new File(
        [[COMPLETE_HEADER, COMPLETE_ROW].join("\n")],
        "customers.csv",
        { type: "text/csv" },
      ),
    );

    await user.click(screen.getByRole("button", { name: /generate pdfs/i }));

    await waitFor(() => {
      expect(
        screen.getByRole("heading", { name: /generated labels/i }),
      ).toBeDefined();
    });

    expect(
      screen.getByRole("heading", { name: /generated forms/i }),
    ).toBeDefined();
    expect(screen.getByTestId("label-pdf-preview")).toBeDefined();
    expect(screen.getByTestId("form-pdf-preview")).toBeDefined();
  });
});
