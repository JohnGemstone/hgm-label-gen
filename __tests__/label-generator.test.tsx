import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { describe, expect, test } from "vitest";
import LabelGenerator from "@/app/components/label-generator";

describe("LabelGenerator", () => {
  test("uploading a valid CSV enables data review and PDF generation", async () => {
    const user = userEvent.setup();

    render(<LabelGenerator />);

    await user.upload(
      screen.getByLabelText(/select a/i),
      new File(
        [
          [
            "name,address line 1,address line 2,address line 3,town,postcode",
            "Alice Example,10 Downing Street,,,London,SW1A 2AA",
          ].join("\n"),
        ],
        "customers.csv",
        { type: "text/csv" },
      ),
    );

    const generateButton = screen.getByRole("button", {
      name: /generate pdf/i,
    }) as HTMLButtonElement;

    await waitFor(() => {
      expect(generateButton.disabled).toBe(false);
    });

    expect(screen.getByTestId("review-data-table")).toBeDefined();
    expect(screen.getByRole("heading", { name: /review data/i })).toBeDefined();
    expect(screen.getByTestId("valid-label-count").textContent).toBe("1");
  });

  test("shows a blocking error when required headers are missing", async () => {
    const user = userEvent.setup();

    render(<LabelGenerator />);

    await user.upload(
      screen.getByLabelText(/select a/i),
      new File(
        [
          [
            "name,address line 1,address line 2,address line 3,town",
            "Alice Example,10 Downing Street,,,London",
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
      name: /generate pdf/i,
    }) as HTMLButtonElement;

    expect(generateButton.disabled).toBe(true);
  });

  test("reports the correct page count for multi-page CSV uploads", async () => {
    const user = userEvent.setup();
    const rows = Array.from(
      { length: 25 },
      (_, index) =>
        `Customer ${index + 1},${index + 1} Test Street,,,London,EC1N ${index + 1}`,
    );

    render(<LabelGenerator />);

    await user.upload(
      screen.getByLabelText(/select a/i),
      new File(
        [
          [
            "name,address line 1,address line 2,address line 3,town,postcode",
            ...rows,
          ].join("\n"),
        ],
        "bulk.csv",
        { type: "text/csv" },
      ),
    );

    await waitFor(() => {
      expect(screen.getByTestId("valid-label-count").textContent).toBe("25");
    });

    expect(screen.getByTestId("page-count").textContent).toBe("2");
  });
});
