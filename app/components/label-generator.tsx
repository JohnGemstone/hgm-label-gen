"use client";

import Papa from "papaparse";
import { useEffect, useRef, useState, type ChangeEvent } from "react";
import {
  generateLabelPdf,
  LABEL_LAYOUT_CONFIG,
  prepareLabelsForOutput,
  REQUIRED_CSV_HEADERS,
  type PreparedLabelResult,
} from "@/lib/labels";

type ParsedCsvRow = Record<string, string | undefined>;

export default function LabelGenerator() {
  const pdfSectionRef = useRef<HTMLElement | null>(null);
  const [selectedFileName, setSelectedFileName] = useState<string>("");
  const [result, setResult] = useState<PreparedLabelResult | null>(null);
  const [blockingError, setBlockingError] = useState<string | null>(null);
  const [parserWarnings, setParserWarnings] = useState<string[]>([]);
  const [isParsing, setIsParsing] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [pdfUrl, setPdfUrl] = useState<string | null>(null);
  const [pdfFileName, setPdfFileName] = useState("hgm-address-labels.pdf");

  useEffect(() => {
    return () => {
      if (pdfUrl) {
        URL.revokeObjectURL(pdfUrl);
      }
    };
  }, [pdfUrl]);

  useEffect(() => {
    if (!pdfUrl) {
      return;
    }

    pdfSectionRef.current?.scrollIntoView({
      behavior: "smooth",
      block: "start",
    });
  }, [pdfUrl]);

  const hasValidLabels =
    result !== null &&
    result.labels.length > 0 &&
    result.missingHeaders.length === 0 &&
    !blockingError;

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    setIsParsing(true);
    setSelectedFileName(file.name);
    setBlockingError(null);
    setParserWarnings([]);
    setResult(null);
    setPdfFileName(buildPdfFileName(file.name));
    setPdfUrl((currentUrl) => {
      if (currentUrl) {
        URL.revokeObjectURL(currentUrl);
      }

      return null;
    });

    try {
      const csvText = await readFileAsText(file);
      const parsed = Papa.parse<ParsedCsvRow>(csvText, {
        header: true,
        skipEmptyLines: "greedy",
      });

      const nextWarnings = parsed.errors.map(
        (error) => `Row ${error.row || "unknown"}: ${error.message}`,
      );
      const nextResult = await prepareLabelsForOutput({
        headers: parsed.meta.fields ?? [],
        rows: parsed.data,
      });

      setParserWarnings(nextWarnings);
      setResult(nextResult);

      if (nextResult.missingHeaders.length > 0) {
        setBlockingError(
          `Missing required columns: ${nextResult.missingHeaders.join(", ")}`,
        );
        return;
      }

      if (nextResult.labels.length === 0) {
        setBlockingError("No valid labels were found in this CSV.");
      }
    } catch (error) {
      setBlockingError(
        error instanceof Error ? error.message : "Unable to read this CSV file.",
      );
    } finally {
      setIsParsing(false);
    }
  }

  async function handleGeneratePdf() {
    if (!result || result.labels.length === 0) {
      return;
    }

    setIsGenerating(true);
    setBlockingError((currentError) =>
      currentError === "Unable to generate the PDF preview."
        ? null
        : currentError,
    );

    try {
      const pdfBytes = await generateLabelPdf(result.labels);
      const blob = new Blob([pdfBytes], { type: "application/pdf" });
      const nextPdfUrl = URL.createObjectURL(blob);

      setPdfUrl((currentUrl) => {
        if (currentUrl) {
          URL.revokeObjectURL(currentUrl);
        }

        return nextPdfUrl;
      });
    } catch {
      setBlockingError("Unable to generate the PDF preview.");
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <div className="space-y-8">
      <section className="rounded-2xl border border-black/10 bg-white p-6 shadow-sm">
        <div className="space-y-2">
          <h2 className="text-lg font-semibold text-slate-900">
            Upload customer CSV
          </h2>
          <p className="text-sm text-slate-600">
            This runs entirely in the browser. Your CSV stays on the local
            machine and the PDF is generated client-side.
          </p>
        </div>

        <div className="mt-5 space-y-4">
          <label
            className="flex cursor-pointer flex-col gap-3 rounded-xl border border-dashed border-slate-300 bg-slate-50 p-4 text-sm text-slate-700"
            htmlFor="csv-upload"
          >
            <span className="font-medium text-slate-900">
              Select a `.csv` file
            </span>
            <span>
              Required columns: <code>{REQUIRED_CSV_HEADERS.join(", ")}</code>
            </span>
            <input
              id="csv-upload"
              type="file"
              accept=".csv,text/csv"
              className="block text-sm text-slate-700 file:mr-4 file:rounded-lg file:border-0 file:bg-slate-900 file:px-4 file:py-2 file:text-sm file:font-medium file:text-white"
              onChange={handleFileChange}
            />
          </label>

          <div
            className="min-h-6 text-sm text-slate-600"
            aria-live="polite"
            data-testid="upload-status"
          >
            {isParsing
              ? "Parsing CSV and preparing label preview..."
              : selectedFileName
                ? `Loaded file: ${selectedFileName}`
                : "No file selected yet."}
          </div>

          {blockingError ? (
            <div className="rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
              {blockingError}
            </div>
          ) : null}

          {parserWarnings.length > 0 ? (
            <div className="rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-800">
              <p className="font-medium">CSV parser warnings</p>
              <ul className="mt-2 space-y-1">
                {parserWarnings.slice(0, 3).map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            </div>
          ) : null}
        </div>
      </section>

      {result ? (
        <section className="rounded-2xl border border-black/10 bg-white p-6 shadow-sm">
          <h2 className="text-lg font-semibold text-slate-900">Job summary</h2>

          <dl className="mt-5 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
            <StatCard
              label="Valid labels"
              value={String(result.labels.length)}
              testId="valid-label-count"
            />
            <StatCard
              label="PDF pages"
              value={String(result.pages.length || 0)}
              testId="page-count"
            />
            <StatCard
              label="Skipped rows"
              value={String(result.skippedRows.length)}
            />
            <StatCard
              label="Overflow warnings"
              value={String(result.warningCount)}
            />
          </dl>

          <div className="mt-4 text-sm text-slate-600">
            <p>Total CSV rows: {result.totalRowCount}</p>
            <p>Ignored empty rows: {result.emptyRowCount}</p>
          </div>

          {result.skippedRows.length > 0 ? (
            <div className="mt-5 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900">
              <p className="font-medium">Skipped rows</p>
              <ul className="mt-2 space-y-1">
                {result.skippedRows.slice(0, 5).map((row) => (
                  <li key={`${row.rowNumber}-${row.reason}`}>
                    Row {row.rowNumber}: {row.reason}
                  </li>
                ))}
              </ul>
              {result.skippedRows.length > 5 ? (
                <p className="mt-2">
                  {result.skippedRows.length - 5} more rows were skipped.
                </p>
              ) : null}
            </div>
          ) : null}
        </section>
      ) : null}

      <section className="rounded-2xl border border-black/10 bg-white p-6 shadow-sm">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <h2 className="text-lg font-semibold text-slate-900">
              Review data
            </h2>
            <p className="text-sm text-slate-600">
              Check the parsed recipient details before generating the PDF. The
              final file will place them on an A4{" "}
              {LABEL_LAYOUT_CONFIG.columns} x {LABEL_LAYOUT_CONFIG.rows} label
              sheet.
            </p>
          </div>

          <button
            type="button"
            className="inline-flex items-center justify-center gap-2 rounded-lg bg-slate-900 px-4 py-2 text-sm font-medium text-white transition hover:bg-slate-700 disabled:cursor-not-allowed disabled:bg-slate-300"
            onClick={handleGeneratePdf}
            disabled={!hasValidLabels || isGenerating || isParsing}
          >
            {isGenerating ? (
              <>
                <span
                  className="h-4 w-4 animate-spin rounded-full border-2 border-white/40 border-t-white"
                  aria-hidden="true"
                />
                Generating PDF...
              </>
            ) : (
              "Generate PDF"
            )}
          </button>
        </div>

        {isGenerating ? (
          <div
            className="mt-5 rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-800"
            aria-live="polite"
          >
            Building the PDF preview now. This can take a moment for larger CSV
            files.
          </div>
        ) : null}

        {result?.labels.length ? (
          <div
            className="mt-5 overflow-hidden rounded-xl border border-slate-200"
            data-testid="review-data-table"
          >
            <div className="max-h-[480px] overflow-auto">
              <table className="min-w-full divide-y divide-slate-200 text-left text-sm">
                <thead className="sticky top-0 bg-slate-50 text-slate-600">
                  <tr>
                    <th className="px-4 py-3 font-medium">#</th>
                    <th className="px-4 py-3 font-medium">Recipient</th>
                    <th className="px-4 py-3 font-medium">Address</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 bg-white text-slate-900">
                  {result.labels.map((label, index) => (
                    <tr key={`${label.name}-${index}`} className="align-top">
                      <td className="px-4 py-3 text-slate-500">{index + 1}</td>
                      <td className="px-4 py-3 font-medium">{label.name || "No name"}</td>
                      <td className="px-4 py-3 text-slate-700">
                        <p className="break-words">
                          {label.displayLines.join(", ")}
                        </p>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ) : (
          <div className="mt-5 rounded-xl border border-dashed border-slate-300 bg-slate-50 px-4 py-10 text-center text-sm text-slate-600">
            Upload a CSV to review the parsed recipients before generating the
            PDF.
          </div>
        )}
      </section>

      {pdfUrl ? (
        <section
          ref={pdfSectionRef}
          className="rounded-2xl border border-black/10 bg-white p-6 shadow-sm"
        >
          <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
            <div>
              <h2 className="text-lg font-semibold text-slate-900">
                Generated PDF
              </h2>
              <p className="text-sm text-slate-600">
                Review the finished sheet below, then download or open it in a
                new tab for printing.
              </p>
            </div>

            <div className="flex flex-wrap gap-3">
              <a
                className="inline-flex items-center justify-center rounded-lg border border-slate-300 px-4 py-2 text-sm font-medium text-slate-800 transition hover:bg-slate-50"
                href={pdfUrl}
                download={pdfFileName}
              >
                Download PDF
              </a>
              <a
                className="inline-flex items-center justify-center rounded-lg border border-slate-300 px-4 py-2 text-sm font-medium text-slate-800 transition hover:bg-slate-50"
                href={pdfUrl}
                target="_blank"
                rel="noreferrer"
              >
                Open in new tab
              </a>
            </div>
          </div>

          <iframe
            title="Generated labels PDF preview"
            src={pdfUrl}
            className="mt-5 h-[720px] w-full rounded-xl border border-slate-200"
            data-testid="pdf-preview"
          />
        </section>
      ) : null}
    </div>
  );
}

function StatCard({
  label,
  value,
  testId,
}: {
  label: string;
  value: string;
  testId?: string;
}) {
  return (
    <div className="rounded-xl border border-slate-200 bg-slate-50 p-4">
      <dt className="text-sm text-slate-500">{label}</dt>
      <dd className="mt-2 text-2xl font-semibold text-slate-900" data-testid={testId}>
        {value}
      </dd>
    </div>
  );
}

function buildPdfFileName(fileName: string) {
  const baseName = fileName.replace(/\.[^.]+$/, "").trim() || "hgm-address";

  return `${baseName.replace(/[^a-z0-9]+/gi, "-").replace(/^-+|-+$/g, "").toLowerCase()}-labels.pdf`;
}

async function readFileAsText(file: File) {
  if (typeof file.text === "function") {
    return file.text();
  }

  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      resolve(typeof reader.result === "string" ? reader.result : "");
    };
    reader.onerror = () => {
      reject(new Error("Unable to read this CSV file."));
    };

    reader.readAsText(file);
  });
}
