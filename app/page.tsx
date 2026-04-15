import LabelGenerator from "@/app/components/label-generator";

export default function Home() {
  return (
    <main className="mx-auto flex min-h-screen w-full max-w-6xl flex-col px-6 py-10 sm:px-8">
      <div className="max-w-3xl space-y-4">
        <p className="text-sm font-medium uppercase tracking-[0.2em] text-slate-500">
          Hatton Garden Metals
        </p>
        <h1 className="text-4xl font-semibold tracking-tight text-slate-950 sm:text-5xl">
          Address label PDF generator
        </h1>
        <p className="text-base leading-7 text-slate-600 sm:text-lg">
          Upload a customer CSV, review the data, and generate a print-ready PDF
          pack labels.
        </p>
      </div>

      <div className="mt-10">
        <LabelGenerator />
      </div>
    </main>
  );
}
