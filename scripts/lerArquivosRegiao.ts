import path from "node:path";
import process from "node:process";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

function getArg(flag: string): string | undefined {
  const idx = process.argv.indexOf(flag);
  if (idx === -1) return undefined;
  return process.argv[idx + 1];
}

function requiredArg(flag: string): string {
  const value = getArg(flag);
  if (!value) throw new Error(`Parametro obrigatorio ausente: ${flag}`);
  return value;
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function readXlsx(filePath: string) {
  const workbook = XLSX.readFile(filePath);
  const output: Array<{
    sheet: string;
    headers: string[];
    sampleRows: Array<Record<string, string>>;
    totalRows: number;
  }> = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<Array<unknown>>(sheet, {
      header: 1,
      raw: false,
      defval: "",
    });

    const headers = (rows[0] ?? []).map((cell) => normalizeCell(cell));
    const dataRows = rows.slice(1).filter((row) => row.some((cell) => normalizeCell(cell) !== ""));

    const sampleRows = dataRows.slice(0, 5).map((row) => {
      const obj: Record<string, string> = {};
      headers.forEach((h, i) => {
        const key = h || `coluna_${i + 1}`;
        obj[key] = normalizeCell(row[i]);
      });
      return obj;
    });

    output.push({
      sheet: sheetName,
      headers,
      sampleRows,
      totalRows: dataRows.length,
    });
  }

  return output;
}

async function readDocx(filePath: string) {
  const result = await mammoth.extractRawText({ path: filePath });
  const text = result.value.replace(/\s+/g, " ").trim();
  return {
    chars: text.length,
    preview: text.slice(0, 1200),
  };
}

async function main() {
  const folder = requiredArg("--pasta");

  const docxPath = path.join(
    folder,
    "Relatório Técnico - Produto 02 - Região de Conservação 01 DEZ..docx",
  );
  const naoPavPath = path.join(
    folder,
    "ficha de inspeção_rodovias não pavimentadas_R.01 - DEZEMBRO.xlsx",
  );
  const pavPath = path.join(
    folder,
    "ficha de inspeção_rodovias pavimentadas_R.01 - DEZEMBRO.xlsx",
  );

  const [docx, naoPav, pav] = await Promise.all([
    readDocx(docxPath),
    Promise.resolve(readXlsx(naoPavPath)),
    Promise.resolve(readXlsx(pavPath)),
  ]);

  console.log(
    JSON.stringify(
      {
        pasta: folder,
        arquivos: {
          docx: {
            file: docxPath,
            ...docx,
          },
          naoPavimentadas: {
            file: naoPavPath,
            sheets: naoPav,
          },
          pavimentadas: {
            file: pavPath,
            sheets: pav,
          },
        },
      },
      null,
      2,
    ),
  );
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao ler arquivos:", message);
  process.exit(1);
});
