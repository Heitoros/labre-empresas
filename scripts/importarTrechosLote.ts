import "dotenv/config";
import path from "node:path";
import process from "node:process";
import * as XLSX from "xlsx";
import { ConvexHttpClient } from "convex/browser";
import { api } from "../convex/_generated/api";

type TipoFonte = "PAV" | "NAO_PAV";

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

function normalize(text: string): string {
  return text
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function canonicalHeader(text: string): string {
  return normalize(text).replace(/\s+/g, "");
}

function mapHeaderKey(header: string, idx: number): string {
  const n = canonicalHeader(header);
  if (n === "lote") return "LOTE";
  if (n === "n" || n === "numero" || n === "numerotrecho") return "NUMERO";
  if (n === "regiaodeconservacao") return "REGIAO_CONSERVACAO";
  if (n === "cidadesede") return "CIDADE_SEDE";
  if (n === "trecho" || n === "trechos") return "TRECHO";
  if (n === "sre") return "SRE";
  if (n === "subtrechos" || n === "subtrecho" || n === "sbutrecho" || n === "segmentos") return "SUBTRECHOS";
  if (n === "extkm" || n === "extensao" || n === "extensaokm") return "EXT_KM";
  if (n === "tipo") return "TIPO";
  if (n.startsWith("jul")) return "JUL";
  if (n.startsWith("aug") || n.startsWith("ago")) return "AGO";
  if (n.startsWith("sep") || n.startsWith("set")) return "SET";
  if (n.startsWith("oct") || n.startsWith("out")) return "OUT";
  if (n.startsWith("nov")) return "NOV";
  if (n.startsWith("dec") || n.startsWith("dez")) return "DEZ";
  return `COL_${idx + 1}`;
}

function detectarAno(arquivo: string): number | undefined {
  const matches = arquivo.match(/(20\d{2})/g);
  if (!matches || matches.length === 0) return undefined;
  return Number(matches[matches.length - 1]);
}

function detectarMes(arquivo: string): number | undefined {
  const matchNumerico = arquivo.match(/20\d{2}[-_ .]?(0?[1-9]|1[0-2])/);
  if (matchNumerico?.[1]) return Number(matchNumerico[1]);

  const n = normalize(arquivo);
  if (n.includes("janeiro")) return 1;
  if (n.includes("fevereiro")) return 2;
  if (n.includes("marco") || n.includes("marco")) return 3;
  if (n.includes("abril")) return 4;
  if (n.includes("maio")) return 5;
  if (n.includes("junho")) return 6;
  if (n.includes("julho")) return 7;
  if (n.includes("agosto")) return 8;
  if (n.includes("setembro")) return 9;
  if (n.includes("outubro")) return 10;
  if (n.includes("novembro")) return 11;
  if (n.includes("dezembro")) return 12;
  return undefined;
}

function detectarRegiao(arquivo: string): number | undefined {
  const m1 = arquivo.match(/regi[aã]o\s*0?([0-9]{1,2})/i);
  if (m1?.[1]) return Number(m1[1]);
  const m2 = arquivo.match(/r\.?\s*0?([0-9]{1,2})/i);
  if (m2?.[1]) return Number(m2[1]);
  return undefined;
}

function lerLinhasTrechos(arquivo: string, sheetName = "Trechos") {
  const workbook = XLSX.readFile(arquivo);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`A aba "${sheetName}" nao existe em "${arquivo}".`);

  const matriz = XLSX.utils.sheet_to_json<Array<unknown>>(sheet, {
    header: 1,
    raw: false,
    defval: "",
  });

  const headerIdx = matriz.findIndex((row) => {
    const asText = row.map((c) => canonicalHeader(String(c ?? "")));
    return asText.some((cell) => cell === "trecho" || cell === "trechos") && asText.includes("sre");
  });

  if (headerIdx === -1) {
    throw new Error(`Nao foi possivel localizar cabecalho da aba Trechos em "${arquivo}".`);
  }

  const header = matriz[headerIdx].map((c) => String(c ?? "").trim());
  const linhas = matriz
    .slice(headerIdx + 1)
    .filter((row) => row.some((c) => String(c ?? "").trim() !== ""))
    .map((row) => {
      const obj: Record<string, unknown> = {};
      for (let i = 0; i < header.length; i += 1) {
        const key = mapHeaderKey(header[i] || "", i);
        obj[key] = row[i] ?? "";
      }
      return obj;
    });

  if (!linhas.length) throw new Error(`Nenhuma linha valida encontrada em "${arquivo}".`);
  return linhas;
}

async function importarUm(
  client: ConvexHttpClient,
  tipoFonte: TipoFonte,
  arquivo: string,
  ano: number,
  mes: number,
  regiao: number,
  sheetName: string,
  limparAntes: boolean,
  dryRun: boolean,
) {
  const linhas = lerLinhasTrechos(arquivo, sheetName);
  return client.mutation(api.trechos.importarTrechos, {
    tipoFonte,
    regiao,
    ano,
    mes,
    arquivoOrigem: path.basename(arquivo),
    linhas,
    limparAntes,
    dryRun,
  });
}

async function main() {
  const arquivoPav = requiredArg("--pav");
  const arquivoNaoPav = requiredArg("--naoPav");

  const anoArg = getArg("--ano");
  const mesArg = getArg("--mes");
  const regiaoArg = getArg("--regiao");

  const ano = anoArg ? Number(anoArg) : detectarAno(arquivoPav) ?? detectarAno(arquivoNaoPav);
  const mes = mesArg ? Number(mesArg) : detectarMes(arquivoPav) ?? detectarMes(arquivoNaoPav);
  const regiao = regiaoArg
    ? Number(regiaoArg)
    : detectarRegiao(arquivoPav) ?? detectarRegiao(arquivoNaoPav);

  if (!ano || !Number.isInteger(ano)) {
    throw new Error("Nao foi possivel definir o ano. Informe --ano ou use arquivo com ano no nome.");
  }
  if (!mes || !Number.isInteger(mes)) {
    throw new Error("Nao foi possivel definir o mes. Informe --mes ou use arquivo com mes no nome.");
  }
  if (!regiao || !Number.isInteger(regiao)) {
    throw new Error("Nao foi possivel definir a regiao. Informe --regiao ou use arquivo com regiao no nome.");
  }

  const sheetName = getArg("--sheet") ?? "Trechos";
  const limparAntes = getArg("--limparAntes") !== "false";
  const dryRun = getArg("--dryRun") === "true";

  const convexUrl = process.env.CONVEX_URL;
  if (!convexUrl) throw new Error("Defina CONVEX_URL no ambiente (.env ou shell).");

  const client = new ConvexHttpClient(convexUrl);

  const pav = await importarUm(
    client,
    "PAV",
    arquivoPav,
    ano,
    mes,
    regiao,
    sheetName,
    limparAntes,
    dryRun,
  );

  const naoPav = await importarUm(
    client,
    "NAO_PAV",
    arquivoNaoPav,
    ano,
    mes,
    regiao,
    sheetName,
    limparAntes,
    dryRun,
  );

  console.log("Importacao em lote concluida:");
  console.log(JSON.stringify({ PAV: pav, NAO_PAV: naoPav }, null, 2));
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro na importacao em lote:", message);
  process.exit(1);
});
