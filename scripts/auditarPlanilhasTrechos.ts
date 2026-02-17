import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import * as XLSX from "xlsx";

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

function detectarRegiaoPorArquivo(arquivo: string): number | undefined {
  const m1 = arquivo.match(/regi[aã]o\s*0?([0-9]{1,2})/i);
  if (m1?.[1]) return Number(m1[1]);
  const m2 = arquivo.match(/r\.?\s*0?([0-9]{1,2})/i);
  if (m2?.[1]) return Number(m2[1]);
  return undefined;
}

function detectarTipoFontePorArquivo(arquivo: string): "PAV" | "NAO_PAV" | "DESCONHECIDO" {
  const n = normalize(path.basename(arquivo));
  if (n.includes("nao paviment") || n.includes("n paviment")) return "NAO_PAV";
  if (n.includes("paviment")) return "PAV";
  return "DESCONHECIDO";
}

function listXlsxRecursively(baseFolder: string): string[] {
  const output: string[] = [];

  const walk = (folder: string) => {
    for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
      const full = path.join(folder, entry.name);
      if (entry.isDirectory()) {
        walk(full);
        continue;
      }

      if (!entry.name.toLowerCase().endsWith(".xlsx")) continue;
      if (entry.name.startsWith("~$")) continue;
      output.push(full);
    }
  };

  walk(baseFolder);
  return output;
}

function auditarArquivo(filePath: string) {
  const workbook = XLSX.readFile(filePath);
  const trechosSheet = workbook.Sheets.Trechos;

  const regiao = detectarRegiaoPorArquivo(filePath);
  const tipoFonte = detectarTipoFontePorArquivo(filePath);
  const avisos: string[] = [];
  const erros: string[] = [];

  if (!trechosSheet) {
    erros.push('Aba "Trechos" nao encontrada.');
    return {
      arquivo: filePath,
      regiao,
      tipoFonte,
      status: "ERRO",
      totalLinhas: 0,
      cabecalhoOriginal: [] as string[],
      cabecalhoMapeado: [] as string[],
      avisos,
      erros,
    };
  }

  const matriz = XLSX.utils.sheet_to_json<Array<unknown>>(trechosSheet, {
    header: 1,
    raw: false,
    defval: "",
  });

  const headerIdx = matriz.findIndex((row) => {
    const tokens = row.map((cell) => canonicalHeader(String(cell ?? "")));
    return tokens.some((cell) => cell === "trecho" || cell === "trechos") && tokens.includes("sre");
  });

  if (headerIdx === -1) {
    erros.push("Nao foi possivel localizar cabecalho valido na aba Trechos.");
    return {
      arquivo: filePath,
      regiao,
      tipoFonte,
      status: "ERRO",
      totalLinhas: Math.max(0, matriz.length - 1),
      cabecalhoOriginal: [] as string[],
      cabecalhoMapeado: [] as string[],
      avisos,
      erros,
    };
  }

  const cabecalhoOriginal = matriz[headerIdx].map((cell) => String(cell ?? "").trim());
  const cabecalhoMapeado = cabecalhoOriginal.map((h, i) => mapHeaderKey(h, i));

  const possui = new Set(cabecalhoMapeado);
  for (const obrigatoria of ["TRECHO", "SRE", "EXT_KM"]) {
    if (!possui.has(obrigatoria)) {
      erros.push(`Coluna obrigatoria ausente: ${obrigatoria}.`);
    }
  }

  if (!possui.has("SUBTRECHOS")) {
    avisos.push("Coluna de subtrechos ausente ou com nome nao reconhecido.");
  }

  if (!regiao) avisos.push("Regiao nao detectada pelo nome do arquivo.");
  if (tipoFonte === "DESCONHECIDO") avisos.push("Tipo de fonte nao detectado pelo nome do arquivo.");

  const totalLinhas = matriz.slice(headerIdx + 1).filter((row) => row.some((c) => String(c ?? "").trim() !== "")).length;
  const status = erros.length > 0 ? "ERRO" : avisos.length > 0 ? "ALERTA" : "OK";

  return {
    arquivo: filePath,
    regiao,
    tipoFonte,
    status,
    totalLinhas,
    cabecalhoOriginal,
    cabecalhoMapeado,
    avisos,
    erros,
  };
}

function main() {
  const pasta = requiredArg("--pasta");
  const arquivos = listXlsxRecursively(pasta);
  const resultados = arquivos.map(auditarArquivo);

  const resumo = {
    totalArquivos: resultados.length,
    ok: resultados.filter((r) => r.status === "OK").length,
    alerta: resultados.filter((r) => r.status === "ALERTA").length,
    erro: resultados.filter((r) => r.status === "ERRO").length,
  };

  console.log(JSON.stringify({ pasta, resumo, resultados }, null, 2));
}

try {
  main();
} catch (err: unknown) {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro na auditoria:", message);
  process.exit(1);
}
