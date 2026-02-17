import "dotenv/config";
import fs from "node:fs";
import fsp from "node:fs/promises";
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

function requiredNumberArg(flag: string): number {
  const value = requiredArg(flag);
  const n = Number(value);
  if (!Number.isInteger(n)) throw new Error(`Parametro invalido em ${flag}: ${value}`);
  return n;
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

function detectarTipoFontePorArquivo(arquivo: string): TipoFonte | undefined {
  const n = normalize(path.basename(arquivo));
  if (n.includes("nao paviment") || n.includes("n paviment")) return "NAO_PAV";
  if (n.includes("paviment")) return "PAV";
  return undefined;
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

function lerLinhasTrechos(arquivo: string, sheetName = "Trechos") {
  const workbook = XLSX.readFile(arquivo);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`Aba \"${sheetName}\" nao encontrada em: ${arquivo}`);

  const matriz = XLSX.utils.sheet_to_json<Array<unknown>>(sheet, {
    header: 1,
    raw: false,
    defval: "",
  });

  const headerIdx = matriz.findIndex((row) => {
    const tokens = row.map((cell) => canonicalHeader(String(cell ?? "")));
    return tokens.some((cell) => cell === "trecho" || cell === "trechos") && tokens.includes("sre");
  });

  if (headerIdx === -1) {
    throw new Error(`Cabecalho invalido na aba Trechos: ${arquivo}`);
  }

  const header = matriz[headerIdx].map((c) => String(c ?? "").trim());
  const mapped = header.map((h, i) => mapHeaderKey(h, i));
  const required = new Set(["TRECHO", "SRE", "EXT_KM"]);
  for (const key of required) {
    if (!mapped.includes(key)) {
      throw new Error(`Coluna obrigatoria ausente (${key}) em: ${arquivo}`);
    }
  }

  const linhas = matriz
    .slice(headerIdx + 1)
    .filter((row) => row.some((c) => String(c ?? "").trim() !== ""))
    .map((row) => {
      const obj: Record<string, unknown> = {};
      for (let i = 0; i < header.length; i += 1) {
        obj[mapHeaderKey(header[i] || "", i)] = row[i] ?? "";
      }
      return obj;
    });

  if (!linhas.length) throw new Error(`Sem linhas validas em: ${arquivo}`);
  return linhas;
}

function buildCompetencia(ano: number, mes: number): string {
  return `${ano}-${String(mes).padStart(2, "0")}`;
}

async function gerarRelatorioRegional(client: ConvexHttpClient, regiao: number, ano: number, mes: number) {
  const payload = await client.query(api.trechos.gerarPayloadRelatorio, { regiao, ano, mes });
  const graficos = await client.query(api.trechos.obterGraficosCompetencia, { regiao, ano, mes });
  const inconsistencias = await client.query(api.trechos.obterInconsistenciasImportacao, { regiao, ano, mes });

  const competencia = buildCompetencia(ano, mes);
  const outFile = path.join(process.cwd(), "relatorios", `relatorio_regiao_${regiao}_${competencia}.md`);
  await fsp.mkdir(path.dirname(outFile), { recursive: true });

  const md = [
    `# Relatorio Tecnico - Regiao ${regiao}`,
    "",
    `- Competencia: ${competencia}`,
    `- Gerado em: ${new Date(payload.metadata.geradoEm).toISOString()}`,
    "",
    "## Indicadores",
    `- Total de trechos: ${graficos.kpis.totalTrechos}`,
    `- Total de extensao (km): ${graficos.kpis.totalKm.toFixed(2)}`,
    `- Programados no mes: ${graficos.kpis.programadosNoMes}`,
    `- Nao programados no mes: ${graficos.kpis.naoProgramadosNoMes}`,
    `- Percentual programados: ${payload.graficos.kpis.percentualProgramados}%`,
    "",
    "## Distribuicao por Tipo de Fonte",
    ...graficos.series.porTipoFonte.map(
      (item) => `- ${item.tipoFonte}: ${item.totalTrechos} trechos / ${item.totalKm.toFixed(2)} km`,
    ),
    "",
    "## Top SRE por KM",
    ...graficos.series.topSrePorKm.map((item) => `- ${item.sre}: ${item.km.toFixed(2)} km`),
    "",
    "## Inconsistencias de Importacao",
    `- Total de importacoes: ${inconsistencias.resumo.totalImportacoes}`,
    `- Importacoes com erro: ${inconsistencias.resumo.importacoesComErro}`,
    `- Total de erros: ${inconsistencias.resumo.totalErros}`,
    ...inconsistencias.porCodigo.map((item) => `- ${item.codigo}: ${item.total}`),
    "",
    "## Observacoes Automaticas",
    ...payload.observacoesAutomaticas.map((obs) => `- ${obs}`),
    "",
  ].join("\n");

  await fsp.writeFile(outFile, md, "utf8");
  return outFile;
}

async function gerarConsolidado(client: ConvexHttpClient, ano: number, mes: number, regioes: number[]) {
  const competencia = buildCompetencia(ano, mes);
  const outFile = path.join(process.cwd(), "relatorios", `relatorio_consolidado_${competencia}.md`);

  const resultados = [] as Array<{
    regiao: number;
    totalTrechos: number;
    totalKm: number;
    programadosNoMes: number;
    naoProgramadosNoMes: number;
    totalImportacoes: number;
    importacoesComErro: number;
    totalErros: number;
    porTipoFonte: Array<{ tipoFonte: string; totalTrechos: number; totalKm: number }>;
  }>;

  for (const regiao of regioes) {
    const [graficos, inconsistencias] = await Promise.all([
      client.query(api.trechos.obterGraficosCompetencia, { regiao, ano, mes }),
      client.query(api.trechos.obterInconsistenciasImportacao, { regiao, ano, mes }),
    ]);
    resultados.push({
      regiao,
      totalTrechos: graficos.kpis.totalTrechos,
      totalKm: graficos.kpis.totalKm,
      programadosNoMes: graficos.kpis.programadosNoMes,
      naoProgramadosNoMes: graficos.kpis.naoProgramadosNoMes,
      totalImportacoes: inconsistencias.resumo.totalImportacoes,
      importacoesComErro: inconsistencias.resumo.importacoesComErro,
      totalErros: inconsistencias.resumo.totalErros,
      porTipoFonte: graficos.series.porTipoFonte,
    });
  }

  const totais = resultados.reduce(
    (acc, item) => {
      acc.totalTrechos += item.totalTrechos;
      acc.totalKm += item.totalKm;
      acc.programadosNoMes += item.programadosNoMes;
      acc.naoProgramadosNoMes += item.naoProgramadosNoMes;
      acc.totalImportacoes += item.totalImportacoes;
      acc.importacoesComErro += item.importacoesComErro;
      acc.totalErros += item.totalErros;
      return acc;
    },
    {
      totalTrechos: 0,
      totalKm: 0,
      programadosNoMes: 0,
      naoProgramadosNoMes: 0,
      totalImportacoes: 0,
      importacoesComErro: 0,
      totalErros: 0,
    },
  );

  const md = [
    "# Relatorio Tecnico Consolidado",
    "",
    `- Competencia: ${competencia}`,
    `- Regioes: ${regioes.join(", ")}`,
    `- Gerado em: ${new Date().toISOString()}`,
    "",
    "## Totais Consolidados",
    `- Total de trechos: ${totais.totalTrechos}`,
    `- Total de extensao (km): ${totais.totalKm.toFixed(2)}`,
    `- Programados no mes: ${totais.programadosNoMes}`,
    `- Nao programados no mes: ${totais.naoProgramadosNoMes}`,
    `- Total de importacoes: ${totais.totalImportacoes}`,
    `- Importacoes com erro: ${totais.importacoesComErro}`,
    `- Total de erros: ${totais.totalErros}`,
    "",
    "## Resumo por Regiao",
    ...resultados.flatMap((item) => [
      `### Regiao ${item.regiao}`,
      `- Total de trechos: ${item.totalTrechos}`,
      `- Total de extensao (km): ${item.totalKm.toFixed(2)}`,
      `- Programados no mes: ${item.programadosNoMes}`,
      `- Nao programados no mes: ${item.naoProgramadosNoMes}`,
      `- Total de importacoes: ${item.totalImportacoes}`,
      `- Importacoes com erro: ${item.importacoesComErro}`,
      `- Total de erros: ${item.totalErros}`,
      "",
      "- Distribuicao por tipo de fonte:",
      ...item.porTipoFonte.map((f) => `  - ${f.tipoFonte}: ${f.totalTrechos} trechos / ${f.totalKm.toFixed(2)} km`),
      "",
      `- Relatorio detalhado: relatorio_regiao_${item.regiao}_${competencia}.md`,
      "",
    ]),
  ].join("\n");

  await fsp.mkdir(path.dirname(outFile), { recursive: true });
  await fsp.writeFile(outFile, md, "utf8");
  return outFile;
}

async function main() {
  const pasta = requiredArg("--pasta");
  const ano = requiredNumberArg("--ano");
  const mes = requiredNumberArg("--mes");
  const regioes = (getArg("--regioes") ?? "1,2,3,11,12,13")
    .split(",")
    .map((r) => Number(r.trim()))
    .filter((r) => Number.isInteger(r));
  const dryRun = getArg("--dryRun") === "true";
  const incluirComplementar = getArg("--complementar") !== "false";
  const operador = getArg("--operador");
  const perfil = (getArg("--perfil") as "OPERADOR" | "GESTOR" | "ADMIN" | undefined) ?? "OPERADOR";
  const email = getArg("--email");
  const senha = getArg("--senha");

  const convexUrl = process.env.CONVEX_URL;
  if (!convexUrl) throw new Error("Defina CONVEX_URL no ambiente (.env ou shell).");

  const arquivos = listXlsxRecursively(pasta);
  const porRegiao = new Map<number, Partial<Record<TipoFonte, string>>>();

  for (const arquivo of arquivos) {
    const regiao = detectarRegiaoPorArquivo(arquivo);
    const tipo = detectarTipoFontePorArquivo(arquivo);
    if (!regiao || !tipo) continue;

    lerLinhasTrechos(arquivo, "Trechos");
    const atual = porRegiao.get(regiao) ?? {};
    atual[tipo] = arquivo;
    porRegiao.set(regiao, atual);
  }

  for (const regiao of regioes) {
    const item = porRegiao.get(regiao);
    if (!item?.PAV || !item?.NAO_PAV) {
      throw new Error(`Arquivos PAV/NAO_PAV nao encontrados para regiao ${regiao}.`);
    }
  }

  const client = new ConvexHttpClient(convexUrl);
  let sessionToken: string | undefined;
  if (incluirComplementar && !dryRun) {
    if (!email || !senha) {
      throw new Error("Para --complementar true, informe --email e --senha de um usuario autenticado.");
    }
    const sessao = await client.mutation(api.auth.login, { email, senha });
    sessionToken = String(sessao.token);
  }
  const importacoes: Array<Record<string, unknown>> = [];

  for (const regiao of regioes) {
    const item = porRegiao.get(regiao)!;

    for (const tipoFonte of ["PAV", "NAO_PAV"] as const) {
      const arquivo = item[tipoFonte]!;
      const linhas = lerLinhasTrechos(arquivo, "Trechos");

      const result = await client.mutation(api.trechos.importarTrechos, {
        tipoFonte,
        regiao,
        ano,
        mes,
        arquivoOrigem: path.basename(arquivo),
        linhas,
        limparAntes: true,
        dryRun,
        operador,
        perfil,
      });

      importacoes.push(result as Record<string, unknown>);

      if (!dryRun && incluirComplementar) {
        const arquivoBase64 = fs.readFileSync(arquivo).toString("base64");
        await client.action(api.workbook.importarWorkbookComplementar, {
          sessionToken: sessionToken!,
          regiao,
          ano,
          mes,
          tipoFonte,
          arquivoOrigem: path.basename(arquivo),
          arquivoBase64,
          limparAntes: true,
          operador,
          perfil,
        });
      }
    }
  }

  const relatoriosRegionais: string[] = [];
  if (!dryRun) {
    for (const regiao of regioes) {
      const file = await gerarRelatorioRegional(client, regiao, ano, mes);
      relatoriosRegionais.push(file);
    }
  }

  const consolidado = dryRun ? null : await gerarConsolidado(client, ano, mes, regioes);

  console.log(
    JSON.stringify(
      {
        competencia: buildCompetencia(ano, mes),
        pasta,
        dryRun,
        incluirComplementar,
        operador,
        perfil,
        regioes,
        importacoes,
        relatoriosRegionais,
        consolidado,
      },
      null,
      2,
    ),
  );
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro no fechamento mensal:", message);
  process.exit(1);
});
