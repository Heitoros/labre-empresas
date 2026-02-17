import "dotenv/config";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { ConvexHttpClient } from "convex/browser";
import { api } from "../convex/_generated/api";

function getArg(flag: string): string | undefined {
  const idx = process.argv.indexOf(flag);
  if (idx === -1) return undefined;
  return process.argv[idx + 1];
}

function requiredNumberArg(flag: string): number {
  const value = getArg(flag);
  if (!value) throw new Error(`Parametro obrigatorio ausente: ${flag}`);
  const n = Number(value);
  if (!Number.isInteger(n)) throw new Error(`Parametro invalido em ${flag}: ${value}`);
  return n;
}

function buildCompetencia(ano: number, mes: number): string {
  return `${ano}-${String(mes).padStart(2, "0")}`;
}

type ResumoRegiao = {
  regiao: number;
  totalTrechos: number;
  totalKm: number;
  programadosNoMes: number;
  naoProgramadosNoMes: number;
  totalImportacoes: number;
  importacoesComErro: number;
  totalErros: number;
  porTipoFonte: Array<{ tipoFonte: string; totalTrechos: number; totalKm: number }>;
  topSrePorKm: Array<{ sre: string; km: number }>;
};

async function main() {
  const ano = requiredNumberArg("--ano");
  const mes = requiredNumberArg("--mes");
  const outputArg = getArg("--out");
  const regioesArg = getArg("--regioes");
  const regioes = (regioesArg ?? "1,2,3,11,12,13")
    .split(",")
    .map((r) => Number(r.trim()))
    .filter((r) => Number.isInteger(r));

  if (regioes.length === 0) {
    throw new Error("Nenhuma regiao valida informada em --regioes.");
  }

  const convexUrl = process.env.CONVEX_URL;
  if (!convexUrl) throw new Error("Defina CONVEX_URL no ambiente (.env ou shell).");

  const client = new ConvexHttpClient(convexUrl);
  const competencia = buildCompetencia(ano, mes);
  const outFile = outputArg ?? path.join(process.cwd(), "relatorios", `relatorio_consolidado_${competencia}.md`);

  const resultados: ResumoRegiao[] = [];

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
      topSrePorKm: graficos.series.topSrePorKm,
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

  await fs.mkdir(path.dirname(outFile), { recursive: true });

  const md = [
    `# Relatorio Tecnico Consolidado`,
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
      "- Top SRE por KM:",
      ...item.topSrePorKm.slice(0, 5).map((s) => `  - ${s.sre}: ${s.km.toFixed(2)} km`),
      "",
      `- Relatorio detalhado: relatorio_regiao_${item.regiao}_${competencia}.md`,
      "",
    ]),
  ].join("\n");

  await fs.writeFile(outFile, md, "utf8");
  console.log(`Relatorio consolidado gerado em: ${outFile}`);
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao gerar relatorio consolidado:", message);
  process.exit(1);
});
