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

async function main() {
  const regiao = requiredNumberArg("--regiao");
  const ano = requiredNumberArg("--ano");
  const mes = requiredNumberArg("--mes");
  const outputArg = getArg("--out");

  const convexUrl = process.env.CONVEX_URL;
  if (!convexUrl) throw new Error("Defina CONVEX_URL no ambiente (.env ou shell).");

  const client = new ConvexHttpClient(convexUrl);
  const payload = await client.query(api.trechos.gerarPayloadRelatorio, { regiao, ano, mes });
  const graficos = await client.query(api.trechos.obterGraficosCompetencia, { regiao, ano, mes });
  const inconsistencias = await client.query(api.trechos.obterInconsistenciasImportacao, { regiao, ano, mes });

  const competencia = buildCompetencia(ano, mes);
  const outFile =
    outputArg ?? path.join(process.cwd(), "relatorios", `relatorio_regiao_${regiao}_${competencia}.md`);

  await fs.mkdir(path.dirname(outFile), { recursive: true });

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

  await fs.writeFile(outFile, md, "utf8");
  console.log(`Relatorio gerado em: ${outFile}`);
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao gerar relatorio:", message);
  process.exit(1);
});
