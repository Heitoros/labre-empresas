import "dotenv/config";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { ConvexHttpClient } from "convex/browser";
import { Document, HeadingLevel, Packer, Paragraph } from "docx";
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

function bullet(text: string): Paragraph {
  return new Paragraph({ text, bullet: { level: 0 } });
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
    outputArg ?? path.join(process.cwd(), "relatorios", `relatorio_regiao_${regiao}_${competencia}.docx`);

  await fs.mkdir(path.dirname(outFile), { recursive: true });

  const children: Paragraph[] = [
    new Paragraph({ text: `Relatorio Tecnico - Regiao ${regiao}`, heading: HeadingLevel.TITLE }),
    bullet(`Competencia: ${competencia}`),
    bullet(`Gerado em: ${new Date(payload.metadata.geradoEm).toISOString()}`),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "Indicadores", heading: HeadingLevel.HEADING_1 }),
    bullet(`Total de trechos: ${graficos.kpis.totalTrechos}`),
    bullet(`Total de extensao (km): ${graficos.kpis.totalKm.toFixed(2)}`),
    bullet(`Programados no mes: ${graficos.kpis.programadosNoMes}`),
    bullet(`Nao programados no mes: ${graficos.kpis.naoProgramadosNoMes}`),
    bullet(`Percentual programados: ${payload.graficos.kpis.percentualProgramados}%`),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "Distribuicao por Tipo de Fonte", heading: HeadingLevel.HEADING_1 }),
    ...graficos.series.porTipoFonte.map((item) =>
      bullet(`${item.tipoFonte}: ${item.totalTrechos} trechos / ${item.totalKm.toFixed(2)} km`),
    ),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "Top SRE por KM", heading: HeadingLevel.HEADING_1 }),
    ...graficos.series.topSrePorKm.map((item) => bullet(`${item.sre}: ${item.km.toFixed(2)} km`)),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "Inconsistencias de Importacao", heading: HeadingLevel.HEADING_1 }),
    bullet(`Total de importacoes: ${inconsistencias.resumo.totalImportacoes}`),
    bullet(`Importacoes com erro: ${inconsistencias.resumo.importacoesComErro}`),
    bullet(`Total de erros: ${inconsistencias.resumo.totalErros}`),
    ...inconsistencias.porCodigo.map((item) => bullet(`${item.codigo}: ${item.total}`)),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "Observacoes Automaticas", heading: HeadingLevel.HEADING_1 }),
    ...payload.observacoesAutomaticas.map((obs) => bullet(obs)),
  ];

  const doc = new Document({
    sections: [{ children }],
  });

  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(outFile, buffer);
  console.log(`Relatorio DOCX gerado em: ${outFile}`);
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao gerar relatorio DOCX:", message);
  process.exit(1);
});
