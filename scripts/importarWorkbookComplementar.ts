import "dotenv/config";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { ConvexHttpClient } from "convex/browser";
import { api } from "../convex/_generated/api";

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
  const n = Number(requiredArg(flag));
  if (!Number.isInteger(n)) throw new Error(`Parametro invalido em ${flag}`);
  return n;
}

async function main() {
  const arquivo = requiredArg("--arquivo");
  const regiao = requiredNumberArg("--regiao");
  const ano = requiredNumberArg("--ano");
  const mes = requiredNumberArg("--mes");
  const tipoFonte = requiredArg("--tipoFonte") as "PAV" | "NAO_PAV";
  const limparAntes = getArg("--limparAntes") !== "false";
  const operador = getArg("--operador");
  const perfil = (getArg("--perfil") as "OPERADOR" | "GESTOR" | "ADMIN" | undefined) ?? "OPERADOR";
  const email = requiredArg("--email");
  const senha = requiredArg("--senha");

  const convexUrl = process.env.CONVEX_URL;
  if (!convexUrl) throw new Error("Defina CONVEX_URL no ambiente.");

  const arquivoBase64 = fs.readFileSync(arquivo).toString("base64");
  const client = new ConvexHttpClient(convexUrl);

  const sessao = await client.mutation(api.auth.login, { email, senha });

  const result = await client.action(api.workbook.importarWorkbookComplementar, {
    sessionToken: String(sessao.token),
    regiao,
    ano,
    mes,
    tipoFonte,
    arquivoOrigem: path.basename(arquivo),
    arquivoBase64,
    limparAntes,
    operador,
    perfil,
  });

  console.log(JSON.stringify(result, null, 2));
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao importar workbook complementar:", message);
  process.exit(1);
});
