import { defineSchema, defineTable } from "convex/server";
import { v } from "convex/values";

export default defineSchema({
  usuarios: defineTable({
    nome: v.string(),
    email: v.string(),
    senhaHash: v.string(),
    perfil: v.union(v.literal("OPERADOR"), v.literal("GESTOR"), v.literal("ADMIN")),
    ativo: v.boolean(),
    forcarTrocaSenha: v.optional(v.boolean()),
    senhaAtualizadaEm: v.optional(v.number()),
    criadoEm: v.number(),
    atualizadoEm: v.number(),
  }).index("by_email", ["email"]),

  sessoes: defineTable({
    usuarioId: v.id("usuarios"),
    token: v.string(),
    nome: v.string(),
    email: v.string(),
    perfil: v.union(v.literal("OPERADOR"), v.literal("GESTOR"), v.literal("ADMIN")),
    criadoEm: v.number(),
    expiraEm: v.number(),
    revogadaEm: v.optional(v.number()),
  })
    .index("by_token", ["token"])
    .index("by_usuario", ["usuarioId"]),

  competencias: defineTable({
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    competencia: v.string(),
    criadoEm: v.number(),
    atualizadoEm: v.number(),
  })
    .index("by_competencia", ["competencia"])
    .index("by_regiao_ano_mes", ["regiao", "ano", "mes"]),

  importacoes: defineTable({
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    competencia: v.string(),

    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
    arquivoOrigem: v.string(),
    dryRun: v.boolean(),
    limparAntes: v.boolean(),

    status: v.union(
      v.literal("PROCESSANDO"),
      v.literal("SUCESSO"),
      v.literal("SUCESSO_COM_ERROS"),
      v.literal("ERRO"),
    ),

    totalLinhasRecebidas: v.number(),
    linhasValidas: v.number(),
    linhasIgnoradas: v.number(),
    linhasComErro: v.number(),
    gravados: v.number(),

    iniciadoEm: v.number(),
    finalizadoEm: v.optional(v.number()),
    erroFatal: v.optional(v.string()),
    operador: v.optional(v.string()),
    perfil: v.optional(v.union(v.literal("OPERADOR"), v.literal("GESTOR"), v.literal("ADMIN"))),
  })
    .index("by_competencia", ["competencia"])
    .index("by_competencia_tipo", ["competencia", "tipoFonte"])
    .index("by_regiao_ano_mes", ["regiao", "ano", "mes"])
    .index("by_regiao_ano_mes_tipo", ["regiao", "ano", "mes", "tipoFonte"]),

  ttAvaliacoes: defineTable({
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    competencia: v.string(),
    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
    numeroTrecho: v.optional(v.string()),
    trecho: v.string(),
    grupo: v.string(),
    classificacao: v.string(),
    valor: v.number(),
    importacaoId: v.optional(v.id("importacoes")),
    arquivoOrigem: v.string(),
    importadoEm: v.number(),
  })
    .index("by_competencia", ["competencia"])
    .index("by_regiao_ano_mes", ["regiao", "ano", "mes"])
    .index("by_regiao_ano_mes_tipo", ["regiao", "ano", "mes", "tipoFonte"])
    .index("by_trecho", ["trecho"]),

  workbookGraficos: defineTable({
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    competencia: v.string(),
    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
    aba: v.string(),
    ordem: v.number(),
    titulo: v.string(),
    tipoGrafico: v.string(),
    trecho: v.optional(v.string()),
    labels: v.array(v.string()),
    valores: v.array(v.number()),
    arquivoOrigem: v.string(),
    importadoEm: v.number(),
  })
    .index("by_competencia", ["competencia"])
    .index("by_regiao_ano_mes_tipo", ["regiao", "ano", "mes", "tipoFonte"])
    .index("by_aba", ["aba"]),

  auditoriaEventos: defineTable({
    acao: v.string(),
    entidade: v.string(),
    regiao: v.optional(v.number()),
    ano: v.optional(v.number()),
    mes: v.optional(v.number()),
    competencia: v.optional(v.string()),
    operador: v.optional(v.string()),
    perfil: v.optional(v.string()),
    detalhes: v.optional(v.string()),
    criadoEm: v.number(),
  })
    .index("by_criadoEm", ["criadoEm"])
    .index("by_competencia", ["competencia"]),

  importacaoErros: defineTable({
    importacaoId: v.id("importacoes"),
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    competencia: v.string(),

    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
    linhaPlanilha: v.number(),
    codigo: v.string(),
    mensagem: v.string(),
    coluna: v.optional(v.string()),
    valor: v.optional(v.string()),
    criadoEm: v.number(),
  })
    .index("by_importacao", ["importacaoId"])
    .index("by_competencia", ["competencia"]),

  trechos: defineTable({
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    competencia: v.string(),

    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),

    lote: v.optional(v.string()),
    numero: v.optional(v.number()),
    regiaoConservacao: v.optional(v.string()),
    cidadeSede: v.optional(v.string()),

    trecho: v.string(),
    sre: v.optional(v.string()),
    subtrechos: v.optional(v.string()),
    extKm: v.optional(v.number()),
    tipo: v.optional(v.string()),

    programacao: v.record(v.string(), v.boolean()),

    linhaPlanilha: v.number(),
    importacaoId: v.id("importacoes"),
    arquivoOrigem: v.string(),
    importadoEm: v.number(),
  })
    .index("by_competencia", ["competencia"])
    .index("by_competencia_tipo", ["competencia", "tipoFonte"])
    .index("by_regiao_ano_mes", ["regiao", "ano", "mes"])
    .index("by_regiao_ano_mes_tipo", ["regiao", "ano", "mes", "tipoFonte"])
    .index("by_tipo_sre", ["tipoFonte", "sre"])
    .index("by_sre", ["sre"])
    .index("by_trecho", ["trecho"]),
});
