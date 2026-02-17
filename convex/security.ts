import type { MutationCtx, QueryCtx } from "./_generated/server";

type Ctx = MutationCtx | QueryCtx | any;

export async function requireSession(
  ctx: Ctx,
  sessionToken: string,
  allowedProfiles?: Array<"OPERADOR" | "GESTOR" | "ADMIN">,
  options?: { allowIfMustChangePassword?: boolean },
) {
  const sessao = await ctx.db.query("sessoes").withIndex("by_token", (q: any) => q.eq("token", sessionToken)).unique();
  if (!sessao || sessao.revogadaEm || sessao.expiraEm < Date.now()) {
    throw new Error("Sessao invalida ou expirada.");
  }

  if (allowedProfiles && !allowedProfiles.includes(sessao.perfil)) {
    throw new Error("Permissao insuficiente para esta operacao.");
  }

  const usuario = await ctx.db.get(sessao.usuarioId);
  if (!usuario || !usuario.ativo) {
    throw new Error("Usuario inativo ou inexistente.");
  }

  if (usuario.forcarTrocaSenha === true && options?.allowIfMustChangePassword !== true) {
    throw new Error("Troca de senha obrigatoria antes de continuar.");
  }

  return {
    usuarioId: sessao.usuarioId,
    nome: sessao.nome,
    email: sessao.email,
    perfil: sessao.perfil,
    token: sessao.token,
  };
}
