import { CommonModule } from "@angular/common";
import { Component, OnInit } from "@angular/core";
import { FormsModule } from "@angular/forms";
import { ConvexHttpClient } from "convex/browser";
import {
  AlignmentType,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableOfContents,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import { api } from "../../convex/_generated/api";
import { environment } from "../environments/environment";

type FonteResumo = { tipoFonte: string; totalTrechos: number; totalKm: number };
type SreResumo = { sre: string; km: number };
type CodigoInconsistenciaResumo = { codigo: string; total: number };
type ImportacaoResumo = {
  id: string;
  tipoFonte: string;
  status: string;
  arquivoOrigem: string;
  iniciadoEm?: number;
  finalizadoEm?: number;
  linhasValidas: number;
  linhasIgnoradas: number;
  linhasComErro: number;
  gravados: number;
};
type AuditoriaEvento = {
  id: string;
  acao: string;
  entidade: string;
  competencia?: string;
  regiao?: number;
  operador?: string;
  perfil?: string;
  detalhes?: string;
  criadoEm: number;
};
type ImportacaoStatus = {
  id: string;
  status: string;
  gravados: number;
  linhasIgnoradas: number;
  linhasComErro: number;
  erroFatal?: string;
};
type SessaoAtual = {
  token: string;
  usuario: {
    id: string;
    nome: string;
    email: string;
    perfil: "OPERADOR" | "GESTOR" | "ADMIN";
  };
  expiraEm: number;
  precisaTrocaSenha?: boolean;
};
type UsuarioAdmin = {
  id: string;
  nome: string;
  email: string;
  perfil: "OPERADOR" | "GESTOR" | "ADMIN";
  ativo: boolean;
  forcarTrocaSenha: boolean;
  atualizadoEm: number;
};
type SaudeOperacional = {
  periodoHoras: number;
  totalImportacoes: number;
  processando: number;
  erros: number;
  sucessoComErros: number;
  taxaSucesso: number;
};
type WorkbookSerie = { label: string; valor: number; percentual: number };
type WorkbookGrafico = {
  id: string;
  aba: string;
  ordem: number;
  titulo: string;
  tipoGrafico: string;
  total: number;
  series: WorkbookSerie[];
};

@Component({
  selector: "app-root",
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: "./app.component.html",
  styleUrl: "./app.component.css",
})
export class AppComponent implements OnInit {
  regiao = 1;
  ano = 2025;
  mes = 12;
  loading = false;
  erro = "";

  totalTrechos = 0;
  totalKm = 0;
  programados = 0;
  naoProgramados = 0;

  porTipoFonte: FonteResumo[] = [];
  topSrePorKm: SreResumo[] = [];
  porCodigoInconsistencia: CodigoInconsistenciaResumo[] = [];
  importacoes: ImportacaoResumo[] = [];
  mostrarHistoricoCompleto = false;
  gerandoConsolidado = false;
  uploadTipoFonte: "PAV" | "NAO_PAV" = "PAV";
  uploadLimparAntes = true;
  uploadProcessarComplementar = true;
  operador = "";
  perfil: "OPERADOR" | "GESTOR" | "ADMIN" = "OPERADOR";
  emailLogin = "";
  senhaLogin = "";
  authMensagem = "";
  sessaoAtual: SessaoAtual | null = null;
  senhaAtualTroca = "";
  novaSenhaTroca = "";
  confirmarSenhaTroca = "";
  arquivoSelecionado: File | null = null;
  arquivoPavSelecionado: File | null = null;
  arquivoNaoPavSelecionado: File | null = null;
  uploadEmAndamento = false;
  uploadImportacaoId = "";
  uploadMensagem = "";
  auditoriaRecente: AuditoriaEvento[] = [];
  usuariosAdmin: UsuarioAdmin[] = [];
  saudeOperacional: SaudeOperacional | null = null;
  filtroUsuarios = "";
  novoUsuarioNome = "";
  novoUsuarioEmail = "";
  novoUsuarioSenha = "Temp@123";
  novoUsuarioPerfil: "OPERADOR" | "GESTOR" | "ADMIN" = "OPERADOR";
  confirmacaoModalAberto = false;
  confirmacaoModalTitulo = "";
  confirmacaoModalMensagem = "";
  confirmacaoAcao: "RESET_SENHA" | "ALTERAR_ATIVO" | null = null;
  confirmacaoUsuario: UsuarioAdmin | null = null;
  confirmacaoNovoAtivo: boolean | null = null;
  private sessionExpiryTimer: ReturnType<typeof setTimeout> | null = null;
  resumoInconsistencias = {
    totalImportacoes: 0,
    importacoesComErro: 0,
    totalErros: 0,
  };
  abaGraficos: "tipoFonte" | "programacao" = "tipoFonte";
  workbookTipoFonte: "PAV" | "NAO_PAV" = "PAV";
  workbookAbaSelecionada = "";
  workbookGraficos: WorkbookGrafico[] = [];

  meses = [
    { value: 1, label: "Janeiro" },
    { value: 2, label: "Fevereiro" },
    { value: 3, label: "Marco" },
    { value: 4, label: "Abril" },
    { value: 5, label: "Maio" },
    { value: 6, label: "Junho" },
    { value: 7, label: "Julho" },
    { value: 8, label: "Agosto" },
    { value: 9, label: "Setembro" },
    { value: 10, label: "Outubro" },
    { value: 11, label: "Novembro" },
    { value: 12, label: "Dezembro" },
  ];

  regioes = [1, 2, 3, 11, 12, 13];

  private client = new ConvexHttpClient(environment.convexUrl);

  ngOnInit(): void {
    this.operador = localStorage.getItem("labre_operador") ?? "";
    this.perfil = (localStorage.getItem("labre_perfil") as "OPERADOR" | "GESTOR" | "ADMIN") ?? "OPERADOR";
    void this.inicializarSessao();
  }

  private tokenSessao(): string {
    if (!this.sessaoAtual?.token) throw new Error("Sessao nao autenticada.");
    return this.sessaoAtual.token;
  }

  private async inicializarSessao(): Promise<void> {
    const token = localStorage.getItem("labre_session_token");
    if (!token) return;

    try {
      const sessao = await this.client.query(api.auth.me, { sessionToken: token });
      if (!sessao) {
        localStorage.removeItem("labre_session_token");
        return;
      }
      this.sessaoAtual = sessao as SessaoAtual;
      this.agendarExpiracaoSessao();
      if (!this.operador) this.operador = this.sessaoAtual.usuario.nome;
      this.perfil = this.sessaoAtual.usuario.perfil;
      this.salvarPerfilOperacao();
      if (!this.sessaoAtual.precisaTrocaSenha) {
        await this.recarregar();
      }
    } catch {
      localStorage.removeItem("labre_session_token");
    }
  }

  private agendarExpiracaoSessao(): void {
    if (this.sessionExpiryTimer) {
      clearTimeout(this.sessionExpiryTimer);
      this.sessionExpiryTimer = null;
    }
    if (!this.sessaoAtual?.expiraEm) return;

    const ms = this.sessaoAtual.expiraEm - Date.now();
    if (ms <= 0) {
      void this.logoutPorExpiracao();
      return;
    }

    this.sessionExpiryTimer = setTimeout(() => {
      void this.logoutPorExpiracao();
    }, ms + 1000);
  }

  private async logoutPorExpiracao(): Promise<void> {
    await this.logout();
    this.authMensagem = "Sessao expirada. Faca login novamente.";
  }

  private erroAutenticacao(message: string): boolean {
    const m = message.toLowerCase();
    return m.includes("sessao") || m.includes("sessão") || m.includes("expirada") || m.includes("autenticada");
  }

  private async tratarErroAutenticacao(e: unknown): Promise<boolean> {
    const msg = e instanceof Error ? e.message : String(e);
    if (!this.erroAutenticacao(msg)) return false;
    await this.logout();
    this.authMensagem = "Sua sessao expirou ou ficou invalida. Entre novamente.";
    return true;
  }

  async login(): Promise<void> {
    this.authMensagem = "";
    try {
      const sessao = await this.client.mutation(api.auth.login, {
        email: this.emailLogin,
        senha: this.senhaLogin,
      });
      this.sessaoAtual = sessao as SessaoAtual;
      this.agendarExpiracaoSessao();
      localStorage.setItem("labre_session_token", this.sessaoAtual.token);
      this.operador = this.sessaoAtual.usuario.nome;
      this.perfil = this.sessaoAtual.usuario.perfil;
      this.salvarPerfilOperacao();
      this.authMensagem = "Login realizado com sucesso.";
      this.senhaLogin = "";
      if (!this.sessaoAtual.precisaTrocaSenha) {
        await this.recarregar();
      } else {
        this.authMensagem = "Troca de senha obrigatoria no primeiro acesso.";
      }
    } catch (e) {
      this.authMensagem = e instanceof Error ? e.message : String(e);
    }
  }

  async logout(): Promise<void> {
    try {
      if (this.sessaoAtual?.token) {
        await this.client.mutation(api.auth.logout, { sessionToken: this.sessaoAtual.token });
      }
    } finally {
      this.sessaoAtual = null;
      this.usuariosAdmin = [];
      localStorage.removeItem("labre_session_token");
      if (this.sessionExpiryTimer) {
        clearTimeout(this.sessionExpiryTimer);
        this.sessionExpiryTimer = null;
      }
    }
  }

  async criarUsuarioAdmin(): Promise<void> {
    if (!this.sessaoAtual?.token) return;
    try {
      await this.client.mutation(api.auth.criarUsuario, {
        sessionToken: this.sessaoAtual.token,
        nome: this.novoUsuarioNome,
        email: this.novoUsuarioEmail,
        senha: this.novoUsuarioSenha,
        perfil: this.novoUsuarioPerfil,
      });
      this.authMensagem = "Usuario criado com sucesso.";
      this.novoUsuarioNome = "";
      this.novoUsuarioEmail = "";
      this.novoUsuarioSenha = "Temp@123";
      this.novoUsuarioPerfil = "OPERADOR";
      await this.carregarUsuariosAdmin();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.authMensagem = e instanceof Error ? e.message : String(e);
    }
  }

  async atualizarUsuarioAdmin(user: UsuarioAdmin): Promise<void> {
    if (!this.sessaoAtual?.token) return;
    try {
      await this.client.mutation(api.auth.atualizarUsuario, {
        sessionToken: this.sessaoAtual.token,
        usuarioId: user.id as any,
        nome: user.nome,
        perfil: user.perfil,
        ativo: user.ativo,
      });
      await this.carregarUsuariosAdmin();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.authMensagem = e instanceof Error ? e.message : String(e);
    }
  }

  async redefinirSenhaUsuarioAdmin(user: UsuarioAdmin): Promise<void> {
    if (!this.sessaoAtual?.token) return;
    try {
      await this.client.mutation(api.auth.redefinirSenhaUsuario, {
        sessionToken: this.sessaoAtual.token,
        usuarioId: user.id as any,
        novaSenhaTemporaria: "Temp@123",
      });
      this.authMensagem = `Senha temporaria redefinida para ${user.email}.`;
      await this.carregarUsuariosAdmin();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.authMensagem = e instanceof Error ? e.message : String(e);
    }
  }

  private async carregarUsuariosAdmin(): Promise<void> {
    if (!this.sessaoAtual?.token || this.sessaoAtual.usuario.perfil !== "ADMIN") {
      this.usuariosAdmin = [];
      return;
    }
    const users = await this.client.query(api.auth.listarUsuarios, {
      sessionToken: this.sessaoAtual.token,
    });
    this.usuariosAdmin = users as UsuarioAdmin[];
  }

  usuariosFiltrados(): UsuarioAdmin[] {
    const filtro = this.filtroUsuarios.trim().toLowerCase();
    if (!filtro) return this.usuariosAdmin;
    return this.usuariosAdmin.filter(
      (u) => u.nome.toLowerCase().includes(filtro) || u.email.toLowerCase().includes(filtro) || u.perfil.toLowerCase().includes(filtro),
    );
  }

  abrirConfirmacaoReset(user: UsuarioAdmin): void {
    this.confirmacaoModalAberto = true;
    this.confirmacaoAcao = "RESET_SENHA";
    this.confirmacaoUsuario = user;
    this.confirmacaoNovoAtivo = null;
    this.confirmacaoModalTitulo = "Confirmar reset de senha";
    this.confirmacaoModalMensagem = `Redefinir senha de ${user.email} para a temporaria Temp@123 e forcar troca no proximo login?`;
  }

  abrirConfirmacaoAtivo(user: UsuarioAdmin, novoAtivo: boolean): void {
    this.confirmacaoModalAberto = true;
    this.confirmacaoAcao = "ALTERAR_ATIVO";
    this.confirmacaoUsuario = user;
    this.confirmacaoNovoAtivo = novoAtivo;
    this.confirmacaoModalTitulo = novoAtivo ? "Confirmar ativacao" : "Confirmar inativacao";
    this.confirmacaoModalMensagem = novoAtivo
      ? `Ativar usuario ${user.email}?`
      : `Inativar usuario ${user.email}? Ele perdera acesso imediato.`;
  }

  fecharConfirmacao(): void {
    this.confirmacaoModalAberto = false;
    this.confirmacaoAcao = null;
    this.confirmacaoUsuario = null;
    this.confirmacaoNovoAtivo = null;
  }

  async confirmarAcaoModal(): Promise<void> {
    if (!this.confirmacaoUsuario || !this.confirmacaoAcao) {
      this.fecharConfirmacao();
      return;
    }

    try {
      if (this.confirmacaoAcao === "RESET_SENHA") {
        await this.redefinirSenhaUsuarioAdmin(this.confirmacaoUsuario);
      }

      if (this.confirmacaoAcao === "ALTERAR_ATIVO" && this.confirmacaoNovoAtivo !== null) {
        await this.client.mutation(api.auth.atualizarUsuario, {
          sessionToken: this.tokenSessao(),
          usuarioId: this.confirmacaoUsuario.id as any,
          nome: this.confirmacaoUsuario.nome,
          perfil: this.confirmacaoUsuario.perfil,
          ativo: this.confirmacaoNovoAtivo,
        });
        await this.carregarUsuariosAdmin();
      }
    } catch (e) {
      if (!(await this.tratarErroAutenticacao(e))) {
        this.authMensagem = e instanceof Error ? e.message : String(e);
      }
    } finally {
      this.fecharConfirmacao();
    }
  }

  async alterarSenhaObrigatoria(): Promise<void> {
    this.authMensagem = "";
    if (!this.sessaoAtual?.token) return;
    if (this.novaSenhaTroca !== this.confirmarSenhaTroca) {
      this.authMensagem = "A confirmacao da nova senha nao confere.";
      return;
    }

    try {
      await this.client.mutation(api.auth.alterarSenha, {
        sessionToken: this.sessaoAtual.token,
        senhaAtual: this.senhaAtualTroca,
        novaSenha: this.novaSenhaTroca,
      });

      const sessao = await this.client.query(api.auth.me, { sessionToken: this.sessaoAtual.token });
      this.sessaoAtual = sessao as SessaoAtual;
      this.senhaAtualTroca = "";
      this.novaSenhaTroca = "";
      this.confirmarSenhaTroca = "";
      this.authMensagem = "Senha alterada com sucesso.";
      await this.recarregar();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.authMensagem = e instanceof Error ? e.message : String(e);
    }
  }

  salvarPerfilOperacao(): void {
    localStorage.setItem("labre_operador", this.operador);
    localStorage.setItem("labre_perfil", this.perfil);
  }

  private fileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = String(reader.result ?? "");
        const base64 = result.includes(",") ? result.split(",")[1] : "";
        if (!base64) {
          reject(new Error("Falha ao converter arquivo para base64."));
          return;
        }
        resolve(base64);
      };
      reader.onerror = () => reject(new Error("Falha ao ler arquivo selecionado."));
      reader.readAsDataURL(file);
    });
  }

  private validarTamanhoArquivo(file: File): boolean {
    const maxBytes = 25 * 1024 * 1024;
    if (file.size > maxBytes) {
      this.uploadMensagem = `Arquivo ${file.name} excede 25MB.`;
      return false;
    }
    return true;
  }

  onArquivoSelecionado(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.arquivoSelecionado = input.files?.[0] ?? null;
    if (this.arquivoSelecionado && !this.validarTamanhoArquivo(this.arquivoSelecionado)) {
      this.arquivoSelecionado = null;
      return;
    }
    this.uploadMensagem = this.arquivoSelecionado ? `Arquivo selecionado: ${this.arquivoSelecionado.name}` : "";
  }

  onArquivoPavSelecionado(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.arquivoPavSelecionado = input.files?.[0] ?? null;
    if (this.arquivoPavSelecionado && !this.validarTamanhoArquivo(this.arquivoPavSelecionado)) {
      this.arquivoPavSelecionado = null;
    }
  }

  onArquivoNaoPavSelecionado(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.arquivoNaoPavSelecionado = input.files?.[0] ?? null;
    if (this.arquivoNaoPavSelecionado && !this.validarTamanhoArquivo(this.arquivoNaoPavSelecionado)) {
      this.arquivoNaoPavSelecionado = null;
    }
  }

  async importarArquivoSelecionado(): Promise<void> {
    if (!this.arquivoSelecionado) {
      this.uploadMensagem = "Selecione um arquivo antes de importar.";
      return;
    }

    this.uploadEmAndamento = true;
    this.uploadMensagem = "Enviando arquivo e processando no servidor...";

    try {
      const arquivoBase64 = await this.fileToBase64(this.arquivoSelecionado);

      const resultado = await this.client.mutation(api.trechos.iniciarImportacaoArquivoAsync, {
        sessionToken: this.tokenSessao(),
        tipoFonte: this.uploadTipoFonte,
        regiao: this.regiao,
        ano: this.ano,
        mes: this.mes,
        arquivoOrigem: this.arquivoSelecionado.name,
        arquivoBase64,
        limparAntes: this.uploadLimparAntes,
        dryRun: false,
        operador: this.operador || undefined,
        perfil: this.perfil,
      });

      this.uploadImportacaoId = String(resultado.importacaoId ?? "");
      this.uploadMensagem = `Importacao enfileirada (ID: ${this.uploadImportacaoId}). Aguardando processamento...`;

      if (this.uploadProcessarComplementar) {
        await this.client.action(api.workbook.importarWorkbookComplementar, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: this.uploadTipoFonte,
          arquivoOrigem: this.arquivoSelecionado.name,
          arquivoBase64,
          limparAntes: true,
          operador: this.operador || undefined,
          perfil: this.perfil,
        });
      }

      await this.acompanharImportacoes([this.uploadImportacaoId]);
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.uploadMensagem = e instanceof Error ? `Falha na importacao: ${e.message}` : String(e);
    } finally {
      this.uploadEmAndamento = false;
    }
  }

  async importarLoteSelecionado(): Promise<void> {
    if (!this.arquivoPavSelecionado || !this.arquivoNaoPavSelecionado) {
      this.uploadMensagem = "Selecione os dois arquivos (PAV e NAO_PAV).";
      return;
    }

    this.uploadEmAndamento = true;
    this.uploadMensagem = "Enviando lote PAV + NAO_PAV para processamento...";

    try {
      const [pavArquivoBase64, naoPavArquivoBase64] = await Promise.all([
        this.fileToBase64(this.arquivoPavSelecionado),
        this.fileToBase64(this.arquivoNaoPavSelecionado),
      ]);

      const resultado = await this.client.mutation(api.trechos.iniciarImportacaoLoteAsync, {
        sessionToken: this.tokenSessao(),
        regiao: this.regiao,
        ano: this.ano,
        mes: this.mes,
        pavArquivoOrigem: this.arquivoPavSelecionado.name,
        pavArquivoBase64,
        naoPavArquivoOrigem: this.arquivoNaoPavSelecionado.name,
        naoPavArquivoBase64,
        limparAntes: this.uploadLimparAntes,
        dryRun: false,
        operador: this.operador || undefined,
        perfil: this.perfil,
      });

      const ids = [String(resultado.importacoes.PAV), String(resultado.importacoes.NAO_PAV)];
      this.uploadMensagem = `Lote enfileirado. IDs: ${ids.join(", ")}`;

      if (this.uploadProcessarComplementar) {
        await this.client.action(api.workbook.importarWorkbookComplementar, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: "PAV",
          arquivoOrigem: this.arquivoPavSelecionado.name,
          arquivoBase64: pavArquivoBase64,
          limparAntes: true,
          operador: this.operador || undefined,
          perfil: this.perfil,
        });
        await this.client.action(api.workbook.importarWorkbookComplementar, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: "NAO_PAV",
          arquivoOrigem: this.arquivoNaoPavSelecionado.name,
          arquivoBase64: naoPavArquivoBase64,
          limparAntes: true,
          operador: this.operador || undefined,
          perfil: this.perfil,
        });
      }

      await this.acompanharImportacoes(ids);
      this.arquivoPavSelecionado = null;
      this.arquivoNaoPavSelecionado = null;
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.uploadMensagem = e instanceof Error ? `Falha no lote: ${e.message}` : String(e);
    } finally {
      this.uploadEmAndamento = false;
    }
  }

  private async acompanharImportacoes(importacaoIds: string[]): Promise<void> {
    const espera = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

    for (let tentativa = 0; tentativa < 120; tentativa += 1) {
      const statusList = (await Promise.all(
        importacaoIds.map((id) =>
          this.client.query(api.trechos.obterStatusImportacao, {
            importacaoId: id as any,
          }),
        ),
      )) as ImportacaoStatus[];

      const pendentes = statusList.filter((s: ImportacaoStatus) => s.status === "PROCESSANDO");
      if (pendentes.length > 0) {
        this.uploadMensagem = `Processando ${pendentes.length}/${statusList.length} importacao(oes)...`;
        await espera(1500);
        continue;
      }

      const totalGravados = statusList.reduce((acc: number, s: ImportacaoStatus) => acc + s.gravados, 0);
      const totalIgnoradas = statusList.reduce((acc: number, s: ImportacaoStatus) => acc + s.linhasIgnoradas, 0);
      const totalErros = statusList.reduce((acc: number, s: ImportacaoStatus) => acc + s.linhasComErro, 0);
      const erros = statusList.filter((s: ImportacaoStatus) => s.status === "ERRO");

      if (erros.length > 0) {
        this.uploadMensagem = `Concluido com falha em ${erros.length} importacao(oes). Consulte o historico.`;
      } else {
        this.uploadMensagem = `Importacao concluida: ${totalGravados} gravados, ${totalIgnoradas} ignoradas, ${totalErros} erros.`;
      }

      this.arquivoSelecionado = null;
      await this.recarregar();
      return;
    }

    this.uploadMensagem = "Processamento ainda em andamento. Atualize em alguns segundos.";
  }

  async recarregar(): Promise<void> {
    this.loading = true;
    this.erro = "";
    try {
      const [graficos, inconsistencias, auditoria, saude] = await Promise.all([
        this.client.query(api.trechos.obterGraficosCompetencia, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.trechos.obterInconsistenciasImportacao, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.trechos.listarAuditoriaRecente, {
          sessionToken: this.tokenSessao(),
          limite: 30,
        }),
        this.client.query(api.trechos.obterSaudeOperacional, {
          sessionToken: this.tokenSessao(),
        }),
      ]);

      this.totalTrechos = graficos.kpis.totalTrechos;
      this.totalKm = graficos.kpis.totalKm;
      this.programados = graficos.kpis.programadosNoMes;
      this.naoProgramados = graficos.kpis.naoProgramadosNoMes;

      this.porTipoFonte = graficos.series.porTipoFonte;
      this.topSrePorKm = graficos.series.topSrePorKm;

      this.importacoes = inconsistencias.importacoes as ImportacaoResumo[];
      this.porCodigoInconsistencia = inconsistencias.porCodigo as CodigoInconsistenciaResumo[];
      this.resumoInconsistencias = inconsistencias.resumo;
      this.auditoriaRecente = auditoria as AuditoriaEvento[];
      this.saudeOperacional = saude as SaudeOperacional;
      await this.carregarUsuariosAdmin();
      await this.recarregarGraficosWorkbook();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.erro = e instanceof Error ? e.message : String(e);
    } finally {
      this.loading = false;
    }
  }

  maxKm(): number {
    const max = this.topSrePorKm.reduce((acc, item) => Math.max(acc, item.km), 0);
    return max || 1;
  }

  barraKm(valor: number): string {
    return `${Math.max(4, (valor / this.maxKm()) * 100)}%`;
  }

  maxInconsistenciasPorCodigo(): number {
    const max = this.porCodigoInconsistencia.reduce((acc, item) => Math.max(acc, item.total), 0);
    return max || 1;
  }

  barraInconsistencia(total: number): string {
    return `${Math.max(4, (total / this.maxInconsistenciasPorCodigo()) * 100)}%`;
  }

  importacoesOrdenadas(): ImportacaoResumo[] {
    return [...this.importacoes].sort((a, b) => {
      const aTempo = a.finalizadoEm ?? a.iniciadoEm ?? 0;
      const bTempo = b.finalizadoEm ?? b.iniciadoEm ?? 0;
      return bTempo - aTempo;
    });
  }

  ultimasImportacoesPorTipo(): ImportacaoResumo[] {
    const latestByTipo = new Map<string, ImportacaoResumo>();

    for (const item of this.importacoesOrdenadas()) {
      if (!latestByTipo.has(item.tipoFonte)) {
        latestByTipo.set(item.tipoFonte, item);
      }
    }

    return Array.from(latestByTipo.values());
  }

  importacoesVisiveis(): ImportacaoResumo[] {
    return this.mostrarHistoricoCompleto ? this.importacoesOrdenadas() : this.ultimasImportacoesPorTipo();
  }

  totalIgnoradasUltimasImportacoes(): number {
    return this.ultimasImportacoesPorTipo().reduce((acc, item) => acc + item.linhasIgnoradas, 0);
  }

  totalRecebidasUltimasImportacoes(): number {
    return this.ultimasImportacoesPorTipo().reduce(
      (acc, item) => acc + item.linhasValidas + item.linhasIgnoradas + item.linhasComErro,
      0,
    );
  }

  alertaIgnoradasElevadas(): boolean {
    const ignoradas = this.totalIgnoradasUltimasImportacoes();
    const recebidas = this.totalRecebidasUltimasImportacoes();
    if (recebidas === 0) return false;

    const percentualIgnoradas = ignoradas / recebidas;
    return ignoradas >= 20 && percentualIgnoradas >= 0.25;
  }

  textoAlertaIgnoradas(): string {
    const ignoradas = this.totalIgnoradasUltimasImportacoes();
    const recebidas = this.totalRecebidasUltimasImportacoes();
    const percentual = recebidas === 0 ? 0 : Math.round((ignoradas / recebidas) * 100);

    return `${ignoradas} linhas ignoradas (${percentual}% da carga recente). Verifique se a planilha possui linhas de outras regioes.`;
  }

  private percent(part: number, total: number): number {
    if (total <= 0) return 0;
    return Number(((part / total) * 100).toFixed(1));
  }

  dadosPizzaTipoFonte(): Array<{ label: string; valor: number; percentual: number; color: string }> {
    const total = this.porTipoFonte.reduce((acc, item) => acc + item.totalTrechos, 0);
    const palette = ["#d95f02", "#1b9e77", "#457b9d", "#e9c46a"];
    return this.porTipoFonte.map((item, idx) => ({
      label: item.tipoFonte,
      valor: item.totalTrechos,
      percentual: this.percent(item.totalTrechos, total),
      color: palette[idx % palette.length],
    }));
  }

  dadosPizzaProgramacao(): Array<{ label: string; valor: number; percentual: number; color: string }> {
    const total = this.programados + this.naoProgramados;
    return [
      {
        label: "Programados",
        valor: this.programados,
        percentual: this.percent(this.programados, total),
        color: "#2a9d8f",
      },
      {
        label: "Nao programados",
        valor: this.naoProgramados,
        percentual: this.percent(this.naoProgramados, total),
        color: "#e76f51",
      },
    ];
  }

  estiloPizzaConica(data: Array<{ percentual: number; color: string }>): string {
    let angle = 0;
    const slices = data
      .map((item) => {
        const start = angle;
        const inc = (item.percentual / 100) * 360;
        angle += inc;
        return `${item.color} ${start}deg ${angle}deg`;
      })
      .join(", ");
    return `conic-gradient(${slices || "#e2e8f0 0deg 360deg"})`;
  }

  private paletteWorkbook = ["#d95f02", "#1b9e77", "#457b9d", "#e9c46a", "#e76f51", "#6d597a", "#264653"];

  workbookAbasDisponiveis(): string[] {
    const abas = Array.from(new Set(this.workbookGraficos.map((g) => g.aba)));
    return abas.sort((a, b) => {
      const na = Number(a);
      const nb = Number(b);
      if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
      if (Number.isFinite(na)) return -1;
      if (Number.isFinite(nb)) return 1;
      return a.localeCompare(b, "pt-BR");
    });
  }

  graficosWorkbookFiltrados(): WorkbookGrafico[] {
    if (!this.workbookAbaSelecionada) return this.workbookGraficos;
    return this.workbookGraficos.filter((g) => g.aba === this.workbookAbaSelecionada);
  }

  estiloPizzaWorkbook(grafico: WorkbookGrafico): string {
    const data = grafico.series.map((s, i) => ({ percentual: s.percentual, color: this.paletteWorkbook[i % this.paletteWorkbook.length] }));
    return this.estiloPizzaConica(data);
  }

  corSerieWorkbook(index: number): string {
    return this.paletteWorkbook[index % this.paletteWorkbook.length];
  }

  async recarregarGraficosWorkbook(): Promise<void> {
    if (!this.sessaoAtual?.token) return;

    const result = await this.client.query(api.workbook.listarGraficosWorkbook, {
      sessionToken: this.tokenSessao(),
      regiao: this.regiao,
      ano: this.ano,
      mes: this.mes,
      tipoFonte: this.workbookTipoFonte,
    });

    this.workbookGraficos = result as WorkbookGrafico[];
    const abas = this.workbookAbasDisponiveis();
    if (abas.length === 0) {
      this.workbookAbaSelecionada = "";
    } else if (!abas.includes(this.workbookAbaSelecionada)) {
      this.workbookAbaSelecionada = abas[0];
    }
  }

  formatarDataImportacao(item: ImportacaoResumo): string {
    const ts = item.finalizadoEm ?? item.iniciadoEm;
    if (!ts) return "Sem horario";

    return new Intl.DateTimeFormat("pt-BR", {
      dateStyle: "short",
      timeStyle: "short",
    }).format(ts);
  }

  nomeArquivoMarkdown(): string {
    return `relatorio_regiao_${this.regiao}_${this.competencia()}.md`;
  }

  gerarMarkdownCompetencia(): string {
    const linhas = [
      `# Relatorio Tecnico - Regiao ${this.regiao}`,
      "",
      `- Competencia: ${this.competencia()}`,
      `- Gerado em: ${new Date().toISOString()}`,
      "",
      "## Indicadores",
      `- Total de trechos: ${this.totalTrechos}`,
      `- Total de extensao (km): ${this.totalKm.toFixed(2)}`,
      `- Programados no mes: ${this.programados}`,
      `- Nao programados no mes: ${this.naoProgramados}`,
      "",
      "## Distribuicao por Tipo de Fonte",
      ...this.porTipoFonte.map(
        (item) => `- ${item.tipoFonte}: ${item.totalTrechos} trechos / ${item.totalKm.toFixed(2)} km`,
      ),
      "",
      "## Top SRE por KM",
      ...this.topSrePorKm.map((item) => `- ${item.sre}: ${item.km.toFixed(2)} km`),
      "",
      "## Inconsistencias de Importacao",
      `- Total de importacoes: ${this.resumoInconsistencias.totalImportacoes}`,
      `- Importacoes com erro: ${this.resumoInconsistencias.importacoesComErro}`,
      `- Total de erros: ${this.resumoInconsistencias.totalErros}`,
      ...this.porCodigoInconsistencia.map((item) => `- ${item.codigo}: ${item.total}`),
      "",
      "## Historico de Importacoes",
      ...this.importacoesOrdenadas().flatMap((item) => [
        `- ${this.formatarDataImportacao(item)} | ${item.tipoFonte} | ${item.status}`,
        `  - Arquivo: ${item.arquivoOrigem}`,
        `  - Validas: ${item.linhasValidas} | Ignoradas: ${item.linhasIgnoradas} | Erros: ${item.linhasComErro} | Gravados: ${item.gravados}`,
      ]),
      "",
    ];

    return linhas.join("\n");
  }

  baixarMarkdownCompetencia(): void {
    const blob = new Blob([this.gerarMarkdownCompetencia()], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = this.nomeArquivoMarkdown();
    a.click();
    URL.revokeObjectURL(url);
  }

  private async baixarDocx(nomeArquivo: string, linhas: string[]): Promise<void> {
    const children: Paragraph[] = [];
    for (const linha of linhas) {
      if (linha.startsWith("# ")) {
        children.push(new Paragraph({ text: linha.replace(/^#\s+/, ""), heading: HeadingLevel.TITLE }));
      } else if (linha.startsWith("## ")) {
        children.push(new Paragraph({ text: linha.replace(/^##\s+/, ""), heading: HeadingLevel.HEADING_1 }));
      } else if (linha.startsWith("- ")) {
        children.push(new Paragraph({ text: linha.slice(2), bullet: { level: 0 } }));
      } else if (linha.trim() === "") {
        children.push(new Paragraph({ text: "" }));
      } else {
        children.push(new Paragraph({ text: linha }));
      }
    }

    const doc = new Document({ sections: [{ children }] });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = nomeArquivo;
    a.click();
    URL.revokeObjectURL(url);
  }

  private mesLabel(value: number): string {
    return this.meses.find((m) => m.value === value)?.label ?? String(value);
  }

  private formatarNumero(value: number, casas = 2): string {
    return new Intl.NumberFormat("pt-BR", {
      minimumFractionDigits: casas,
      maximumFractionDigits: casas,
    }).format(value);
  }

  private formatarDataHora(ts?: number): string {
    if (!ts) return "-";
    return new Intl.DateTimeFormat("pt-BR", {
      dateStyle: "short",
      timeStyle: "short",
    }).format(ts);
  }

  private dataUrlParaBytes(dataUrl: string): Uint8Array {
    const base64 = dataUrl.split(",")[1] ?? "";
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }

  private async renderizarPizzaWorkbook(titulo: string, series: WorkbookSerie[]): Promise<Uint8Array> {
    const width = 980;
    const height = 560;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Falha ao inicializar contexto grafico do relatorio.");

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    ctx.fillStyle = "#1f2933";
    ctx.font = "bold 24px DM Sans, Segoe UI, sans-serif";
    ctx.fillText(titulo.slice(0, 82), 30, 42);

    const total = series.reduce((acc, s) => acc + (s.valor ?? 0), 0);
    const cx = 270;
    const cy = 300;
    const radius = 170;
    let current = -Math.PI / 2;

    for (let i = 0; i < series.length; i += 1) {
      const s = series[i];
      const percentual = total > 0 ? (s.valor / total) * 100 : s.percentual;
      const angle = (percentual / 100) * Math.PI * 2;
      const color = this.paletteWorkbook[i % this.paletteWorkbook.length];

      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, radius, current, current + angle);
      ctx.closePath();
      ctx.fillStyle = color;
      ctx.fill();
      current += angle;
    }

    ctx.fillStyle = "#ffffff";
    ctx.beginPath();
    ctx.arc(cx, cy, 70, 0, Math.PI * 2);
    ctx.fill();

    ctx.fillStyle = "#334e68";
    ctx.font = "bold 18px DM Sans, Segoe UI, sans-serif";
    ctx.fillText("Total", cx - 24, cy - 6);
    ctx.font = "bold 20px DM Sans, Segoe UI, sans-serif";
    ctx.fillText(this.formatarNumero(total, 2), cx - 42, cy + 20);

    let y = 110;
    for (let i = 0; i < series.length; i += 1) {
      const s = series[i];
      const color = this.paletteWorkbook[i % this.paletteWorkbook.length];

      ctx.fillStyle = color;
      ctx.fillRect(520, y - 12, 18, 18);
      ctx.fillStyle = "#102a43";
      ctx.font = "16px DM Sans, Segoe UI, sans-serif";
      const linha = `${s.label}: ${this.formatarNumero(s.valor, 2)} (${this.formatarNumero(s.percentual, 1)}%)`;
      ctx.fillText(linha.slice(0, 60), 548, y + 2);
      y += 28;
      if (y > 520) break;
    }

    return this.dataUrlParaBytes(canvas.toDataURL("image/png"));
  }

  private selecionarPrimeiroGraficoPorTitulo(graficos: WorkbookGrafico[]): WorkbookGrafico[] {
    const ordenarAba = (a: WorkbookGrafico, b: WorkbookGrafico): number => {
      if (a.aba === b.aba) return a.ordem - b.ordem;
      if (a.aba === "TT" && b.aba !== "TT") return 1;
      if (b.aba === "TT" && a.aba !== "TT") return -1;

      const na = Number(a.aba);
      const nb = Number(b.aba);
      if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
      if (Number.isFinite(na)) return -1;
      if (Number.isFinite(nb)) return 1;
      return a.aba.localeCompare(b.aba, "pt-BR");
    };

    const porTitulo = new Map<string, WorkbookGrafico[]>();
    for (const g of graficos) {
      const titulo = g.titulo.trim().toUpperCase();
      const lista = porTitulo.get(titulo) ?? [];
      lista.push(g);
      porTitulo.set(titulo, lista);
    }

    const selecionados: WorkbookGrafico[] = [];
    for (const lista of porTitulo.values()) {
      const ordenada = [...lista].sort(ordenarAba);
      const preferido = ordenada.find((g) => g.aba !== "TT") ?? ordenada[0];
      if (preferido) selecionados.push(preferido);
    }

    return selecionados.sort((a, b) => a.titulo.localeCompare(b.titulo, "pt-BR"));
  }

  private tabela2Colunas(linhas: Array<[string, string]>): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: linhas.map(
        ([k, v]) =>
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: k, bold: true })] })] }),
              new TableCell({ children: [new Paragraph(v)] }),
            ],
          }),
      ),
    });
  }

  async baixarDocxCompetencia(): Promise<void> {
    if (!this.sessaoAtual?.token) return;

    try {
      const [payload, graficos, inconsistencias, base, workbookPavRaw, workbookNaoPavRaw] = await Promise.all([
        this.client.query(api.trechos.gerarPayloadRelatorio, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.trechos.obterGraficosCompetencia, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.trechos.obterInconsistenciasImportacao, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.trechos.obterBaseConsolidada, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.workbook.listarGraficosWorkbook, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: "PAV",
        }),
        this.client.query(api.workbook.listarGraficosWorkbook, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: "NAO_PAV",
        }),
      ]);

      const trechos = (base as any).trechos as Array<any>;
      const importacoes = (base as any).importacoes as Array<any>;

      const pav = trechos.filter((t) => t.tipoFonte === "PAV");
      const naoPav = trechos.filter((t) => t.tipoFonte === "NAO_PAV");

      const pavKm = pav.reduce((acc, t) => acc + (t.extKm ?? 0), 0);
      const naoPavKm = naoPav.reduce((acc, t) => acc + (t.extKm ?? 0), 0);

      const workbookPav = this.selecionarPrimeiroGraficoPorTitulo(workbookPavRaw as WorkbookGrafico[]);
      const workbookNaoPav = this.selecionarPrimeiroGraficoPorTitulo(workbookNaoPavRaw as WorkbookGrafico[]);

      const children: Array<Paragraph | Table> = [
        new Paragraph({
          text: "ESTADO DO TOCANTINS",
          alignment: AlignmentType.CENTER,
          heading: HeadingLevel.HEADING_1,
        }),
        new Paragraph({
          text: "AGENCIA DE TRANSPORTES, OBRAS E INFRAESTRUTURA - AGETO",
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({ text: "" }),
        new Paragraph({
          text: "PRODUTO 02 - RELATORIO DE ACOMPANHAMENTO TECNICO / AMBIENTAL",
          alignment: AlignmentType.CENTER,
          heading: HeadingLevel.HEADING_2,
        }),
        new Paragraph({
          text: `REGIAO ${String(this.regiao).padStart(2, "0")} DE CONSERVACAO`,
          alignment: AlignmentType.CENTER,
          heading: HeadingLevel.HEADING_2,
        }),
        new Paragraph({
          text: `${this.mesLabel(this.mes).toUpperCase()} DE ${this.ano}`,
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({ children: [new PageBreak()] }),

        new Paragraph({ text: "SUMARIO", heading: HeadingLevel.HEADING_1 }),
        new TableOfContents("", { headingStyleRange: "1-3", hyperlink: true }),
        new Paragraph({ children: [new PageBreak()] }),

        new Paragraph({ text: "1. Apresentacao", heading: HeadingLevel.HEADING_1 }),
        new Paragraph(
          `Este relatorio sintetiza os resultados da competencia ${this.competencia()} para a Regiao ${this.regiao}, com base nas importacoes de planilhas, auditoria operacional e graficos complementares do workbook.`,
        ),
        new Paragraph({ text: "" }),

        new Paragraph({ text: "2. Indicadores Gerais", heading: HeadingLevel.HEADING_1 }),
        this.tabela2Colunas([
          ["Competencia", this.competencia()],
          ["Total de trechos", String((graficos as any).kpis.totalTrechos)],
          ["Total de extensao (km)", this.formatarNumero((graficos as any).kpis.totalKm, 2)],
          ["Programados no mes", String((graficos as any).kpis.programadosNoMes)],
          ["Nao programados no mes", String((graficos as any).kpis.naoProgramadosNoMes)],
          ["Percentual programados", `${(payload as any).graficos.kpis.percentualProgramados}%`],
        ]),
        new Paragraph({ text: "" }),

        new Paragraph({ text: "3. Vistorias Rotineiras no Periodo", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: "3.1 Rodovias Pavimentadas", heading: HeadingLevel.HEADING_2 }),
        this.tabela2Colunas([
          ["Trechos (linhas)", String(pav.length)],
          ["Trechos unicos", String(new Set(pav.map((t) => t.trecho)).size)],
          ["Extensao total (km)", this.formatarNumero(pavKm, 2)],
          ["Graficos workbook (total)", String((workbookPavRaw as WorkbookGrafico[]).length)],
          ["Abas com grafico", String(new Set((workbookPavRaw as WorkbookGrafico[]).map((g) => g.aba)).size)],
        ]),
        new Paragraph({ text: "" }),

        new Paragraph({ text: "3.2 Rodovias Nao Pavimentadas", heading: HeadingLevel.HEADING_2 }),
        this.tabela2Colunas([
          ["Trechos (linhas)", String(naoPav.length)],
          ["Trechos unicos", String(new Set(naoPav.map((t) => t.trecho)).size)],
          ["Extensao total (km)", this.formatarNumero(naoPavKm, 2)],
          ["Graficos workbook (total)", String((workbookNaoPavRaw as WorkbookGrafico[]).length)],
          ["Abas com grafico", String(new Set((workbookNaoPavRaw as WorkbookGrafico[]).map((g) => g.aba)).size)],
        ]),
        new Paragraph({ text: "" }),

        new Paragraph({ text: "4. Controle de Importacoes e Inconsistencias", heading: HeadingLevel.HEADING_1 }),
        this.tabela2Colunas([
          ["Total de importacoes", String((inconsistencias as any).resumo.totalImportacoes)],
          ["Importacoes com erro", String((inconsistencias as any).resumo.importacoesComErro)],
          ["Total de erros", String((inconsistencias as any).resumo.totalErros)],
          ["Ultima importacao", this.formatarDataHora(importacoes[0]?.finalizadoEm ?? importacoes[0]?.iniciadoEm)],
        ]),
      ];

      children.push(new Paragraph({ text: "" }));
      children.push(new Paragraph({ text: "5. Graficos do Workbook", heading: HeadingLevel.HEADING_1 }));
      children.push(
        new Paragraph(
          "Para padronizacao da emissao automatica, esta secao inclui o primeiro grafico de cada titulo no workbook (priorizando as abas numericas e usando TT como fallback), mantendo os valores e percentuais extraidos.",
        ),
      );

      children.push(new Paragraph({ text: "" }));
      children.push(new Paragraph({ text: "5.1 Rodovias Pavimentadas", heading: HeadingLevel.HEADING_2 }));
      if (workbookPav.length === 0) {
        children.push(
          new Paragraph(
            "Nenhum grafico de workbook foi encontrado para Rodovias Pavimentadas nesta competencia. Verifique se a importacao complementar PAV foi executada.",
          ),
        );
      } else {
        for (const g of workbookPav) {
          const image = await this.renderizarPizzaWorkbook(g.titulo, g.series);
          children.push(new Paragraph({ text: `Aba ${g.aba} - ${g.titulo}`, heading: HeadingLevel.HEADING_3 }));
          children.push(
            new Paragraph({
              children: [new ImageRun({ type: "png", data: image, transformation: { width: 620, height: 355 } })],
            }),
          );
        }
      }

      children.push(new Paragraph({ text: "" }));
      children.push(new Paragraph({ text: "5.2 Rodovias Nao Pavimentadas", heading: HeadingLevel.HEADING_2 }));
      if (workbookNaoPav.length === 0) {
        children.push(
          new Paragraph(
            "Nenhum grafico de workbook foi encontrado para Rodovias Nao Pavimentadas nesta competencia. Verifique se a importacao complementar NAO_PAV foi executada.",
          ),
        );
      } else {
        for (const g of workbookNaoPav) {
          const image = await this.renderizarPizzaWorkbook(g.titulo, g.series);
          children.push(new Paragraph({ text: `Aba ${g.aba} - ${g.titulo}`, heading: HeadingLevel.HEADING_3 }));
          children.push(
            new Paragraph({
              children: [new ImageRun({ type: "png", data: image, transformation: { width: 620, height: 355 } })],
            }),
          );
        }
      }

      children.push(new Paragraph({ children: [new PageBreak()] }));
      children.push(new Paragraph({ text: "6. Conclusoes Automaticas", heading: HeadingLevel.HEADING_1 }));
      for (const obs of (payload as any).observacoesAutomaticas as string[]) {
        children.push(new Paragraph({ text: obs, bullet: { level: 0 } }));
      }

      const doc = new Document({ sections: [{ children }] });
      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `relatorio_regiao_${this.regiao}_${this.competencia()}.docx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      this.erro = e instanceof Error ? e.message : String(e);
    }
  }

  async baixarMarkdownConsolidado(): Promise<void> {
    this.gerandoConsolidado = true;
    this.erro = "";

    try {
      const resultados = await Promise.all(
        this.regioes.map(async (regiao) => {
          const [graficos, inconsistencias] = await Promise.all([
            this.client.query(api.trechos.obterGraficosCompetencia, {
              regiao,
              ano: this.ano,
              mes: this.mes,
            }),
            this.client.query(api.trechos.obterInconsistenciasImportacao, {
              regiao,
              ano: this.ano,
              mes: this.mes,
            }),
          ]);

          return {
            regiao,
            totalTrechos: graficos.kpis.totalTrechos,
            totalKm: graficos.kpis.totalKm,
            totalImportacoes: inconsistencias.resumo.totalImportacoes,
            importacoesComErro: inconsistencias.resumo.importacoesComErro,
            totalErros: inconsistencias.resumo.totalErros,
            porTipoFonte: graficos.series.porTipoFonte as FonteResumo[],
          };
        }),
      );

      const totais = resultados.reduce(
        (acc, item) => {
          acc.totalTrechos += item.totalTrechos;
          acc.totalKm += item.totalKm;
          acc.totalImportacoes += item.totalImportacoes;
          acc.importacoesComErro += item.importacoesComErro;
          acc.totalErros += item.totalErros;
          return acc;
        },
        {
          totalTrechos: 0,
          totalKm: 0,
          totalImportacoes: 0,
          importacoesComErro: 0,
          totalErros: 0,
        },
      );

      const linhas = [
        "# Relatorio Tecnico Consolidado",
        "",
        `- Competencia: ${this.competencia()}`,
        `- Regioes: ${this.regioes.join(", ")}`,
        `- Gerado em: ${new Date().toISOString()}`,
        "",
        "## Totais Consolidados",
        `- Total de trechos: ${totais.totalTrechos}`,
        `- Total de extensao (km): ${totais.totalKm.toFixed(2)}`,
        `- Total de importacoes: ${totais.totalImportacoes}`,
        `- Importacoes com erro: ${totais.importacoesComErro}`,
        `- Total de erros: ${totais.totalErros}`,
        "",
        "## Resumo por Regiao",
        ...resultados.flatMap((item) => [
          `### Regiao ${item.regiao}`,
          `- Total de trechos: ${item.totalTrechos}`,
          `- Total de extensao (km): ${item.totalKm.toFixed(2)}`,
          `- Total de importacoes: ${item.totalImportacoes}`,
          `- Importacoes com erro: ${item.importacoesComErro}`,
          `- Total de erros: ${item.totalErros}`,
          "",
          "- Distribuicao por tipo de fonte:",
          ...item.porTipoFonte.map(
            (fonte) => `  - ${fonte.tipoFonte}: ${fonte.totalTrechos} trechos / ${fonte.totalKm.toFixed(2)} km`,
          ),
          "",
          `- Relatorio detalhado: relatorio_regiao_${item.regiao}_${this.competencia()}.md`,
          "",
        ]),
      ];

      const blob = new Blob([linhas.join("\n")], { type: "text/markdown;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `relatorio_consolidado_${this.competencia()}.md`;
      a.click();
      URL.revokeObjectURL(url);

      this._ultimasLinhasConsolidado = linhas;
    } catch (e) {
      this.erro = e instanceof Error ? e.message : String(e);
    } finally {
      this.gerandoConsolidado = false;
    }
  }

  private _ultimasLinhasConsolidado: string[] | null = null;

  async baixarDocxConsolidado(): Promise<void> {
    if (!this._ultimasLinhasConsolidado) {
      await this.baixarMarkdownConsolidado();
    }
    if (!this._ultimasLinhasConsolidado) return;
    await this.baixarDocx(`relatorio_consolidado_${this.competencia()}.docx`, this._ultimasLinhasConsolidado);
  }

  competencia(): string {
    return `${this.ano}-${String(this.mes).padStart(2, "0")}`;
  }
}
