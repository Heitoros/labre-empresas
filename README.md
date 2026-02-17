 LABRE Empresas
Sistema web para ingestão, validação e consolidação de dados rodoviários, com geração automatizada de relatórios técnicos (`.md` e `.docx`), incluindo relatórios regionais, consolidados e workbook completo de gráficos.
 Funcionalidades
- Upload em tempo real de planilhas (`.xlsx`)
- Importação assíncrona (individual e lote PAV + NAO_PAV)
- Validação e tratamento de inconsistências
- Painel com:
  - KPIs por competência
  - inconsistências por código
  - histórico de importações
  - jobs e auditoria operacional
  - saúde operacional (últimas 24h)
- Autenticação e sessão com perfis:
  - `OPERADOR`
  - `GESTOR`
  - `ADMIN`
- Troca obrigatória de senha no primeiro acesso
- Gestão de usuários (ADMIN)
- Geração de relatórios:
  - regional (`markdown` e `docx`)
  - consolidado (`markdown` e `docx`)
  - oficial (`docx`)
  - workbook completo com todos os gráficos
 Stack
- Frontend: Angular
- Backend/DB: Convex
- Scripts: Node + TypeScript
- Geração DOCX: `docx`
- Leitura planilhas: `xlsx`
- Renderização de gráficos em lote: `chart.js` + `chartjs-node-canvas`
 Estrutura principal
- `src/` – aplicação Angular
- `convex/` – schema, autenticação, segurança e regras de negócio
- `scripts/` – automações de importação e relatórios
- `relatorios/` – saídas geradas
- `OPERACAO_DIARIA.txt` – playbook operacional
- `DEPLOY_CHECKLIST.txt` – checklist de deploy/homologação
- `RELATORIO_IMPLEMENTACAO.txt` – histórico técnico de implementação
 Requisitos
- Node.js 20+
- npm
- Conta/Deploy Convex configurado (`CONVEX_URL`)
 Como rodar local
npm install
npm run dev
Build
npm run build
Deploy backend Convex
npm run convex:deploy
Validação pré-release
npm run release:check
Fluxo operacional principal
- Importação mensal (dry-run):
npm run fechamento:dezembro:dry
- Importação mensal oficial:
npm run fechamento:dezembro
- DOCX regionais + consolidado:
npm run fechamento:dezembro:docx
- Workbooks completos (todos os gráficos):
npm run workbook:docx:todos
Segurança
- Sessão obrigatória para operações críticas
- Controle por perfil
- Política de senha mínima:
  - 8+ caracteres
  - maiúscula
  - minúscula
  - número
  - caractere especial
- Auditoria de eventos operacionais
Status do projeto
Em fase de hardening/deploy, com backend em produção e frontend publicado.