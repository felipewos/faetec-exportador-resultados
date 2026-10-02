# Chrome Web Store Listing — FAETEC: Exportador de Resultados

> Última Atualização: Outubro de 2026
> Versão do Pacote: 1.0.0
> Arquivo pronto para upload: `faetec-exportador-v1.0.0.zip`

---

## 1. Ficha da Loja (Store Listing)

**Nome da Extensão (Extension Name)** [OBRIGATÓRIO - máx 75 caracteres]
```text
FAETEC - Exportador de Resultados
```

**Descrição Curta (Short Description)** [OBRIGATÓRIO - máx 132 caracteres]
```text
Exporta resultados de processos seletivos da FAETEC diretamente para planilhas nativas do Excel (.xlsx).
```

**Descrição Detalhada (Detailed Description)** [OBRIGATÓRIO - máx 16.000 caracteres]
```text
Exporte os resultados e convocações dos processos seletivos da FAETEC diretamente para planilhas do Microsoft Excel (.xlsx) de forma rápida, organizada e sem complicação.

O FAETEC - Exportador de Resultados foi criado para facilitar a vida de candidatos e comissões que precisam analisar listas de classificação com dezenas de disciplinas e regiões.

Principais Recursos:

- Exportação Completa em 1 Clique: Varre todos os cargos, disciplinas e regiões do edital, consolidando milhares de registros em um único arquivo Excel sem duplicatas.
- Planilhas Nativas com Filtros Ativos: O arquivo .xlsx gerado já vem com AutoFiltro habilitado em todas as 13 colunas e a primeira linha de cabeçalho congelada, permitindo filtrar por Convocados, Reservas, Polos ou Notas instantaneamente.
- Navegação e Paginação Automática: Percorre todas as páginas de resultados de cada disciplina sem perder nenhum candidato e sem misturar nomes entre polos vizinhos.
- Formatos Flexíveis: Escolha entre Excel (.xlsx) nativo ou CSV delimitado por ponto e vírgula.
- Painel Flutuante Prático: Acompanhe o progresso da coleta em tempo real diretamente na página da FAETEC, com botão de parada e download imediato a qualquer momento.

Como Usar:

1. Acesse a página oficial de resultados do processo seletivo da FAETEC.
2. O painel do Exportador aparecerá automaticamente no canto inferior direito.
3. Clique em "Exportar Edital Completo" para baixar todo o concurso ou selecione um cargo/disciplina e clique em "Exportar Todo este Cargo".
4. Ao concluir, a planilha formatada será baixada diretamente no seu computador.

Privacidade e Segurança:

- 100% Local e Seguro: Todo o processamento é feito exclusivamente no seu navegador. A extensão não coleta, não armazena e não envia dados pessoais para servidores externos.
- Conexão Direta: Os dados são lidos exclusivamente das tabelas públicas do portal oficial da FAETEC.
```

**Categoria (Category)** [OBRIGATÓRIO]
```text
Ferramentas de Pesquisa (Search Tools) ou Produtividade (Productivity)
```

**Propósito Único (Single Purpose)** [OBRIGATÓRIO]
```text
Extrair e formatar resultados públicos do portal de seleção da FAETEC em planilhas Excel (.xlsx) locais.
```

**Idioma Principal (Primary Language)** [OBRIGATÓRIO]
```text
Português (Brasil)
```

---

## 2. Recursos Gráficos e Imagens (Graphics & Assets)

| Recurso | Dimensões | Status | Arquivo no Projeto |
| :--- | :--- | :--- | :--- |
| **Ícone da Loja** [OBRIGATÓRIO] | 128×128 PNG | ✅ Pronto | `icons/icon128.png` |
| **Screenshot 1** [OBRIGATÓRIO] | 1280×800 PNG | ✅ Pronto | `screenshot-1280x800.png` |
| **Ícone Menor** | 48×48 PNG | ✅ Pronto | `icons/icon48.png` |
| **Ícone de Barra** | 16×16 PNG | ✅ Pronto | `icons/icon16.png` |

---

## 3. Justificativa de Permissões (Permissions Justification)

> A equipe de revisão do Google Chrome Web Store exige uma explicação clara e em inglês/português para cada permissão solicitada no `manifest.json`. Copie e cole os textos abaixo:

### Justificativas em Inglês (Para o formulário do Google):

| Permission | Type | Justification for Chrome Web Store Reviewers |
| :--- | :--- | :--- |
| `storage` | permissions | Used exclusively to save the user's preferred export spreadsheet format (e.g., XLSX vs. CSV) locally in browser storage. |
| `tabs` | permissions | Required by the extension popup to inspect the active tab URL, verify if the user is on the official FAETEC portal, and communicate status. |
| `https://processoseletivo.faetec.rj.gov.br/*` | host_permissions | Strictly scoped to the official FAETEC results portal to inject the floating export widget and parse candidate table results into XLSX files. |

### Justificativas em Português:

* **`storage`**: Memoriza no navegador a preferência de formato escolhida pelo usuário (XLSX ou CSV).
* **`tabs`**: Permite ao popup da barra de ferramentas detectar se a aba atual é a página do portal da FAETEC e enviar comandos de exportação para a página.
* **`host_permissions` (`https://processoseletivo.faetec.rj.gov.br/*`)**: Delimita o funcionamento da extensão estritamente ao portal oficial da FAETEC, impedindo acesso a qualquer outro site.

---

## 4. Práticas de Privacidade e Uso de Dados (Data Usage Form)

No painel do Google Developer Dashboard, preencha a aba **Práticas de Privacidade (Privacy Practices)** com as seguintes respostas:

1. **A extensão coleta dados do usuário?**
   - Responda: **NÃO (No)**. A extensão lê dados públicos exibidos na tela da FAETEC e salva diretamente no computador do usuário; nada é coletado ou transmitido pelo desenvolvedor.
2. **Declarações Obrigatórias de Certificação:**
   - [x] Os dados NÃO são vendidos a terceiros.
   - [x] Os dados NÃO são usados para fins alheios à funcionalidade principal.
   - [x] Os dados NÃO são usados para determinar solvabilidade ou concessão de empréstimos.
3. **URL da Política de Privacidade:**
   - Link público do arquivo `PRIVACY.md` (veja instruções abaixo para hospedar gratuitamente no GitHub).

---

## 5. Passo a Passo para Publicar no Chrome Web Store

### Etapa 1: Acessar o Chrome Developer Dashboard
1. Abra o navegador e acesse: [Chrome Developer Dashboard](https://chrome.google.com/webstore/devconsole/)
2. Faça login com a sua conta Google.
3. Caso seja o seu primeiro cadastro como desenvolvedor, o Google solicita o pagamento de uma taxa única de registro de **US$ 5,00**.

### Etapa 2: Fazer o Upload do Pacote
1. No painel, clique no botão **"Adicionar novo item"** (Add new item).
2. Arraste e solte o arquivo **`faetec-exportador-v1.0.0.zip`** (já gerado na raiz da pasta do projeto).
3. O painel lerá o `manifest.json` e criará o rascunho do item automaticamente.

### Etapa 3: Preencher a Ficha da Loja (Store Listing)
1. Preencha **Nome**, **Descrição Curta** e **Descrição Detalhada** copiando os textos da Seção 1 deste documento.
2. No campo de **Ícone**, envie o arquivo `icons/icon128.png`.
3. No campo de **Capturas de Tela**, envie o arquivo `screenshot-1280x800.png`.
4. Selecione a categoria **Ferramentas de Pesquisa** ou **Produtividade**.

### Etapa 4: Preencher as Práticas de Privacidade
1. Cole as justificativas de permissões da Seção 3.
2. Marque que a extensão **NÃO** coleta dados de usuários.
3. Cole a URL pública da sua Política de Privacidade (`PRIVACY.md`).

### Etapa 5: Enviar para Revisão
1. Clique em **Salvar rascunho** e confira se todas as abas estão com visto verde (sem pendências).
2. Clique no botão azul **"Enviar para publicação"** (Submit for Review).
3. A aprovação da Google costuma levar de **24 a 72 horas**. Assim que aprovada, ela estará disponível publicamente na Chrome Web Store para qualquer pessoa instalar com 1 clique!
