# Política de Privacidade — FAETEC: Exportador de Resultados

**Última atualização:** Outubro de 2026

A extensão **FAETEC - Exportador de Resultados** foi desenvolvida com foco total na privacidade, segurança e autonomia dos usuários e candidatos de processos seletivos públicos.

---

### 1. Coleta e Armazenamento de Dados

* **Nenhum Dado Pessoal é Coletado:** A extensão **NÃO coleta, NÃO armazena em servidores externos e NÃO monitora** quaisquer dados pessoais, credenciais, histórico de navegação ou informações do dispositivo.
* **Processamento 100% Local:** A leitura das tabelas públicas de resultados e a conversão para planilha nativa do Excel (`.xlsx`) ou `.csv` ocorrem exclusivamente na memória local do navegador do usuário.
* **Download Direto:** Os arquivos gerados são baixados diretamente no computador do usuário, sem passar por nenhum intermediário ou servidor externo.

---

### 2. Uso de Permissões no Navegador

A extensão solicita apenas as permissões estritamente necessárias para cumprir seu propósito único:

| Permissão | Finalidade |
| :--- | :--- |
| `storage` | Utilizada exclusivamente para memorizar localmente no navegador a preferência de formato de arquivo (ex.: XLSX ou CSV) escolhida pelo usuário. |
| `tabs` | Utilizada pelo popup da extensão para identificar se a aba ativa é o portal oficial da FAETEC e comunicar o status dos dados da página. |
| `host_permissions` (`https://processoseletivo.faetec.rj.gov.br/*`) | Permite a injeção do painel de exportação e a leitura dos resultados públicos exibidos exclusivamente no portal oficial de seleção da FAETEC. |

---

### 3. Comunicação com Serviços Externos

A extensão **NÃO possui servidores próprios**, **NÃO utiliza bibliotecas de telemetria/analytics** e **NÃO realiza conexões com APIs de terceiros**. Ela opera de maneira 100% autônoma e offline após carregada no navegador.

---

### 4. Alterações nesta Política

Caso haja atualizações funcionais que alterem a forma como as permissões são utilizadas, esta política será prontamente atualizada com aviso na nova versão do manifesto.

---

### 5. Contato e Suporte

Para dúvidas, sugestões ou suporte sobre o uso da extensão, utilize a seção de suporte no Chrome Web Store ou abra uma solicitação no repositório oficial do projeto.
