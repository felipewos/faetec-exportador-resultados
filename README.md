# FAETEC - Exportador de Resultados para Excel (.xlsx)

Extensão do Google Chrome (Manifest V3) desenvolvida para extrair com praticidade e precisão a tabela de resultados do processo seletivo da FAETEC (`processoseletivo.faetec.rj.gov.br`) diretamente para planilhas nativas do Excel (.xlsx).

---

## 📋 Colunas Exportadas

O arquivo gerado contém exatamente a estrutura solicitada:

1. **Cargo** *(capturado da seleção do filtro)*
2. **Disciplina** *(capturado da seleção do filtro)*
3. **Região** *(capturado da seleção do filtro)*
4. **Classificação**
5. **Nº de Inscrição**
6. **Nome do Candidato**
7. **Titulação Pref.**
8. **Cota**
9. **Total de Pontos**
10. **Pontos Titulação**
11. **Pontos Experiência**
12. **Data de Nascimento**
13. **Resultados**

---

## 🚀 Como Instalar no Google Chrome

1. Abra o navegador Google Chrome.
2. Na barra de endereços, digite: `chrome://extensions` e pressione **Enter**.
3. No canto superior direito da tela, ative a chave **"Modo do desenvolvedor"**.
4. No canto superior esquerdo, clique no botão **"Carregar sem compactação"**.
5. Selecione a pasta deste projeto:
   `c:\Users\cefet\OneDrive\Documentos\ExtensãoChome\Faetec`
6. Pronto! A extensão estará instalada e ativa.

---

## 💡 Como Usar

1. Acesse a página do processo seletivo:
   `https://processoseletivo.faetec.rj.gov.br/ResultadoTitulacaoExperienciaCotaPne/60`
2. Selecione o **Cargo**, a **Disciplina** e a **Região** desejados e clique no botão **"Buscar"**.
3. Um painel flutuante elegante aparecerá automaticamente no canto inferior direito da tela (ou você pode clicar no ícone da extensão na barra de ferramentas do Chrome).
4. Escolha a opção de exportação desejada:
   - **🌐 Exportar Edital Completo (Tudo)**: A extensão percorre **todos os 4 cargos**, **todas as 136 disciplinas** e **todas as regiões** do Estado, consolidando o resultado geral de todo o processo seletivo em uma única planilha Excel (.xlsx)!
   - **🏢 Exportar Todo este Cargo**: Percorre todas as disciplinas e regiões do cargo atualmente selecionado em uma única planilha.
   - **🚀 Exportar Filtro Atual (Todas as Páginas)**: Varre todas as páginas do filtro específico que você pesquisou.
   - **📥 Exportar Página Atual**: Baixa apenas os candidatos visíveis na tela.
5. Durante a varredura, você acompanha o progresso e o número de candidatos coletados em tempo real, com o botão **💾 Parar e Baixar Agora** para salvar imediatamente a qualquer momento.
6. O download do arquivo `.xlsx` iniciará automaticamente na sua pasta de Downloads.

---

## ⚙️ Formato Nativo do Microsoft Excel (.xlsx)

- O arquivo é gerado diretamente como pasta de trabalho do Excel (**`.xlsx`**) via biblioteca offline embutida (SheetJS).
- **Filtros Automáticos em todas as colunas:** O Excel já abre com as setinhas de filtro ativadas em todas as 13 colunas (*Cargo, Disciplina, Região, Classificação, Cota, Pontuação, etc.*).
- **Cabeçalho Congelado no Topo:** A linha 1 de títulos fica fixada; ao rolar para baixo por milhares de candidatos, os títulos e filtros continuam visíveis.
- **Larguras de colunas pré-ajustadas** automaticamente para visualização limpa e sem textos cortados.
- Sem necessidade de conversão, importação de texto ou configuração de delimitador: abre de forma 100% nativa e com formatação perfeita no Microsoft Excel, LibreOffice Calc e Google Planilhas.
- Também inclui opção de alternar para `.csv` caso você precise importar em sistemas legados.
