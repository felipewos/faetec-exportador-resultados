// FAETEC - Exportador de Resultados
// Content Script para Manifest V3

(() => {
  // Evitar injeção duplicada
  if (window.__faetecExporterInjected) return;
  window.__faetecExporterInjected = true;

  const HEADERS = [
    "Cargo",
    "Disciplina",
    "Região",
    "Classificação",
    "Nº de Inscrição",
    "Nome do Candidato",
    "Titulação Pref.",
    "Cota",
    "Total de Pontos",
    "Pontos Titulação",
    "Pontos Experiência",
    "Data de Nascimento",
    "Resultados"
  ];

  let isCrawlingAllPages = false;
  let cancelCrawling = false;
  let shouldSaveOnCancel = true;

  // Obter valores selecionados dos filtros
  function getFilterValues() {
    const selects = Array.from(document.querySelectorAll("form select"));
    let cargo = "";
    let disciplina = "";
    let regiao = "";

    if (selects.length >= 1) {
      const opt = selects[0].options[selects[0].selectedIndex];
      if (opt && opt.value && !opt.text.includes("-Antes")) cargo = opt.text.trim();
    }
    if (selects.length >= 2) {
      const opt = selects[1].options[selects[1].selectedIndex];
      if (opt && opt.value && !opt.text.includes("-Antes")) disciplina = opt.text.trim();
    }
    if (selects.length >= 3) {
      const opt = selects[2].options[selects[2].selectedIndex];
      if (opt && opt.value && !opt.text.includes("-Antes")) regiao = opt.text.trim();
    }

    return { cargo, disciplina, regiao };
  }

  // Obter linhas visíveis na tabela atual
  function getCurrentPageRows() {
    const table = document.querySelector("table.tabela-resultado") || document.querySelector("table");
    if (!table) return [];

    const { cargo, disciplina, regiao } = getFilterValues();
    const trs = Array.from(table.querySelectorAll("tbody tr"));
    const dataRows = [];

    for (const tr of trs) {
      const tds = Array.from(tr.querySelectorAll("td")).map(td => td.innerText.trim());
      // Se não for linha de candidato (ex: mensagem de aviso ou vazia), pula
      if (tds.length < 5) continue;

      dataRows.push([
        cargo || "Não identificado",
        disciplina || "Não identificada",
        regiao || "Não identificada",
        tds[0] || "",
        tds[1] || "",
        tds[2] || "",
        tds[3] || "",
        tds[4] || "",
        tds[5] || "",
        tds[6] || "",
        tds[7] || "",
        tds[8] || "",
        tds[9] || ""
      ]);
    }

    return dataRows;
  }

  // Detectar informações de paginação
  function getPaginationInfo() {
    const pagination = document.querySelector(".pagination");
    if (!pagination) return { hasPagination: false, currentPage: 1, totalPages: 1 };

    const items = Array.from(pagination.querySelectorAll(".page-link")).map(el => el.innerText.trim());
    const activeItem = pagination.querySelector(".page-item.active .page-link");
    const currentPage = activeItem ? parseInt(activeItem.innerText.trim(), 10) || 1 : 1;

    // Números de páginas encontrados
    const pageNumbers = items.map(t => parseInt(t, 10)).filter(n => !isNaN(n));
    const maxPage = pageNumbers.length > 0 ? Math.max(...pageNumbers) : 1;

    return {
      hasPagination: items.length > 1,
      currentPage,
      totalPages: items.includes("Última") ? "Várias" : maxPage
    };
  }

  // Escapar valores para formato CSV
  function escapeCsvCell(value, delimiter = ";") {
    if (value === null || value === undefined) return "";
    const str = String(value).trim();
    if (str.includes(delimiter) || str.includes('"') || str.includes("\n") || str.includes("\r")) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  }

  // Gerar e fazer download do CSV com UTF-8 BOM
  function downloadCsv(rows, filename, delimiter = ";") {
    const allRows = [HEADERS, ...rows];
    const csvContent = allRows
      .map(row => row.map(cell => escapeCsvCell(cell, delimiter)).join(delimiter))
      .join("\r\n");

    const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  // Gerar e fazer download de planilha nativa XLSX (.xlsx)
  function downloadXlsx(rows, filename) {
    if (typeof XLSX === "undefined") {
      console.warn("Biblioteca XLSX não carregada no momento, utilizando fallback CSV.");
      downloadCsv(rows, filename.replace(/\.xlsx$/i, ".csv"));
      return;
    }

    const allData = [HEADERS, ...rows];
    const worksheet = XLSX.utils.aoa_to_sheet(allData);

    // Definir larguras de colunas proporcionais para visualização limpa no Excel
    const colWidths = HEADERS.map((header, colIdx) => {
      let maxLen = header.length;
      const sampleCount = Math.min(rows.length, 60);
      for (let i = 0; i < sampleCount; i++) {
        const val = rows[i][colIdx] ? String(rows[i][colIdx]) : "";
        if (val.length > maxLen) maxLen = val.length;
      }
      return { wch: Math.min(Math.max(maxLen + 3, 10), 45) };
    });
    worksheet["!cols"] = colWidths;

    // Habilitar AutoFiltro em todas as colunas (da linha 1 até o final)
    const filterRange = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(rows.length, 1), c: HEADERS.length - 1 }
    });
    worksheet["!autofilter"] = { ref: filterRange };

    // Congelar a linha de cabeçalho para fixar os títulos e filtros ao rolar a planilha
    worksheet["!views"] = [
      { state: "frozen", xSplit: 0, ySplit: 1, topLeftCell: "A2", activeCell: "A2" }
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Resultados FAETEC");

    const finalFilename = filename.endsWith(".xlsx") ? filename : filename.replace(/\.csv$/i, "") + ".xlsx";

    // Criar arquivo binário XLSX e acionar download direto
    const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = finalFilename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  // Criar nome amigável para o arquivo
  function buildFilename(cargo, disciplina, regiao, suffix = "", extension = "xlsx") {
    const sanitize = str =>
      (str || "GERAL")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-zA-Z0-9_-]/g, "_")
        .replace(/_+/g, "_")
        .substring(0, 25);

    const now = new Date();
    const dateStr = now.toLocaleDateString("pt-BR").replace(/\//g, "-");
    const timeStr = now.toTimeString().slice(0, 5).replace(/:/g, "h");

    const c = sanitize(cargo);
    const d = sanitize(disciplina);
    const r = sanitize(regiao);

    return `FAETEC_${c}_${d}_${r}${suffix ? "_" + suffix : ""}_${dateStr}_${timeStr}.${extension}`;
  }

  // Exportar registros no formato preferido (XLSX por padrão)
  function exportRows(rows, cargo, disciplina, regiao, suffix = "") {
    const selectFormat = document.getElementById("faetec-select-format");
    const format = selectFormat ? selectFormat.value : "xlsx";

    if (format === "csv") {
      const filename = buildFilename(cargo, disciplina, regiao, suffix, "csv");
      downloadCsv(rows, filename, ";");
    } else {
      const filename = buildFilename(cargo, disciplina, regiao, suffix, "xlsx");
      downloadXlsx(rows, filename);
    }
  }

  // Exibir notificação rápida (toast)
  function showToast(message, isError = false) {
    let toast = document.getElementById("faetec-toast");
    if (!toast) {
      toast = document.createElement("div");
      toast.id = "faetec-toast";
      toast.className = "faetec-toast";
      document.body.appendChild(toast);
    }

    toast.style.background = isError ? "#b91c1c" : "#1e293b";
    toast.textContent = message;
    toast.classList.add("show");

    setTimeout(() => {
      toast.classList.remove("show");
    }, 4000);
  }

  // Disparar o botão de busca oficial do portal FAETEC
  function triggerSearch() {
    const btn = document.querySelector(".btn-buscar") || document.querySelector("button[type='submit']");
    if (btn) {
      btn.click();
      return true;
    }
    return false;
  }

  // Pausa com timeout
  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  // Aguardar opções serem preenchidas ou atualizadas no select comparando com assinatura anterior
  async function waitForSelectOptions(selectEl, previousSignature = "", minOptions = 2, timeoutMs = 4000) {
    const start = Date.now();
    await sleep(20);
    while (Date.now() - start < timeoutMs) {
      const currentSignature = Array.from(selectEl.options).map(o => o.value).join(",");
      if (selectEl.options.length >= minOptions && (currentSignature !== previousSignature || previousSignature === "")) {
        return true;
      }
      await sleep(20);
    }
  }

  // Aguardar mudança de página na tabela
  async function waitForPageChange(expectedPage, prevFirstRowText = "", timeoutMs = 4000) {
    const start = Date.now();
    await sleep(40);
    while (Date.now() - start < timeoutMs) {
      const activeLink = document.querySelector(".pagination .page-item.active .page-link");
      const currentActivePage = activeLink ? parseInt(activeLink.innerText.trim(), 10) : 0;
      const currentFirstRow = document.querySelector("tbody tr td:nth-child(3)")?.innerText.trim();

      if ((currentActivePage === expectedPage || !expectedPage) && 
          (currentFirstRow !== prevFirstRowText || prevFirstRowText === "")) {
        await sleep(35);
        return true;
      }
      await sleep(20);
    }
    return false;
  }

  // Aguardar resposta do servidor Blazor após o clique em Buscar com verificação de mutação real
  async function waitForTableResults(hadDataBefore = false, timeoutMs = 4000) {
    const start = Date.now();
    const tableContainer = document.querySelector(".table-responsive") || document.querySelector("table") || document.body;

    let mutated = false;
    const observer = new MutationObserver(() => { mutated = true; });
    observer.observe(tableContainer, { childList: true, subtree: true, characterData: true });

    const maxWait = hadDataBefore ? timeoutMs : 350;

    while (Date.now() - start < maxWait) {
      if (mutated) {
        await sleep(40);
        break;
      }
      await sleep(15);
    }
    observer.disconnect();

    const tbody = document.querySelector("tbody");
    const text = tbody ? tbody.innerText.trim() : "";
    return text !== "" && !text.includes("Clique em");
  }

  // Extrair registros atribuindo explicitamente Cargo, Disciplina e Região da iteração
  function getCurrentPageRowsWithNames(cargo, disciplina, regiao) {
    const table = document.querySelector("table.tabela-resultado") || document.querySelector("table");
    if (!table) return [];

    const trs = Array.from(table.querySelectorAll("tbody tr"));
    const dataRows = [];

    for (const tr of trs) {
      const tds = Array.from(tr.querySelectorAll("td")).map(td => td.innerText.trim());
      if (tds.length < 5) continue;

      dataRows.push([
        cargo || "Não identificado",
        disciplina || "Não identificada",
        regiao || "Não identificada",
        tds[0] || "",
        tds[1] || "",
        tds[2] || "",
        tds[3] || "",
        tds[4] || "",
        tds[5] || "",
        tds[6] || "",
        tds[7] || "",
        tds[8] || "",
        tds[9] || ""
      ]);
    }

    return dataRows;
  }

  // Coletar registros da seleção atual incluindo todas as páginas da tabela
  async function collectCurrentSelectionWithPagination(cargoName, discName, regName, allCandidateRows) {
    let currentPage = 1;
    let hasNext = true;

    while (hasNext && !cancelCrawling) {
      const pageRows = getCurrentPageRowsWithNames(cargoName, discName, regName);
      for (const row of pageRows) {
        const inscricao = row[4] || Math.random().toString();
        // Chave única composta para evitar duplicatas mesmo em recarregamento
        const key = `${inscricao}_${row[0]}_${row[1]}_${row[2]}_${row[3]}`;
        allCandidateRows.set(key, row);
      }

      const pag = document.querySelector(".pagination");
      if (!pag) break;

      const nextNum = String(currentPage + 1);
      const links = Array.from(pag.querySelectorAll(".page-link"));
      let nextBtn = links.find(el => el.innerText.trim() === nextNum);

      if (!nextBtn) {
        const proxima = links.find(el => el.innerText.trim().toLowerCase().includes("próxima"));
        if (proxima && !proxima.parentElement.classList.contains("disabled")) {
          nextBtn = proxima;
        }
      }

      if (nextBtn && !nextBtn.parentElement.classList.contains("disabled")) {
        const prevFirstRow = document.querySelector("tbody tr td:nth-child(3)")?.innerText.trim() || document.querySelector("tbody tr td")?.innerText;
        nextBtn.click();
        currentPage++;
        await waitForPageChange(currentPage, prevFirstRow);
      } else {
        hasNext = false;
      }
    }
  }

  // Aguardar os dados serem carregados pelo Blazor após a busca manual
  function waitForSearchResults(timeoutMs = 8000) {
    return new Promise(resolve => {
      const start = Date.now();
      const interval = setInterval(() => {
        const rows = getCurrentPageRows();
        const tbody = document.querySelector("tbody");
        const tableText = tbody ? tbody.innerText.trim() : "";
        const searchDone = rows.length > 0 ||
          (tableText !== "" && !tableText.includes("Clique em") && !tableText.includes("Buscar")) ||
          (tbody && tbody.children.length === 0 && !tableText.includes("Clique em")) ||
          (Date.now() - start > timeoutMs);

        if (searchDone) {
          clearInterval(interval);
          setTimeout(resolve, 350);
        }
      }, 200);
    });
  }

  // Criar e injetar o painel flutuante
  function createWidget() {
    if (document.getElementById("faetec-exporter-widget")) return;

    const widget = document.createElement("div");
    widget.id = "faetec-exporter-widget";
    widget.innerHTML = `
      <div class="faetec-widget-header" id="faetec-widget-header">
        <div class="faetec-widget-title">
          <svg viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>
          Exportador FAETEC
        </div>
        <div class="faetec-widget-controls">
          <button class="faetec-btn-icon" id="faetec-btn-toggle" title="Minimizar/Restaurar">−</button>
        </div>
      </div>
      <div class="faetec-widget-body">
        <div class="faetec-filter-summary">
          <div class="faetec-filter-item"><span>Cargo:</span> <strong id="faetec-lbl-cargo">-</strong></div>
          <div class="faetec-filter-item"><span>Disciplina:</span> <strong id="faetec-lbl-disc">-</strong></div>
          <div class="faetec-filter-item"><span>Região:</span> <strong id="faetec-lbl-reg">-</strong></div>
        </div>

        <div class="faetec-data-badge empty" id="faetec-status-badge">
          <span id="faetec-status-text">Aguardando busca...</span>
          <span id="faetec-status-count">0 registros</span>
        </div>

        <div class="faetec-progress-box" id="faetec-progress-box">
          <div style="display: flex; justify-content: space-between; font-weight: 600;">
            <span id="faetec-progress-title">Exportando dados...</span>
            <span id="faetec-progress-percent">0%</span>
          </div>
          <div class="faetec-progress-bar-bg">
            <div class="faetec-progress-bar-fill" id="faetec-progress-bar"></div>
          </div>
          <div class="faetec-progress-details" id="faetec-progress-details">
            Iniciando varredura...
          </div>
          <div class="faetec-progress-actions">
            <button class="faetec-btn faetec-btn-stop" id="faetec-btn-stop-and-save" title="Parar e gerar a planilha imediatamente com tudo que já foi coletado">
              💾 Parar e Baixar Agora
            </button>
            <button class="faetec-btn faetec-btn-secondary" id="faetec-btn-cancel-crawl" title="Cancelar sem salvar">
              ✖ Cancelar
            </button>
          </div>
        </div>

        <div class="faetec-actions" id="faetec-action-buttons">
          <button class="faetec-btn faetec-btn-full" id="faetec-btn-export-everything" title="Percorre todos os 4 cargos, todas as disciplinas e todas as regiões gerando um único Excel consolidado">
            🌐 Exportar Edital Completo (Tudo)
          </button>
          <button class="faetec-btn faetec-btn-secondary" id="faetec-btn-export-cargo" title="Percorre todas as disciplinas e regiões do cargo atualmente selecionado">
            🏢 Exportar Todo este Cargo
          </button>
          <div class="faetec-section-divider"></div>
          <button class="faetec-btn faetec-btn-search" id="faetec-btn-search" style="display: none;">
            🔍 Buscar Resultados Agora
          </button>
          <button class="faetec-btn faetec-btn-primary" id="faetec-btn-export-all">
            🚀 Exportar Filtro Atual (Todas Páginas)
          </button>
          <button class="faetec-btn faetec-btn-secondary" id="faetec-btn-export-page">
            📥 Exportar Página Atual
          </button>
        </div>

        <div class="faetec-options">
          <span>Formato:</span>
          <select id="faetec-select-format">
            <option value="xlsx" selected>📊 Excel Nativo (.xlsx)</option>
            <option value="csv">📄 Arquivo (.csv)</option>
          </select>
        </div>
      </div>
    `;

    document.body.appendChild(widget);

    // Configurar eventos
    const btnToggle = widget.querySelector("#faetec-btn-toggle");
    btnToggle.addEventListener("click", () => {
      widget.classList.toggle("minimized");
      btnToggle.textContent = widget.classList.contains("minimized") ? "+" : "−";
    });

    const btnExportEverything = widget.querySelector("#faetec-btn-export-everything");
    btnExportEverything.addEventListener("click", () => handleExportEverything("ALL"));

    const btnExportCargo = widget.querySelector("#faetec-btn-export-cargo");
    btnExportCargo.addEventListener("click", () => handleExportEverything("CARGO"));

    const btnStopAndSave = widget.querySelector("#faetec-btn-stop-and-save");
    btnStopAndSave.addEventListener("click", () => {
      cancelCrawling = true;
      shouldSaveOnCancel = true;
      showToast("💾 Parando varredura e gerando planilha com os dados coletados...", false);
    });

    const btnCancel = widget.querySelector("#faetec-btn-cancel-crawl");
    btnCancel.addEventListener("click", () => {
      cancelCrawling = true;
      shouldSaveOnCancel = false;
      showToast("Varredura cancelada.", true);
    });

    const btnSearch = widget.querySelector("#faetec-btn-search");
    btnSearch.addEventListener("click", async () => {
      const { cargo, disciplina, regiao } = getFilterValues();
      if (!cargo || !disciplina || !regiao) {
        showToast("Selecione Cargo, Disciplina e Região antes de buscar.", true);
        return;
      }
      btnSearch.textContent = "⏳ Buscando no site...";
      btnSearch.disabled = true;
      triggerSearch();
      await waitForSearchResults();
      btnSearch.textContent = "🔍 Buscar Resultados Agora";
      btnSearch.disabled = false;
      updateWidgetState();
    });

    const btnExportPage = widget.querySelector("#faetec-btn-export-page");
    btnExportPage.addEventListener("click", handleExportCurrentPage);

    const btnExportAll = widget.querySelector("#faetec-btn-export-all");
    btnExportAll.addEventListener("click", handleExportAllPages);

    // Salvar e carregar preferência de formato
    const selectFormat = widget.querySelector("#faetec-select-format");
    if (chrome.storage && chrome.storage.local) {
      chrome.storage.local.get("faetec_format", res => {
        if (res.faetec_format) selectFormat.value = res.faetec_format;
      });
      selectFormat.addEventListener("change", () => {
        chrome.storage.local.set({ faetec_format: selectFormat.value });
      });
    }

    updateWidgetState();
  }

  // Atualizar informações visuais do painel
  function updateWidgetState() {
    const widget = document.getElementById("faetec-exporter-widget");
    if (!widget) return;

    const { cargo, disciplina, regiao } = getFilterValues();
    const rows = getCurrentPageRows();
    const pagInfo = getPaginationInfo();

    const lblCargo = widget.querySelector("#faetec-lbl-cargo");
    const lblDisc = widget.querySelector("#faetec-lbl-disc");
    const lblReg = widget.querySelector("#faetec-lbl-reg");
    const badge = widget.querySelector("#faetec-status-badge");
    const statusText = widget.querySelector("#faetec-status-text");
    const btnExportEverything = widget.querySelector("#faetec-btn-export-everything");
    const btnExportCargo = widget.querySelector("#faetec-btn-export-cargo");
    const btnSearch = widget.querySelector("#faetec-btn-search");
    const btnExportPage = widget.querySelector("#faetec-btn-export-page");
    const btnExportAll = widget.querySelector("#faetec-btn-export-all");

    if (btnExportEverything) btnExportEverything.disabled = isCrawlingAllPages;
    if (btnExportCargo) btnExportCargo.disabled = !cargo || isCrawlingAllPages;

    const newCargo = cargo || "Não selecionado";
    const newDisc = disciplina || "Não selecionada";
    const newReg = regiao || "Não selecionada";

    if (lblCargo && lblCargo.textContent !== newCargo) lblCargo.textContent = newCargo;
    if (lblDisc && lblDisc.textContent !== newDisc) lblDisc.textContent = newDisc;
    if (lblReg && lblReg.textContent !== newReg) lblReg.textContent = newReg;

    const hasData = rows.length > 0;
    const filtersReady = !!(cargo && disciplina && regiao);

    let newStatusText = "";
    let newStatusCount = "";

    if (hasData) {
      newStatusText = pagInfo.hasPagination
        ? `Página ${pagInfo.currentPage} de ${pagInfo.totalPages}`
        : "Dados prontos";
      newStatusCount = `${rows.length} nesta página`;

      if (badge) {
        badge.className = "faetec-data-badge";
      }
      if (btnSearch) btnSearch.style.display = "none";
      if (btnExportPage && btnExportPage.disabled) btnExportPage.disabled = false;
      if (btnExportAll && btnExportAll.disabled) btnExportAll.disabled = false;
    } else {
      if (filtersReady) {
        newStatusText = "👉 Clique em 'Buscar'";
        newStatusCount = "Pronto para buscar";

        if (badge) {
          badge.className = "faetec-data-badge action-needed";
        }
        if (btnSearch) btnSearch.style.display = "flex";
        if (btnExportAll && btnExportAll.disabled) btnExportAll.disabled = false;
        if (btnExportPage && !btnExportPage.disabled) btnExportPage.disabled = true;
      } else {
        newStatusText = "Aguardando busca...";
        newStatusCount = "0 registros";

        if (badge) {
          badge.className = "faetec-data-badge empty";
        }
        if (btnSearch) btnSearch.style.display = "none";
        if (btnExportAll && !btnExportAll.disabled) btnExportAll.disabled = true;
        if (btnExportPage && !btnExportPage.disabled) btnExportPage.disabled = true;
      }
    }

    if (statusText && statusText.textContent !== newStatusText) statusText.textContent = newStatusText;
    if (statusCount && statusCount.textContent !== newStatusCount) statusCount.textContent = newStatusCount;
  }

  // Exportar apenas a página atual
  function handleExportCurrentPage() {
    const rows = getCurrentPageRows();
    if (rows.length === 0) {
      showToast("Nenhum registro para exportar nesta página. Faça a busca primeiro.", true);
      return;
    }

    const { cargo, disciplina, regiao } = getFilterValues();
    const pagInfo = getPaginationInfo();
    const suffix = pagInfo.hasPagination ? `pag_${pagInfo.currentPage}` : "completo";

    exportRows(rows, cargo, disciplina, regiao, suffix);
    showToast(`✅ ${rows.length} candidatos exportados para Excel (.xlsx)!`);
  }

  // Exportar todas as páginas automaticamente
  async function handleExportAllPages() {
    if (isCrawlingAllPages) return;

    let initialRows = getCurrentPageRows();
    const { cargo, disciplina, regiao } = getFilterValues();

    // Se ainda não buscou mas selecionou os filtros, busca automaticamente
    if (initialRows.length === 0) {
      if (!cargo || !disciplina || !regiao) {
        showToast("Selecione Cargo, Disciplina e Região antes de exportar.", true);
        return;
      }

      showToast("🔍 Buscando candidatos no portal FAETEC...");
      triggerSearch();
      await waitForSearchResults();
      initialRows = getCurrentPageRows();

      if (initialRows.length === 0) {
        showToast("Nenhum candidato encontrado para essa pesquisa.", true);
        updateWidgetState();
        return;
      }
    }

    const selectDelim = document.getElementById("faetec-select-delimiter");
    const delimiter = selectDelim ? selectDelim.value : ";";

    const widget = document.getElementById("faetec-exporter-widget");
    const progressBox = widget.querySelector("#faetec-progress-box");
    const progressText = widget.querySelector("#faetec-progress-text");
    const progressBar = widget.querySelector("#faetec-progress-bar");
    const progressPercent = widget.querySelector("#faetec-progress-percent");
    const actionButtons = widget.querySelector("#faetec-action-buttons");

    // Verificar se há paginação
    const pagInfo = getPaginationInfo();
    if (!pagInfo.hasPagination) {
      // Página única, exporta direto
      handleExportCurrentPage();
      return;
    }

    isCrawlingAllPages = true;
    cancelCrawling = false;
    progressBox.style.display = "flex";
    actionButtons.style.display = "none";

    const allCandidateRows = new Map(); // Usar mapa por Nº de Inscrição para evitar duplicatas

    try {
      // 1. Se não estiver na página 1, navegar até a primeira página
      const pagination = document.querySelector(".pagination");
      const primeiraBtn = Array.from(pagination.querySelectorAll(".page-link")).find(
        el => el.innerText.trim() === "Primeira" || el.innerText.trim() === "1"
      );

      const activePage = pagination.querySelector(".page-item.active .page-link");
      if (activePage && activePage.innerText.trim() !== "1" && primeiraBtn) {
        progressText.textContent = "Indo para a primeira página...";
        const prevFirstRow = document.querySelector("tbody tr td:nth-child(3)")?.innerText.trim() || document.querySelector("tbody tr td")?.innerText;
        primeiraBtn.click();
        await waitForPageChange(1, prevFirstRow);
      }

      let currentPage = 1;
      let hasNextPage = true;

      while (hasNextPage && !cancelCrawling) {
        // Coletar linhas da página atual
        const pageRows = getCurrentPageRows();
        for (const row of pageRows) {
          // Coluna 4 é o Nº de Inscrição
          const inscricao = row[4] || Math.random().toString();
          allCandidateRows.set(inscricao, row);
        }

        progressText.textContent = `Página ${currentPage} coletada (${allCandidateRows.size} registros)...`;
        progressBar.style.width = "50%"; // animação de atividade

        // Procurar botão da próxima página
        const pag = document.querySelector(".pagination");
        if (!pag) break;

        const nextNum = String(currentPage + 1);
        const links = Array.from(pag.querySelectorAll(".page-link"));
        
        let nextBtn = links.find(el => el.innerText.trim() === nextNum);

        // Se não achou número exato, procurar "Próxima"
        if (!nextBtn) {
          const proxima = links.find(el => el.innerText.trim().toLowerCase().includes("próxima"));
          if (proxima && !proxima.parentElement.classList.contains("disabled")) {
            nextBtn = proxima;
          }
        }

        if (nextBtn && !nextBtn.parentElement.classList.contains("disabled")) {
          const prevFirstRow = document.querySelector("tbody tr td:nth-child(3)")?.innerText.trim() || document.querySelector("tbody tr td")?.innerText;
          nextBtn.click();
          currentPage++;

          // Aguardar atualização do Blazor
          await waitForPageChange(currentPage, prevFirstRow);
        } else {
          hasNextPage = false;
        }
      }

      if (cancelCrawling) {
        showToast("Varredura interrompida pelo usuário.", true);
      }

      const rowsToExport = Array.from(allCandidateRows.values());
      if (rowsToExport.length > 0) {
        progressBar.style.width = "100%";
        progressPercent.textContent = "100%";
        progressText.textContent = `Concluído! ${rowsToExport.length} registros.`;

        exportRows(rowsToExport, cargo, disciplina, regiao, "TODAS_PAGINAS");
        showToast(`🎉 Exportação completa! ${rowsToExport.length} candidatos salvos em Excel (.xlsx).`);
      }
    } catch (err) {
      console.error("Erro na varredura:", err);
      showToast("Erro durante a varredura automática.", true);
    } finally {
      isCrawlingAllPages = false;
      cancelCrawling = false;
      setTimeout(() => {
        progressBox.style.display = "none";
        actionButtons.style.display = "flex";
        progressBar.style.width = "0%";
        progressPercent.textContent = "0%";
        updateWidgetState();
      }, 1500);
    }
  }

  // Exportação em massa (Edital Completo ou Todo o Cargo)
  async function handleExportEverything(scope = "ALL") {
    if (isCrawlingAllPages) return;

    isCrawlingAllPages = true;
    cancelCrawling = false;
    shouldSaveOnCancel = true;

    const widget = document.getElementById("faetec-exporter-widget");
    const progressBox = widget.querySelector("#faetec-progress-box");
    const progressTitle = widget.querySelector("#faetec-progress-title");
    const progressDetails = widget.querySelector("#faetec-progress-details");
    const progressBar = widget.querySelector("#faetec-progress-bar");
    const progressPercent = widget.querySelector("#faetec-progress-percent");
    const actionButtons = widget.querySelector("#faetec-action-buttons");

    progressBox.style.display = "flex";
    actionButtons.style.display = "none";
    progressTitle.textContent = scope === "ALL" ? "Exportando Edital Completo..." : "Exportando Cargo Completo...";

    const allCandidateRows = new Map();

    const selects = document.querySelectorAll("select");
    const cargoSelect = selects[0];
    const discSelect = selects[1];
    const regSelect = selects[2];
    const btnBuscar = document.querySelector(".btn-buscar") || document.querySelector("button[type='submit']");

    try {
      let cargosToProcess = [];
      if (scope === "ALL") {
        cargosToProcess = Array.from(cargoSelect.options).filter(o => o.value && !o.text.includes("-Antes"));
      } else {
        const cur = cargoSelect.options[cargoSelect.selectedIndex];
        if (cur && cur.value && !cur.text.includes("-Antes")) {
          cargosToProcess = [cur];
        } else {
          showToast("Selecione um Cargo primeiro.", true);
          return;
        }
      }

      let totalCombinationsEstimated = scope === "ALL" ? 2800 : 1400;
      let processedCombinations = 0;

      for (let cIdx = 0; cIdx < cargosToProcess.length; cIdx++) {
        if (cancelCrawling) break;

        const cargoOpt = cargosToProcess[cIdx];
        const cargoName = cargoOpt.text.trim();

        // Mudar Cargo
        const oldDiscSig = Array.from(discSelect.options).map(o => o.value).join(",");
        cargoSelect.value = cargoOpt.value;
        cargoSelect.dispatchEvent(new Event("change", { bubbles: true }));
        await waitForSelectOptions(discSelect, oldDiscSig, 2);
        await sleep(40);

        const discsToProcess = Array.from(discSelect.options).filter(o => o.value && !o.text.includes("-Antes"));

        for (let dIdx = 0; dIdx < discsToProcess.length; dIdx++) {
          if (cancelCrawling) break;

          const discOpt = discsToProcess[dIdx];
          const discName = discOpt.text.trim();

          // Mudar Disciplina
          const oldRegSig = Array.from(regSelect.options).map(o => o.value).join(",");
          discSelect.value = discOpt.value;
          discSelect.dispatchEvent(new Event("change", { bubbles: true }));
          await waitForSelectOptions(regSelect, oldRegSig, 2);
          await sleep(40);

          const regioesToProcess = Array.from(regSelect.options).filter(o => o.value && !o.text.includes("-Antes"));

          for (let rIdx = 0; rIdx < regioesToProcess.length; rIdx++) {
            if (cancelCrawling) break;

            const regOpt = regioesToProcess[rIdx];
            const regName = regOpt.text.trim();
            processedCombinations++;

            // Atualização visual otimizada
            const pct = Math.min(Math.round((processedCombinations / totalCombinationsEstimated) * 100), 99);
            progressBar.style.width = `${pct}%`;
            progressPercent.textContent = `${pct}%`;
            progressDetails.innerHTML = `
              <strong>Cargo (${cIdx + 1}/${cargosToProcess.length}):</strong> ${cargoName}<br>
              <strong>Disciplina (${dIdx + 1}/${discsToProcess.length}):</strong> ${discName}<br>
              <strong>Região (${rIdx + 1}/${regioesToProcess.length}):</strong> ${regName}<br>
              <div style="margin-top: 4px; font-weight: bold; color: #15803d;">
                👥 ${allCandidateRows.size} candidatos acumulados na planilha
              </div>
            `;

            // Mudar Região
            const prevCand = document.querySelector("tbody tr td:nth-child(3)")?.innerText.trim() || "";
            const hadDataBefore = prevCand !== "" && !document.querySelector("tbody")?.innerText.includes("Nenhum Resultado");

            regSelect.value = regOpt.value;
            regSelect.dispatchEvent(new Event("change", { bubbles: true }));
            await sleep(80);

            // Clicar em Buscar
            if (btnBuscar) btnBuscar.click();
            await waitForTableResults(hadDataBefore, 4000);

            // Ler candidatos (com paginação)
            await collectCurrentSelectionWithPagination(cargoName, discName, regName, allCandidateRows);
          }
        }
      }

      const rowsToExport = Array.from(allCandidateRows.values());
      if (rowsToExport.length > 0 && shouldSaveOnCancel) {
        progressBar.style.width = "100%";
        progressPercent.textContent = "100%";
        progressTitle.textContent = "Gerando planilha Excel...";

        const scopeLabel = scope === "ALL" ? "EDITAL_COMPLETO" : "CARGO_COMPLETO";
        exportRows(rowsToExport, "FAETEC", scopeLabel, "GERAL", `${rowsToExport.length}_REGISTROS`);

        if (cancelCrawling) {
          showToast(`💾 Planilha salva com ${rowsToExport.length} candidatos coletados até a pausa!`);
        } else {
          showToast(`🎉 Concluído! ${rowsToExport.length} candidatos de todos os cargos/disciplinas salvos no Excel!`);
        }
      } else if (!shouldSaveOnCancel) {
        showToast("Varredura cancelada pelo usuário.", true);
      } else {
        showToast("Nenhum candidato encontrado no portal.", true);
      }
    } catch (err) {
      console.error("Erro na exportação global:", err);
      showToast("Ocorreu um erro durante a varredura.", true);
    } finally {
      isCrawlingAllPages = false;
      cancelCrawling = false;
      shouldSaveOnCancel = true;
      setTimeout(() => {
        progressBox.style.display = "none";
        actionButtons.style.display = "flex";
        progressBar.style.width = "0%";
        progressPercent.textContent = "0%";
        updateWidgetState();
      }, 2500);
    }
  }

  // Esperar o Blazor atualizar a tabela (otimizado para alta velocidade)
  function waitForPageChange(expectedPage, previousFirstCell = null) {
    return new Promise(resolve => {
      let attempts = 0;
      const interval = setInterval(() => {
        attempts++;
        const currentActive = document.querySelector(".pagination .page-item.active .page-link");
        const newFirstCell = document.querySelector("tbody tr td")?.innerText;

        const pageMatches = currentActive && parseInt(currentActive.innerText.trim(), 10) === expectedPage;
        const cellChanged = previousFirstCell && newFirstCell && newFirstCell !== previousFirstCell;

        if (pageMatches || cellChanged || attempts > 50) {
          clearInterval(interval);
          // Folga mínima ultra ágil para Blazor Server
          setTimeout(resolve, 35);
        }
      }, 20);
    });
  }

  // Monitorar alterações no DOM sem travar a página
  let updateDebounceTimer = null;

  function scheduleUpdate(delay = 300) {
    if (isCrawlingAllPages) return;
    if (updateDebounceTimer) clearTimeout(updateDebounceTimer);
    updateDebounceTimer = setTimeout(() => {
      updateWidgetState();
    }, delay);
  }

  function initObservers() {
    createWidget();

    // Observador protegido contra loops infinitos
    const observer = new MutationObserver(mutations => {
      // Ignorar se a mutação veio do próprio widget da extensão
      const isWidgetMutation = mutations.every(m => {
        const target = m.target;
        return (
          target &&
          (target.id === "faetec-exporter-widget" ||
            target.id === "faetec-toast" ||
            (target.closest && (target.closest("#faetec-exporter-widget") || target.closest("#faetec-toast"))))
        );
      });

      if (isWidgetMutation) return;

      // Se o Blazor removeu o widget em algum momento, recriar
      if (!document.getElementById("faetec-exporter-widget")) {
        createWidget();
      }

      // Agendar atualização com debounce de 350ms para não engasgar o Blazor
      scheduleUpdate(350);
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true
    });

    // Escutar eventos normais de interação com o formulário
    document.addEventListener("change", e => {
      if (e.target && e.target.matches("select")) {
        scheduleUpdate(150);
      }
    });

    document.addEventListener("click", e => {
      if (e.target && (e.target.matches("button") || e.target.closest("button") || e.target.matches(".page-link"))) {
        scheduleUpdate(500);
      }
    });
  }

  // Responder a mensagens do popup
  if (chrome.runtime && chrome.runtime.onMessage) {
    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
      if (message.action === "GET_STATUS") {
        const filters = getFilterValues();
        const rows = getCurrentPageRows();
        const pag = getPaginationInfo();
        sendResponse({
          ready: true,
          filters,
          rowCount: rows.length,
          pagination: pag,
          isCrawling: isCrawlingAllPages
        });
        return false;
      }

      if (message.action === "EXPORT_CURRENT_PAGE") {
        handleExportCurrentPage();
        sendResponse({ success: true });
        return false;
      }

      if (message.action === "EXPORT_ALL_PAGES") {
        handleExportAllPages();
        sendResponse({ success: true });
        return false;
      }

      if (message.action === "EXPORT_EVERYTHING") {
        handleExportEverything("ALL");
        sendResponse({ success: true });
        return false;
      }

      if (message.action === "EXPORT_CARGO") {
        handleExportEverything("CARGO");
        sendResponse({ success: true });
        return false;
      }
    });
  }

  // Iniciar
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initObservers);
  } else {
    initObservers();
  }
})();
