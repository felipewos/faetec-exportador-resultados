// Script do Popup da Extensão FAETEC

document.addEventListener("DOMContentLoaded", async () => {
  const viewConnected = document.getElementById("view-connected");
  const viewDisconnected = document.getElementById("view-disconnected");

  const valCargo = document.getElementById("val-cargo");
  const valDisc = document.getElementById("val-disc");
  const valReg = document.getElementById("val-reg");
  const numRows = document.getElementById("num-rows");
  const numPages = document.getElementById("num-pages");

  const btnExportAll = document.getElementById("btn-export-all");
  const btnExportPage = document.getElementById("btn-export-page");
  const btnOpenFaetec = document.getElementById("btn-open-faetec");

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    const isFaetec = tab && tab.url && tab.url.includes("processoseletivo.faetec.rj.gov.br");

    if (!isFaetec) {
      viewDisconnected.classList.remove("hidden");
      viewConnected.classList.add("hidden");

      if (btnOpenFaetec) {
        btnOpenFaetec.addEventListener("click", () => {
          chrome.tabs.create({ url: "https://processoseletivo.faetec.rj.gov.br/ResultadoTitulacaoExperienciaCotaPne/60" });
          window.close();
        });
      }
      return;
    }

    // Está na página da FAETEC
    viewConnected.classList.remove("hidden");
    viewDisconnected.classList.add("hidden");

    // Consultar dados do Content Script
    try {
      const response = await chrome.tabs.sendMessage(tab.id, { action: "GET_STATUS" });

      if (response && response.ready) {
        const { filters, rowCount, pagination, isCrawling } = response;

        valCargo.textContent = filters.cargo || "Não selecionado";
        valDisc.textContent = filters.disciplina || "Não selecionada";
        valReg.textContent = filters.regiao || "Não selecionada";

        numRows.textContent = rowCount;
        numPages.textContent = pagination.hasPagination ? `${pagination.currentPage} de ${pagination.totalPages}` : "1 (Única)";

        const hasData = rowCount > 0;
        const filtersReady = !!(filters.cargo && filters.disciplina && filters.regiao);
        btnExportAll.disabled = (!hasData && !filtersReady) || isCrawling;
        btnExportPage.disabled = !hasData || isCrawling;

        if (isCrawling) {
          btnExportAll.textContent = "⏳ Varredura em andamento...";
        }
      }
    } catch (err) {
      console.warn("Content script ainda não respondeu, recarregando estado:", err);
      valCargo.textContent = "Atualize a página";
    }

    // Ações dos botões
    btnExportPage.addEventListener("click", async () => {
      btnExportPage.disabled = true;
      btnExportPage.textContent = "Baixando...";
      await chrome.tabs.sendMessage(tab.id, { action: "EXPORT_CURRENT_PAGE" });
      setTimeout(() => {
        btnExportPage.disabled = false;
        btnExportPage.textContent = "📥 Exportar Página Atual";
      }, 1000);
    });

    const btnExportEverything = document.getElementById("btn-export-everything");
    if (btnExportEverything) {
      btnExportEverything.addEventListener("click", async () => {
        btnExportEverything.disabled = true;
        btnExportEverything.textContent = "Iniciando varredura completa...";
        await chrome.tabs.sendMessage(tab.id, { action: "EXPORT_EVERYTHING" });
        window.close();
      });
    }

    btnExportAll.addEventListener("click", async () => {
      btnExportAll.disabled = true;
      btnExportAll.textContent = "Iniciando varredura...";
      await chrome.tabs.sendMessage(tab.id, { action: "EXPORT_ALL_PAGES" });
      window.close(); // Fecha o popup para o usuário acompanhar no painel da própria página
    });

  } catch (error) {
    console.error("Erro ao inicializar popup:", error);
  }
});
