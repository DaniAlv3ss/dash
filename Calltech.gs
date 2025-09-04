/**
 * Contém todas as funções do lado do servidor para a página do Dashboard de Calltech.
 */


/**
 * Busca e consolida os dados de um cliente de múltiplas fontes para o histórico.
 */
function getCustomerData(filter) {
  try {
    const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
    const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);
    const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);
    const abaManager = planilhaCalltech.getSheetByName(NOME_ABA_MANAGER);
    const abaAtendimento = planilhaCalltech.getSheetByName(NOME_ABA_ATENDIMENTO);

    if (!abaNPS || !abaManager || !abaAtendimento) throw new Error("Uma ou mais abas necessárias não foram encontradas.");

    const npsData = abaNPS.getRange(2, 1, abaNPS.getLastRow() - 1, abaNPS.getLastColumn()).getDisplayValues();
    const managerData = abaManager.getRange(2, 1, abaManager.getLastRow() - 1, abaManager.getLastColumn()).getDisplayValues();
    const atendimentoData = abaAtendimento.getRange(2, 1, abaAtendimento.getLastRow() - 1, abaAtendimento.getLastColumn()).getDisplayValues();

    const customers = new Map();
    const orderToEmailMap = new Map();

    // 1. Processa Pedidos Manager para criar base de clientes e mapear pedidos
    managerData.forEach(row => {
      const email = row[INDICES_CALLTECH.EMAIL]?.toString().trim().toLowerCase();
      const name = row[INDICES_CALLTECH.CLIENTE]?.toString().trim();
      const pedidoId = row[INDICES_CALLTECH.PEDIDO_ID]?.toString().trim();

      if (email && name) {
        if (!customers.has(email)) {
          customers.set(email, { name, email, history: [] });
        }
        if (pedidoId) {
          orderToEmailMap.set(pedidoId, email);
        }
      }
    });
    
    // 2. Processa NPS para complementar base e mapeamento
    npsData.forEach(row => {
      const email = row[INDICES_NPS.EMAIL]?.toString().trim().toLowerCase();
      const name = row[INDICES_NPS.CLIENTE]?.toString().trim();
      const pedidoId = row[INDICES_NPS.PEDIDO_ID]?.toString().trim();

      if (email && name) {
        if (!customers.has(email)) {
          customers.set(email, { name, email, history: [] });
        }
        if (pedidoId) {
          orderToEmailMap.set(pedidoId, email);
        }
      }
    });
    
    // 3. Filtra clientes com base no critério (letra ou pesquisa)
    let filteredEmails = new Set();
    const filterValue = filter.value.toLowerCase();

    if (filter.type === 'search' && filterValue.length >= 3) {
      customers.forEach((customer, email) => {
        if (customer.name.toLowerCase().includes(filterValue) || email.includes(filterValue)) {
          filteredEmails.add(email);
        }
      });
      orderToEmailMap.forEach((email, pedidoId) => {
        if(pedidoId.includes(filterValue) && customers.has(email)) {
           filteredEmails.add(email);
        }
      });
    }
    if(filteredEmails.size === 0) return [];

    // 4. Monta o histórico APENAS para os clientes filtrados
    const results = new Map();
    filteredEmails.forEach(email => { results.set(email, customers.get(email)); });

    npsData.forEach(row => {
      const email = row[INDICES_NPS.EMAIL]?.toString().trim().toLowerCase();
      const dateStr = row[INDICES_NPS.DATA_AVALIACAO];
      if(results.has(email) && dateStr){
         results.get(email).history.push({
           type: 'NPS', date: new Date(dateStr.split(' ')[0]), pedidoId: row[INDICES_NPS.PEDIDO_ID],
           classificacao: row[INDICES_NPS.CLASSIFICACAO], comentario: row[INDICES_NPS.COMENTARIO]
         });
      }
    });

    managerData.forEach(row => {
      const email = row[INDICES_CALLTECH.EMAIL]?.toString().trim().toLowerCase();
      const dateStr = row[INDICES_CALLTECH.DATA_ABERTURA];
      if(results.has(email) && dateStr){
         results.get(email).history.push({
           type: 'Chamado', date: new Date(dateStr.split(' ')[0].split('/').reverse().join('-')),
           chamadoId: row[INDICES_CALLTECH.CHAMADO_ID], pedidoId: row[INDICES_CALLTECH.PEDIDO_ID],
           status: row[INDICES_CALLTECH.STATUS], dataFinalizacao: row[INDICES_CALLTECH.DATA_FINALIZACAO]
         });
      }
    });

    atendimentoData.forEach(row => {
      const pedidoId = row[INDICES_ATENDIMENTO.PEDIDO_ID]?.toString().trim();
      const email = orderToEmailMap.get(pedidoId);
      const dateStr = row[INDICES_ATENDIMENTO.DATA_ATENDIMENTO];
      if(email && results.has(email) && dateStr){
         results.get(email).history.push({
           type: 'Atendimento', date: new Date(dateStr.split(' ')[0].split('/').reverse().join('-')),
           pedidoId: pedidoId, resolucao: row[INDICES_ATENDIMENTO.RESOLUCAO],
           defeito: row[INDICES_ATENDIMENTO.DEFEITO], relato: row[INDICES_ATENDIMENTO.RELATO_CLIENTE]
         });
      }
    });

    // 5. Finaliza e formata o retorno
    const finalResults = Array.from(results.values());
    finalResults.forEach(customer => {
      customer.history.sort((a, b) => b.date - a.date);
      customer.history.forEach(item => {
        item.date = (item.date instanceof Date && !isNaN(item.date)) ? item.date.toISOString() : null;
      });
    });
    return finalResults.sort((a,b) => a.name.localeCompare(b.name));

  } catch (e) {
    Logger.log(`Erro em getCustomerData: ${e.stack}`);
    return { error: `Erro ao buscar dados dos clientes: ${e.message}` };
  }
}


/**
 * Busca e processa os dados para os KPIs e tabela da página Calltech.
 */
function getCalltechData(dateRange) {
  try {
    const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);
    const abaManager = planilhaCalltech.getSheetByName(NOME_ABA_MANAGER);
    const abaAtendimento = planilhaCalltech.getSheetByName(NOME_ABA_ATENDIMENTO);
    const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
    const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);

    if (!abaManager || !abaAtendimento || !abaNPS) throw new Error(`Uma ou mais abas necessárias não foram encontradas.`);

    const allManagerData = abaManager.getRange(2, 1, abaManager.getLastRow() - 1, abaManager.getLastColumn()).getDisplayValues();
    const allNPSData = abaNPS.getRange(2, 1, abaNPS.getLastRow() - 1, abaNPS.getLastColumn()).getDisplayValues();
    
    // Cria um mapa de Pedido ID para classificação NPS para consulta rápida
    const npsMap = new Map();
    getUniqueValidRows(allNPSData, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO)
      .forEach(row => {
        const pedidoId = row[INDICES_NPS.PEDIDO_ID]?.trim();
        const classification = row[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
        if (pedidoId && classification) {
            npsMap.set(pedidoId, classification);
        }
    });
    
    let openTickets = 0, closedTickets = 0, totalResolutionTime = 0, resolvedTicketsCount = 0;
    const resolutionCounts = { day0: 0, day1: 0, day2: 0, day3: 0, day4plus: 0 };
    const tickets = [];
    const npsFeedback = { total: 0, promoters: 0, neutrals: 0, detractors: 0 };
    const postServiceNps = { total: 0, promoters: 0, neutrals: 0, detractors: 0 };
    const ticketPedidoIds = new Set(), closedTicketPedidoIds = new Set();

    allManagerData.forEach((row, index) => {
      const openDateStr = row[INDICES_CALLTECH.DATA_ABERTURA];
      const chamadoId = row[INDICES_CALLTECH.CHAMADO_ID];
      if (!chamadoId || !openDateStr) return;

      try {
        const [day, month, year] = openDateStr.split(' ')[0].split('/');
        const openDate = new Date(year, month - 1, day);
        if (isNaN(openDate.getTime())) return;

        const openDateISO = openDate.toISOString().split('T')[0];
        if (openDateISO >= dateRange.start && openDateISO <= dateRange.end) {
          const status = row[INDICES_CALLTECH.STATUS] || "";
          const closeDateStr = row[INDICES_CALLTECH.DATA_FINALIZACAO];
          const isClosed = status.toLowerCase().includes('finalizado') || status.toLowerCase().includes('resolvido');
          const pedidoId = row[INDICES_CALLTECH.PEDIDO_ID]?.trim();

          if (isClosed) {
            closedTickets++;
            if (closeDateStr) {
              const [closeDay, closeMonth, closeYear] = closeDateStr.split(' ')[0].split('/');
              const closeDate = new Date(closeYear, closeMonth - 1, closeDay);
              if (!isNaN(closeDate.getTime())) {
                const diffDays = Math.ceil(Math.abs(closeDate - openDate) / (1000 * 60 * 60 * 24));
                totalResolutionTime += diffDays;
                resolvedTicketsCount++;
                if (diffDays === 0) resolutionCounts.day0++;
                else if (diffDays === 1) resolutionCounts.day1++;
                else if (diffDays === 2) resolutionCounts.day2++;
                else if (diffDays === 3) resolutionCounts.day3++;
                else resolutionCounts.day4plus++;
              }
            }
          } else {
            openTickets++;
          }

          tickets.push({
            status, chamadoId, pedidoId,
            dataAbertura: openDateStr.split(' ')[0],
            cliente: row[INDICES_CALLTECH.CLIENTE],
            dataFinalizacao: closeDateStr ? closeDateStr.split(' ')[0] : '',
            email: row[INDICES_CALLTECH.EMAIL],
            hasNps: npsMap.has(pedidoId)
          });

          // Processa NPS geral (baseado em chamados ABERTOS no período)
          if (pedidoId && !ticketPedidoIds.has(pedidoId)) {
              if (npsMap.has(pedidoId)) {
                  const c = npsMap.get(pedidoId);
                  npsFeedback.total++;
                  if (c === 'promotor') npsFeedback.promoters++; else if (c === 'neutro') npsFeedback.neutrals++; else if (c === 'detrator') npsFeedback.detractors++;
              }
              ticketPedidoIds.add(pedidoId);
          }

          // Processa NPS Pós-Atendimento (apenas para chamados FINALIZADOS)
          if (isClosed && pedidoId && !closedTicketPedidoIds.has(pedidoId)) {
            if (npsMap.has(pedidoId)) {
              const c = npsMap.get(pedidoId);
              postServiceNps.total++;
              if (c === 'promotor') postServiceNps.promoters++; else if (c === 'neutro') postServiceNps.neutrals++; else if (c === 'detrator') postServiceNps.detractors++;
            }
            closedTicketPedidoIds.add(pedidoId);
          }
        }
      } catch (e) {
        Logger.log(`Erro ao processar linha ${index + 2} da planilha Manager: ${e.message}`);
      }
    });

    const avgResolutionTime = resolvedTicketsCount > 0 ? (totalResolutionTime / resolvedTicketsCount).toFixed(1) : 0;
    const resolutionRate = {};
    for (const day in resolutionCounts) {
      const count = resolutionCounts[day];
      resolutionRate[day] = { count, percentage: resolvedTicketsCount > 0 ? (count / resolvedTicketsCount) * 100 : 0 };
    }

    postServiceNps.npsScore = postServiceNps.total > 0 ? parseFloat((((postServiceNps.promoters - postServiceNps.detractors) / postServiceNps.total) * 100).toFixed(1)) : 0;

    // Processamento para KPI de Retenção
    const allAtendimentoData = abaAtendimento.getRange(2, 1, abaAtendimento.getLastRow() - 1, abaAtendimento.getLastColumn()).getDisplayValues();
    let retentionValue = 0;
    allAtendimentoData.forEach(row => {
      const dateStr = row[INDICES_ATENDIMENTO.DATA_ATENDIMENTO];
      if (!dateStr) return;
      try {
        const [day, month, year] = dateStr.split(' ')[0].split('/');
        const atendimentoDate = new Date(year, month - 1, day);
        const atendimentoDateISO = atendimentoDate.toISOString().split('T')[0];
        if (atendimentoDateISO >= dateRange.start && atendimentoDateISO <= dateRange.end) {
          if (row[INDICES_ATENDIMENTO.STATUS_ATENDIMENTO] === "Retido no Atendimento (MSPC)") {
            const valorNumerico = parseFloat((row[INDICES_ATENDIMENTO.VALOR_RETIDO] || '0').replace('R$', '').replace(/\./g, '').replace(',', '.').trim());
            if (!isNaN(valorNumerico)) retentionValue += valorNumerico;
          }
        }
      } catch (e) { /* Ignora linhas com datas inválidas */ }
    });

    return {
      tickets: tickets,
      kpis: { total: tickets.length, open: openTickets, closed: closedTickets, avgTime: avgResolutionTime, retentionValue, npsFeedback, postServiceNps },
      resolutionRate
    };
  } catch (e) {
    Logger.log(`Erro fatal em getCalltechData: ${e.stack}`);
    return { tickets: [], kpis: {}, resolutionRate: {} };
  }
}

/**
 * Busca dados para o gráfico de fluxo diário de chamados (abertos vs. fechados).
 */
function getDailyFlowChartData(dateRange) {
  try {
    const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);
    const abaManager = planilhaCalltech.getSheetByName(NOME_ABA_MANAGER);
    if (!abaManager) throw new Error(`Aba ${NOME_ABA_MANAGER} não encontrada.`);

    const allManagerData = abaManager.getRange(2, 1, abaManager.getLastRow() - 1, abaManager.getLastColumn()).getDisplayValues();
    const dailyFlow = {};
    const startDate = new Date(dateRange.start.replace(/-/g, '/'));
    const endDate = new Date(dateRange.end.replace(/-/g, '/'));

    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        dailyFlow[d.toISOString().split('T')[0]] = { opened: 0, closed: 0 };
    }

    allManagerData.forEach(row => {
        try {
            // Abertos
            const openDateStr = row[INDICES_CALLTECH.DATA_ABERTURA];
            if (openDateStr) {
                const [day, month, year] = openDateStr.split(' ')[0].split('/');
                const openDate = new Date(year, month - 1, day);
                const openDateISO = openDate.toISOString().split('T')[0];
                if (dailyFlow.hasOwnProperty(openDateISO)) dailyFlow[openDateISO].opened++;
            }
            // Fechados
            const closeDateStr = row[INDICES_CALLTECH.DATA_FINALIZACAO];
            if (closeDateStr) {
                const [day, month, year] = closeDateStr.split(' ')[0].split('/');
                const closeDate = new Date(year, month - 1, day);
                const closeDateISO = closeDate.toISOString().split('T')[0];
                if (dailyFlow.hasOwnProperty(closeDateISO)) dailyFlow[closeDateISO].closed++;
            }
        } catch (e) { /* Ignora datas inválidas */ }
    });

    const weekDayInitials = ['D', 'S', 'T', 'Q', 'Q', 'S', 'S'];
    const dailyFlowChartData = [['Dia', 'Abertos', 'Fechados']];
    Object.keys(dailyFlow).sort().forEach(dayISO => {
        const date = new Date(dayISO.replace(/-/g, '/'));
        const [year, month, day] = dayISO.split('-');
        const label = `${day}/${month}\n(${weekDayInitials[date.getUTCDay()]})`;
        dailyFlowChartData.push([label, dailyFlow[dayISO].opened, dailyFlow[dayISO].closed]);
    });
    return dailyFlowChartData;
  } catch (e) {
    Logger.log(`Erro fatal em getDailyFlowChartData: ${e.stack}`);
    return [['Dia', 'Abertos', 'Fechados']];
  }
}
