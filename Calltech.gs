/**
 * Contém todas as funções do lado do servidor para a página do Dashboard de Calltech.
 * OTIMIZADO: A função principal agora pré-carrega os históricos dos clientes relevantes.
 */
function getCalltechData(dateRange) {
  try {
    const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);
    const abaManager = planilhaCalltech.getSheetByName(NOME_ABA_MANAGER);
    const abaAtendimento = planilhaCalltech.getSheetByName(NOME_ABA_ATENDIMENTO);
    const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
    const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);
    const planilhaDevolucao = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilhaDevolucao.getSheetByName(NOME_ABA_DEVOLUCAO);

    if (!abaManager || !abaAtendimento || !abaNPS || !abaDevolucao) throw new Error(`Uma ou mais abas necessárias não foram encontradas.`);

    // Carrega todos os dados necessários uma única vez para otimização
    const allManagerData = abaManager.getRange(2, 1, abaManager.getLastRow() - 1, abaManager.getLastColumn()).getDisplayValues();
    const allNPSData = abaNPS.getRange(2, 1, abaNPS.getLastRow() - 1, abaNPS.getLastColumn()).getDisplayValues();
    const allAtendimentoData = abaAtendimento.getRange(2, 1, abaAtendimento.getLastRow() - 1, abaAtendimento.getLastColumn()).getDisplayValues();
    const allDevolucaoData = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, abaDevolucao.getLastColumn()).getValues();
    
    // ==========================================================
    // PARTE 1: Processamento de KPIs e Tabela de Chamados
    // ==========================================================
    const npsMap = new Map();
    getUniqueValidRows(allNPSData, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO)
      .forEach(row => {
        const pedidoId = row[INDICES_NPS.PEDIDO_ID]?.trim();
        const classification = row[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
        if (pedidoId && classification) {
            npsMap.set(pedidoId, classification);
        }
    });

    const devolucaoMap = new Map();
    allDevolucaoData.forEach(row => {
        const pedidoId = row[0]?.toString().trim();
        if(pedidoId) {
            devolucaoMap.set(pedidoId, true);
        }
    });
    
    let openTickets = 0, closedTickets = 0, totalResolutionTime = 0, resolvedTicketsCount = 0;
    const resolutionCounts = { day0: 0, day1: 0, day2: 0, day3: 0, day4plus: 0 };
    const tickets = [];
    const npsFeedback = { total: 0, promoters: 0, neutrals: 0, detractors: 0 };
    const postServiceNps = { total: 0, promoters: 0, neutrals: 0, detractors: 0 };
    const ticketPedidoIds = new Set(), closedTicketPedidoIds = new Set();
    const customerEmailsInPeriod = new Set(); // Coleta e-mails para pré-carregamento do histórico
    const allTicketStatuses = new Set();


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
          const status = row[INDICES_CALLTECH.STATUS] || "Sem Status";
          allTicketStatuses.add(status);

          const closeDateStr = row[INDICES_CALLTECH.DATA_FINALIZACAO];
          const isClosed = status.toLowerCase().includes('finalizado') || status.toLowerCase().includes('resolvido');
          const pedidoId = row[INDICES_CALLTECH.PEDIDO_ID]?.trim();
          const email = row[INDICES_CALLTECH.EMAIL]?.toString().trim().toLowerCase();

          if (email) { customerEmailsInPeriod.add(email); }

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
            hasNps: npsMap.has(pedidoId),
            hasDevolucao: devolucaoMap.has(pedidoId)
          });

          if (pedidoId && !ticketPedidoIds.has(pedidoId)) {
              if (npsMap.has(pedidoId)) {
                  const c = npsMap.get(pedidoId);
                  npsFeedback.total++;
                  if (c === 'promotor') npsFeedback.promoters++; else if (c === 'neutro') npsFeedback.neutrals++; else if (c === 'detrator') npsFeedback.detractors++;
              }
              ticketPedidoIds.add(pedidoId);
          }

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

    // ==========================================================
    // PARTE 2: Pré-carregamento dos históricos de clientes
    // ==========================================================
    const customerHistories = {};
    if (customerEmailsInPeriod.size > 0) {
        
        customerEmailsInPeriod.forEach(email => {
            customerHistories[email] = { name: '', email: email, history: [] };
        });

        const orderToEmailMap = new Map();
        allManagerData.forEach(row => {
            const email = row[INDICES_CALLTECH.EMAIL]?.toString().trim().toLowerCase();
            const pedidoId = row[INDICES_CALLTECH.PEDIDO_ID]?.toString().trim();
            const name = row[INDICES_CALLTECH.CLIENTE]?.toString().trim();
            if (email && pedidoId) orderToEmailMap.set(pedidoId, { email, name });
        });
        allNPSData.forEach(row => {
            const email = row[INDICES_NPS.EMAIL]?.toString().trim().toLowerCase();
            const pedidoId = row[INDICES_NPS.PEDIDO_ID]?.toString().trim();
            const name = row[INDICES_NPS.CLIENTE]?.toString().trim();
            if (email && pedidoId && !orderToEmailMap.has(pedidoId)) orderToEmailMap.set(pedidoId, { email, name });
        });

        allNPSData.forEach(row => {
            const email = row[INDICES_NPS.EMAIL]?.toString().trim().toLowerCase();
            if (customerHistories[email]) {
                const dateStr = row[INDICES_NPS.DATA_AVALIACAO];
                if (dateStr) {
                    customerHistories[email].name = customerHistories[email].name || row[INDICES_NPS.CLIENTE]?.toString().trim();
                    customerHistories[email].history.push({
                      type: 'NPS', date: new Date(dateStr.split(' ')[0]), pedidoId: row[INDICES_NPS.PEDIDO_ID],
                      classificacao: row[INDICES_NPS.CLASSIFICACAO], comentario: row[INDICES_NPS.COMENTARIO]
                    });
                }
            }
        });

        allManagerData.forEach(row => {
            const email = row[INDICES_CALLTECH.EMAIL]?.toString().trim().toLowerCase();
            if (customerHistories[email]) {
                const dateStr = row[INDICES_CALLTECH.DATA_ABERTURA];
                if (dateStr) {
                    customerHistories[email].name = customerHistories[email].name || row[INDICES_CALLTECH.CLIENTE]?.toString().trim();
                    customerHistories[email].history.push({
                      type: 'Chamado', date: new Date(dateStr.split(' ')[0].split('/').reverse().join('-')),
                      chamadoId: row[INDICES_CALLTECH.CHAMADO_ID], pedidoId: row[INDICES_CALLTECH.PEDIDO_ID],
                      status: row[INDICES_CALLTECH.STATUS], dataFinalizacao: row[INDICES_CALLTECH.DATA_FINALIZACAO]
                    });
                }
            }
        });

        allAtendimentoData.forEach(row => {
            const pedidoId = row[INDICES_ATENDIMENTO.PEDIDO_ID]?.toString().trim();
            const mapping = orderToEmailMap.get(pedidoId);
            if (mapping && customerHistories[mapping.email]) {
                const dateStr = row[INDICES_ATENDIMENTO.DATA_ATENDIMENTO];
                if (dateStr) {
                    customerHistories[mapping.email].name = customerHistories[mapping.email].name || mapping.name;
                    customerHistories[mapping.email].history.push({
                      type: 'Atendimento', date: new Date(dateStr.split(' ')[0].split('/').reverse().join('-')),
                      pedidoId: pedidoId, resolucao: row[INDICES_ATENDIMENTO.RESOLUCAO],
                      defeito: row[INDICES_ATENDIMENTO.DEFEITO], relato: row[INDICES_ATENDIMENTO.RELATO_CLIENTE]
                    });
                }
            }
        });

        const INDICES_DEVOLUCAO = {
          PEDIDO_ID: 0,
          DATA_NFE: 3,
          PRODUTO: 10,
          VALOR_DEVOLUCAO: 24,
          MOTIVO: 26,
        };

        const devolucoesAgrupadas = {};
        allDevolucaoData.forEach(row => {
          const pedidoId = row[INDICES_DEVOLUCAO.PEDIDO_ID]?.toString().trim();
          if (!pedidoId) return;

          const mapping = orderToEmailMap.get(pedidoId);
          if (mapping && customerHistories[mapping.email]) {
              const dateValue = row[INDICES_DEVOLUCAO.DATA_NFE];
              if (dateValue instanceof Date) {
                  const key = pedidoId;
                  if (!devolucoesAgrupadas[key]) {
                      devolucoesAgrupadas[key] = {
                          type: 'DevolucaoAgrupada',
                          pedidoId: pedidoId,
                          date: dateValue,
                          email: mapping.email,
                          name: mapping.name,
                          items: []
                      };
                  }
                  devolucoesAgrupadas[key].items.push({
                      produto: row[INDICES_DEVOLUCAO.PRODUTO],
                      motivo: row[INDICES_DEVOLUCAO.MOTIVO],
                      valor: row[INDICES_DEVOLUCAO.VALOR_DEVOLUCAO]
                  });
              }
          }
        });

        Object.values(devolucoesAgrupadas).forEach(devolucao => {
            const email = devolucao.email;
            if (customerHistories[email]) {
                devolucao.totalItens = devolucao.items.length;
                devolucao.valorTotal = devolucao.items.reduce((sum, item) => {
                    const valorRaw = item.valor;
                    const valorNum = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);
                    return sum + valorNum;
                }, 0);

                customerHistories[email].name = customerHistories[email].name || devolucao.name;
                customerHistories[email].history.push(devolucao);
            }
        });
        
        Object.values(customerHistories).forEach(customer => {
            customer.history.sort((a, b) => b.date - a.date);
            customer.history.forEach(item => {
                item.date = (item.date instanceof Date && !isNaN(item.date)) ? item.date.toISOString() : null;
            });
        });
    }

    return {
      tickets: tickets,
      kpis: { total: tickets.length, open: openTickets, closed: closedTickets, avgTime: avgResolutionTime, retentionValue, npsFeedback, postServiceNps },
      resolutionRate,
      customerHistories, // Retorna o objeto de históricos pré-carregado
      uniqueStatuses: Array.from(allTicketStatuses).sort()
    };
  } catch (e) {
    Logger.log(`Erro fatal em getCalltechData: ${e.stack}`);
    return { tickets: [], kpis: {}, resolutionRate: {}, customerHistories: {}, uniqueStatuses: [] };
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

// --- VERSÕES COM CACHE PARA SEREM CHAMADAS PELO CLIENTE ---

function getCalltechDataWithCache(dateRange) {
  const cacheKey = `calltech_data_v4_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getCalltechData, [dateRange]);
}

function getDailyFlowChartDataWithCache(dateRange) {
  const cacheKey = `daily_flow_chart_data_v2_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getDailyFlowChartData, [dateRange]);
}
