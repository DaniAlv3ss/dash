/**
 * Script para servir um dashboard de NPS como um aplicativo da web.
 * VERSÃO REATORADA: Puxa dados de planilhas separadas por ID e carrega páginas dinamicamente.
 */

// === CENTRAL DE IDs DAS PLANILHAS ===
const ID_PLANILHA_NPS = "1ewRARy4u4V0MJMoup0XbPlLLUrdPmR4EZwRwmy_ZECM";
const ID_PLANILHA_CALLTECH = "1bmHgGpAXAB4Sh95t7drXLImfNgAojCHv-o2CYS2d3-g";
// ===================================

// Constantes para nomes de abas
const NOME_ABA_NPS = "Avaliações 2025";
const NOME_ABA_ACOES = "ações 2025"; // NOVA ABA
const NOME_ABA_ATENDIMENTO = "Forms";
const NOME_ABA_OS = "NPS Datas";
const NOME_ABA_MANAGER = "Pedidos Manager"; // Aba para a nova página Calltech

// Índices de colunas da aba de NPS
const INDICES_NPS = {
  DATA_AVALIACAO: 2,  // Coluna C
  CLASSIFICACAO: 5,   // Coluna F
  CLIENTE: 12,        // Coluna M
  EMAIL: 15,          // Coluna P
  PEDIDO_ID: 32,      // Coluna AG
  COMENTARIO: 48,     // Coluna AW
  MOTIVO_FUNCIONAMENTO: 39,
  MOTIVO_QUALIDADE_MONTAGEM: 45,
  MOTIVO_VISUAL_PC: 43,
  MOTIVO_TRANSPORTE: 47
};

// Índices de colunas da aba de Atendimento (Forms)
const INDICES_ATENDIMENTO = {
  PEDIDO_ID: 2,
  RESOLUCAO: 4,       // Coluna E
  STATUS_ATENDIMENTO: 5, // Coluna F
  VALOR_RETIDO: 6,       // Coluna G
  DEFEITO: 7,         // Coluna H
  RELATO_CLIENTE: 8,  // Coluna I
  DATA_ATENDIMENTO: 0,
  OS: 14
};

// Índices de colunas da aba NPS Datas
const INDICES_OS = {
  PEDIDO_ID: 2,
  OS: 3         // Coluna D
};

// NOVOS ÍNDICES: Colunas para a aba Pedidos Manager (Calltech)
const INDICES_CALLTECH = {
  EMAIL: 0,               // Coluna A
  STATUS: 2,              // Coluna C
  CHAMADO_ID: 3,          // Coluna D
  DATA_ABERTURA: 6,       // Coluna G
  PEDIDO_ID: 12,            // Coluna M
  CLIENTE: 14,            // Coluna O
  DATA_FINALIZACAO: 16    // Coluna Q
};


/**
 * Função principal que serve o "casco" da aplicação (menu e área de conteúdo).
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard KaBuM! - Monte o Seu PC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Retorna o conteúdo HTML de uma página específica para ser carregado dinamicamente.
 */
function getPageHtml(pageName) {
  if (pageName === 'Dashboard' || pageName === 'Calltech') {
    return HtmlService.createHtmlOutputFromFile('Page_' + pageName).getContent();
  }
  throw new Error('Página não encontrada.');
}

/**
 * Permite a inclusão de arquivos HTML (usados para os scripts JS) dentro de outro template HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * NOVA FUNÇÃO: Busca e consolida os dados de um cliente de múltiplas fontes.
 */
function getCustomerData(filter) {
  try {
    const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
    const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);

    const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);
    const abaManager = planilhaCalltech.getSheetByName(NOME_ABA_MANAGER);
    const abaAtendimento = planilhaCalltech.getSheetByName(NOME_ABA_ATENDIMENTO);

    // Validação de abas
    if (!abaNPS || !abaManager || !abaAtendimento) {
      throw new Error("Uma ou mais abas necessárias não foram encontradas.");
    }
    
    // Usar getDisplayValues() para garantir que as datas sejam strings formatadas.
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
    
    // 2. Processa NPS para complementar base de clientes e mapear pedidos
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

    // 3. Filtra os clientes com base no critério (letra ou pesquisa)
    let filteredEmails = new Set();
    const filterType = filter.type;
    const filterValue = filter.value.toLowerCase();

    if (filterType === 'letter') {
      customers.forEach((customer, email) => {
        if (customer.name.toLowerCase().startsWith(filterValue)) {
          filteredEmails.add(email);
        }
      });
    } else if (filterType === 'search') {
      if (filterValue.length < 3) return []; // Evita buscas muito amplas
      customers.forEach((customer, email) => {
        if (customer.name.toLowerCase().includes(filterValue) || email.includes(filterValue)) {
          filteredEmails.add(email);
        }
      });
      // Busca por pedido
      orderToEmailMap.forEach((email, pedidoId) => {
        if(pedidoId.includes(filterValue) && customers.has(email)) {
           filteredEmails.add(email);
        }
      });
    }

    if(filteredEmails.size === 0) return [];
    
    // 4. Monta o histórico APENAS para os clientes filtrados
    const results = new Map();
    filteredEmails.forEach(email => {
      results.set(email, customers.get(email));
    });

    // Histórico de NPS
    npsData.forEach(row => {
      const email = row[INDICES_NPS.EMAIL]?.toString().trim().toLowerCase();
      const dateStr = row[INDICES_NPS.DATA_AVALIACAO];
      if(results.has(email) && dateStr){
         results.get(email).history.push({
           type: 'NPS',
           date: new Date(dateStr.split(' ')[0]), // Formato da planilha é YYYY-MM-DD
           pedidoId: row[INDICES_NPS.PEDIDO_ID],
           classificacao: row[INDICES_NPS.CLASSIFICACAO],
           comentario: row[INDICES_NPS.COMENTARIO]
         });
      }
    });

    // Histórico de Chamados (Manager)
    managerData.forEach(row => {
      const email = row[INDICES_CALLTECH.EMAIL]?.toString().trim().toLowerCase();
      const dateStr = row[INDICES_CALLTECH.DATA_ABERTURA];
      if(results.has(email) && dateStr){
         results.get(email).history.push({
           type: 'Chamado',
           date: new Date(dateStr.split(' ')[0].split('/').reverse().join('-')),
           chamadoId: row[INDICES_CALLTECH.CHAMADO_ID],
           pedidoId: row[INDICES_CALLTECH.PEDIDO_ID],
           status: row[INDICES_CALLTECH.STATUS],
           dataFinalizacao: row[INDICES_CALLTECH.DATA_FINALIZACAO]
         });
      }
    });
    
    // Histórico de Atendimentos (Forms)
    atendimentoData.forEach(row => {
      const pedidoId = row[INDICES_ATENDIMENTO.PEDIDO_ID]?.toString().trim();
      const email = orderToEmailMap.get(pedidoId);
      const dateStr = row[INDICES_ATENDIMENTO.DATA_ATENDIMENTO];
      if(email && results.has(email) && dateStr){
         results.get(email).history.push({
           type: 'Atendimento',
           date: new Date(dateStr.split(' ')[0].split('/').reverse().join('-')),
           pedidoId: pedidoId,
           resolucao: row[INDICES_ATENDIMENTO.RESOLUCAO],
           defeito: row[INDICES_ATENDIMENTO.DEFEITO],
           relato: row[INDICES_ATENDIMENTO.RELATO_CLIENTE]
         });
      }
    });

    // 5. Finaliza e formata o retorno
    const finalResults = Array.from(results.values());
    finalResults.forEach(customer => {
      // Ordena o histórico pela data mais recente
      customer.history.sort((a, b) => b.date - a.date);
      // Formata as datas para string no formato ISO para evitar problemas de fuso horário
      customer.history.forEach(item => {
        if (item.date instanceof Date && !isNaN(item.date)) {
           item.date = item.date.toISOString();
        } else {
           item.date = null; // Data inválida
        }
      });
    });

    return finalResults.sort((a,b) => a.name.localeCompare(b.name));

  } catch (e) {
    Logger.log(`Erro fatal na função getCustomerData: ${e.stack}`);
    // Retorna um erro estruturado para o frontend
    return { error: `Erro ao buscar dados dos clientes: ${e.message}` };
  }
}


/**
 * NOVA FUNÇÃO: Busca e processa os dados para a página Calltech.
 * VERSÃO ROBUSTA: Lida com erros de formato de data e células vazias.
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
    const uniqueNPSRows = getUniqueValidRows(allNPSData, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO);
    uniqueNPSRows.forEach(row => {
        const pedidoId = row[INDICES_NPS.PEDIDO_ID]?.trim();
        const classification = row[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
        if (pedidoId && classification) {
            npsMap.set(pedidoId, classification);
        }
    });

    // --- Processamento para KPIs e Tabela ---
    let openTickets = 0;
    let closedTickets = 0;
    let totalResolutionTime = 0;
    let resolvedTicketsCount = 0;
    const resolutionCounts = { day0: 0, day1: 0, day2: 0, day3: 0, day4plus: 0 };
    const tickets = [];
    const npsFeedback = { total: 0, promoters: 0, neutrals: 0, detractors: 0 };
    const postServiceNps = { total: 0, promoters: 0, neutrals: 0, detractors: 0 }; // KPI NOVO
    const ticketPedidoIds = new Set();
    const closedTicketPedidoIds = new Set(); // Controle para o KPI novo


    allManagerData.forEach((row, index) => {
      const openDateStr = row[INDICES_CALLTECH.DATA_ABERTURA];
      const chamadoId = row[INDICES_CALLTECH.CHAMADO_ID];

      if (!chamadoId || !openDateStr) return; 

      try {
        const dateParts = openDateStr.split(' ')[0].split('/');
        if (dateParts.length !== 3) throw new Error(`Formato de data inválido: ${openDateStr}`);
        const [day, month, year] = dateParts;
        const openDate = new Date(year, month - 1, day);
        if (isNaN(openDate.getTime())) throw new Error(`Data inválida: ${openDateStr}`);
        
        const openDateISO = openDate.toISOString().split('T')[0];

        if (openDateISO >= dateRange.start && openDateISO <= dateRange.end) {
          const status = row[INDICES_CALLTECH.STATUS] || "";
          const closeDateStr = row[INDICES_CALLTECH.DATA_FINALIZACAO];
          const isClosed = status.toLowerCase().includes('finalizado') || status.toLowerCase().includes('resolvido');
          const pedidoId = row[INDICES_CALLTECH.PEDIDO_ID]?.trim();

          if (isClosed) {
            closedTickets++;
            if (closeDateStr) {
              const closeDateParts = closeDateStr.split(' ')[0].split('/');
              if (closeDateParts.length === 3) {
                const closeDate = new Date(closeDateParts[2], closeDateParts[1] - 1, closeDateParts[0]);
                if (!isNaN(closeDate.getTime())) {
                  const diffTime = Math.abs(closeDate - openDate);
                  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                  totalResolutionTime += diffDays;
                  resolvedTicketsCount++;
                  
                  if (diffDays === 0) resolutionCounts.day0++;
                  else if (diffDays === 1) resolutionCounts.day1++;
                  else if (diffDays === 2) resolutionCounts.day2++;
                  else if (diffDays === 3) resolutionCounts.day3++;
                  else if (diffDays >= 4) resolutionCounts.day4plus++;
                }
              }
            }
          } else {
            openTickets++;
          }

          tickets.push({
            status: status,
            chamadoId: chamadoId,
            dataAbertura: openDateStr.split(' ')[0],
            pedidoId: pedidoId,
            cliente: row[INDICES_CALLTECH.CLIENTE],
            dataFinalizacao: closeDateStr ? closeDateStr.split(' ')[0] : '',
            email: row[INDICES_CALLTECH.EMAIL],
            hasNps: npsMap.has(pedidoId) // Adiciona a flag se o pedido tem NPS
          });

          // Processa o feedback NPS geral (baseado em chamados ABERTOS no período)
          if (pedidoId && !ticketPedidoIds.has(pedidoId)) {
              if (npsMap.has(pedidoId)) {
                  const classification = npsMap.get(pedidoId);
                  npsFeedback.total++;
                  if (classification === 'promotor') npsFeedback.promoters++;
                  else if (classification === 'neutro') npsFeedback.neutrals++;
                  else if (classification === 'detrator') npsFeedback.detractors++;
              }
              ticketPedidoIds.add(pedidoId);
          }

          // NOVO: Processa o feedback NPS Pós-Atendimento (apenas para chamados FINALIZADOS)
          if (isClosed && pedidoId && !closedTicketPedidoIds.has(pedidoId)) {
            if (npsMap.has(pedidoId)) {
              const classification = npsMap.get(pedidoId);
              postServiceNps.total++;
              if (classification === 'promotor') postServiceNps.promoters++;
              else if (classification === 'neutro') postServiceNps.neutrals++;
              else if (classification === 'detrator') postServiceNps.detractors++;
            }
            closedTicketPedidoIds.add(pedidoId);
          }
        }
      } catch (e) {
        Logger.log(`Erro ao processar linha ${index + 2} da planilha Manager: ${e.message}. Dados da linha: ${row.join(', ')}`);
      }
    });
    
    const avgResolutionTime = resolvedTicketsCount > 0 ? (totalResolutionTime / resolvedTicketsCount).toFixed(1) : 0;
    
    const resolutionRate = {};
    for (const day in resolutionCounts) {
      const count = resolutionCounts[day];
      const percentage = resolvedTicketsCount > 0 ? (count / resolvedTicketsCount) * 100 : 0;
      resolutionRate[day] = { count: count, percentage: percentage };
    }
    
    // NOVO: Calcular o NPS Score para o novo KPI
    const postServiceNpsScore = postServiceNps.total > 0 
      ? parseFloat((((postServiceNps.promoters - postServiceNps.detractors) / postServiceNps.total) * 100).toFixed(1)) 
      : 0;
    postServiceNps.npsScore = postServiceNpsScore;

    // --- Processamento para KPI de Retenção ---
    const allAtendimentoData = abaAtendimento.getRange(2, 1, abaAtendimento.getLastRow() - 1, abaAtendimento.getLastColumn()).getDisplayValues();
    let retentionValue = 0;
    const statusRetido = "Retido no Atendimento (MSPC)";

    allAtendimentoData.forEach((row, index) => {
      const dateStr = row[INDICES_ATENDIMENTO.DATA_ATENDIMENTO];
      if (!dateStr) return;

      try {
        const dateParts = dateStr.split(' ')[0].split('/');
        if (dateParts.length !== 3) return;
        const [day, month, year] = dateParts;
        const atendimentoDate = new Date(year, month - 1, day);
        if (isNaN(atendimentoDate.getTime())) return;

        const atendimentoDateISO = atendimentoDate.toISOString().split('T')[0];
        
        if (atendimentoDateISO >= dateRange.start && atendimentoDateISO <= dateRange.end) {
          const status = row[INDICES_ATENDIMENTO.STATUS_ATENDIMENTO];
          if (status === statusRetido) {
            const valorStr = row[INDICES_ATENDIMENTO.VALOR_RETIDO] || '0';
            const valorNumerico = parseFloat(valorStr.replace('R$', '').replace(/\./g, '').replace(',', '.').trim());
            if (!isNaN(valorNumerico)) {
              retentionValue += valorNumerico;
            }
          }
        }
      } catch (e) {
        Logger.log(`Error processing retention value on row ${index + 2} from Forms: ${e.message}. Data: ${row.join(', ')}`);
      }
    });

    return {
      tickets: tickets,
      kpis: {
        total: tickets.length,
        open: openTickets,
        closed: closedTickets,
        avgTime: avgResolutionTime,
        retentionValue: retentionValue,
        npsFeedback: npsFeedback,
        postServiceNps: postServiceNps // NOVO KPI ADICIONADO
      },
      resolutionRate: resolutionRate
    };
  } catch (e) {
    Logger.log(`Erro fatal na função getCalltechData: ${e.stack}`);
    return { 
      tickets: [], 
      kpis: { total: 0, open: 0, closed: 0, avgTime: 0, retentionValue: 0, npsFeedback: { total: 0, promoters: 0, neutrals: 0, detractors: 0 }, postServiceNps: { total: 0, promoters: 0, neutrals: 0, detractors: 0, npsScore: 0 } },
      resolutionRate: {}
    };
  }
}

// NOVA FUNÇÃO PARA O GRÁFICO DE FLUXO
function getDailyFlowChartData(dateRange) {
  try {
    const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);
    const abaManager = planilhaCalltech.getSheetByName(NOME_ABA_MANAGER);
    if (!abaManager) throw new Error(`Aba ${NOME_ABA_MANAGER} não encontrada.`);
    
    const allManagerData = abaManager.getRange(2, 1, abaManager.getLastRow() - 1, abaManager.getLastColumn()).getDisplayValues();
    const dailyFlow = {};
    const startDate = new Date(dateRange.start.replace(/-/g, '/'));
    const endDate = new Date(dateRange.end.replace(/-/g, '/'));

    // Initialize all days in the range to ensure empty days appear on the chart
    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        const dayKey = d.toISOString().split('T')[0];
        dailyFlow[dayKey] = { opened: 0, closed: 0 };
    }

    allManagerData.forEach((row, index) => {
        try {
            // Check opened tickets
            const openDateStr = row[INDICES_CALLTECH.DATA_ABERTURA];
            if (openDateStr) {
                const openDateParts = openDateStr.split(' ')[0].split('/');
                if (openDateParts.length === 3) {
                    const openDate = new Date(openDateParts[2], openDateParts[1] - 1, openDateParts[0]);
                    if (!isNaN(openDate.getTime())) {
                        const openDateISO = openDate.toISOString().split('T')[0];
                        if (dailyFlow.hasOwnProperty(openDateISO)) {
                           dailyFlow[openDateISO].opened++;
                        }
                    }
                }
            }
            // Check closed tickets
            const closeDateStr = row[INDICES_CALLTECH.DATA_FINALIZACAO];
            if (closeDateStr) {
                const closeDateParts = closeDateStr.split(' ')[0].split('/');
                if (closeDateParts.length === 3) {
                    const closeDate = new Date(closeDateParts[2], closeDateParts[1] - 1, closeDateParts[0]);
                    if (!isNaN(closeDate.getTime())) {
                        const closeDateISO = closeDate.toISOString().split('T')[0];
                        if (dailyFlow.hasOwnProperty(closeDateISO)) {
                           dailyFlow[closeDateISO].closed++;
                        }
                    }
                }
            }
        } catch (e) {
            Logger.log(`Error processing daily flow on row ${index + 2}: ${e.message}`);
        }
    });

    const weekDayInitials = ['D', 'S', 'T', 'Q', 'Q', 'S', 'S']; // Dom, Seg, Ter, Qua, Qui, Sex, Sab
    const dailyFlowChartData = [['Dia', 'Abertos', 'Fechados']];
    Object.keys(dailyFlow).sort().forEach(dayISO => {
        const date = new Date(dayISO.replace(/-/g, '/'));
        const dayOfWeek = date.getUTCDay(); // Usar getUTCDay para consistência
        const weekDayInitial = weekDayInitials[dayOfWeek];

        const [year, month, day] = dayISO.split('-');
        const label = `${day}/${month}\n(${weekDayInitial})`;
        dailyFlowChartData.push([label, dailyFlow[dayISO].opened, dailyFlow[dayISO].closed]);
    });
    
    return dailyFlowChartData;
  } catch (e) {
    Logger.log(`Erro fatal na função getDailyFlowChartData: ${e.stack}`);
    return [['Dia', 'Abertos', 'Fechados']]; // Return headers on error
  }
}

/**
 * NOVA FUNÇÃO: Busca dados para o gráfico de timeline de performance do NPS.
 * Agrega dados de NPS por semana e cruza com ações estratégicas.
 */
function getNpsTimelineData(dateRange) {
  try {
    const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
    const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);
    const abaAcoes = planilhaNPS.getSheetByName(NOME_ABA_ACOES); 

    if (!abaNPS || !abaAcoes) {
      throw new Error("Aba de NPS ou de Ações não encontrada.");
    }

    // Pega todos os dados de NPS e Ações
    const todosDadosNPS = abaNPS.getRange(2, 1, abaNPS.getLastRow() - 1, abaNPS.getLastColumn()).getValues();
    const dadosAcoes = abaAcoes.getRange(2, 1, abaAcoes.getLastRow() - 1, 6).getValues(); // Coluna 1 (A) e 6 (F)

    // Filtra os dados de NPS pelo range de data solicitado
    const dadosFiltrados = todosDadosNPS.filter(linha => {
      const dataAvaliacao = linha[INDICES_NPS.DATA_AVALIACAO];
      if (!dataAvaliacao || !(dataAvaliacao instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataAvaliacao, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });

    const dadosUnicos = getUniqueValidRows(dadosFiltrados, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO);

    // Agrupa as métricas por semana
    const weeklyMetrics = {};
    dadosUnicos.forEach(linha => {
      const dataAvaliacao = linha[INDICES_NPS.DATA_AVALIACAO];
      const date = new Date(dataAvaliacao);
      
      // Chave da semana (baseada no domingo daquela semana)
      const firstDayOfWeek = new Date(date.setDate(date.getDate() - date.getDay()));
      const key = Utilities.formatDate(firstDayOfWeek, "GMT-3", "yyyy-MM-dd");

      if (!weeklyMetrics[key]) {
        weeklyMetrics[key] = { promoters: 0, neutrals: 0, detractors: 0, endDate: new Date(firstDayOfWeek.getTime() + 6 * 24 * 60 * 60 * 1000) };
      }

      const classificacao = linha[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
      if (classificacao === 'promotor') weeklyMetrics[key].promoters++;
      else if (classificacao === 'neutro') weeklyMetrics[key].neutrals++;
      else if (classificacao === 'detrator') weeklyMetrics[key].detractors++;
    });

    // Processa as ações e as associa a cada semana
    const actionsByWeek = {};
    dadosAcoes.forEach(acao => {
        const acaoTexto = acao[0];
        const acaoData = acao[5];
        if (acaoTexto && acaoData instanceof Date) {
            const date = new Date(acaoData);
            const firstDayOfWeek = new Date(date.setDate(date.getDate() - date.getDay()));
            const key = Utilities.formatDate(firstDayOfWeek, "GMT-3", "yyyy-MM-dd");
            if (!actionsByWeek[key]) {
                actionsByWeek[key] = [];
            }
            actionsByWeek[key].push(acaoTexto);
        }
    });

    // Monta o array final, calculando o NPS e as variações
    const sortedKeys = Object.keys(weeklyMetrics).sort();
    const result = [];
    let previousWeekMetrics = null;

    sortedKeys.forEach(key => {
      const metrics = weeklyMetrics[key];
      const total = metrics.promoters + metrics.neutrals + metrics.detractors;
      const nps = total > 0 ? parseFloat((((metrics.promoters - metrics.detractors) / total) * 100).toFixed(1)) : 0;
      
      const weekData = {
        weekLabel: `Semana de ${Utilities.formatDate(new Date(key), "GMT-3", "dd/MM")}`,
        nps: nps,
        detractors: metrics.detractors,
        neutrals: metrics.neutrals,
        actions: actionsByWeek[key] || [],
        npsChange: 0,
        detractorsChange: 0,
        neutralsChange: 0
      };

      if (previousWeekMetrics) {
        weekData.npsChange = parseFloat((nps - previousWeekMetrics.nps).toFixed(1));
        weekData.detractorsChange = metrics.detractors - previousWeekMetrics.detractors;
        weekData.neutralsChange = metrics.neutrals - previousWeekMetrics.neutrals;
      }
      
      result.push(weekData);
      previousWeekMetrics = { nps, detractors: metrics.detractors, neutrals: metrics.neutrals };
    });

    return result;

  } catch (e) {
    Logger.log(`Erro em getNpsTimelineData: ${e.stack}`);
    return { error: e.message };
  }
}


// ==================================================================
// === FUNÇÕES EXISTENTES (NPS, GEMINI, ETC) - SEM ALTERAÇÕES ABAIXO ===
// ==================================================================

function callGeminiAPI(prompt, schema) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error("A chave de API do Google AI não foi configurada.");
  }
  const payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json", responseSchema: schema }
  };
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  const response = UrlFetchApp.fetch(apiUrl, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  if (responseCode !== 200) {
    Logger.log(`Erro na API: ${responseCode} - ${responseText}`);
    throw new Error(`Erro na API: ${responseText}`);
  }
  const result = JSON.parse(responseText);
  if (result.candidates && result.candidates.length > 0 && result.candidates[0].content.parts[0].text) {
    return JSON.parse(result.candidates[0].content.parts[0].text);
  } else {
    Logger.log("Resposta da API inválida: " + responseText);
    throw new Error("Resposta da API inválida.");
  }
}

function getProblemAndSuggestionAnalysis(comments) {
  try {
    const prompt = `
      Aja como um analista de dados sênior, especialista em Customer Experience.
      Analise os seguintes comentários e identifique os 3 problemas mais críticos. Para cada um, sugira uma ação corretiva.
      Comentários: ${JSON.stringify(comments)}
      Retorne a resposta estritamente no formato JSON.`;
    const schema = {
      type: "OBJECT",
      properties: {
        "problemas_recorrentes": { "type": "ARRAY", "items": { "type": "OBJECT", "properties": { "titulo": { "type": "STRING" }, "descricao": { "type": "STRING" } } } },
        "sugestoes_de_acoes": { "type": "ARRAY", "items": { "type": "OBJECT", "properties": { "titulo": { "type": "STRING" }, "descricao": { "type": "STRING" } } } }
      }
    };
    return callGeminiAPI(prompt, schema);
  } catch (e) {
    Logger.log(e);
    return { problemas_recorrentes: [{titulo: "Erro na Análise", descricao: e.message}], sugestoes_de_acoes: [] };
  }
}

function getTopicAnalysis(comments) {
  try {
    const prompt = `
      Aja como um analista de dados sênior.
      Analise os seguintes comentários e identifique os 4 tópicos mais comentados (ex: "Transporte", "Qualidade de Montagem").
      Para cada tópico, inclua a frequência (quantos comentários o mencionam) e o sentimento médio ("Positivo", "Negativo", "Neutro").
      Comentários: ${JSON.stringify(comments)}
      Retorne a resposta estritamente no formato JSON.`;
    const schema = {
      type: "OBJECT",
      properties: {
        "principais_topicos": { "type": "ARRAY", "items": { "type": "OBJECT", "properties": { "topico": { "type": "STRING" }, "frequencia": { "type": "INTEGER" }, "sentimento_medio": { "type": "STRING" } } } }
      }
    };
    return callGeminiAPI(prompt, schema);
  } catch (e) {
    Logger.log(e);
    return { principais_topicos: [] };
  }
}

function getPraiseAnalysis(comments) {
  try {
    const prompt = `
      Aja como um analista de marketing.
      Analise os seguintes comentários e identifique os 3 pontos positivos mais elogiados pelos clientes.
      Comentários: ${JSON.stringify(comments)}
      Retorne a resposta estritamente no formato JSON.`;
    const schema = {
      type: "OBJECT",
      properties: { "elogios_destacados": { "type": "ARRAY", "items": { "type": "OBJECT", "properties": { "titulo": { "type": "STRING" }, "descricao": { "type": "STRING" } } } } }
    };
    return callGeminiAPI(prompt, schema);
  } catch (e) {
    Logger.log(e);
    return { elogios_destacados: [] };
  }
}

function getWordFrequencyAnalysis(comments) {
  const stopWords = new Set(['de', 'a', 'o', 'que', 'e', 'do', 'da', 'em', 'um', 'para', 'com', 'não', 'uma', 'os', 'no', 'na', 'por', 'mais', 'as', 'dos', 'como', 'mas', 'foi', 'ao', 'ele', 'das', 'tem', 'à', 'seu', 'sua', 'ou', 'ser', 'quando', 'muito', 'há', 'nos', 'já', 'está', 'eu', 'também', 'só', 'pelo', 'pela', 'até', 'isso', 'ela', 'entre', 'era', 'depois', 'sem', 'mesmo', 'aos', 'ter', 'seus', 'quem', 'nas', 'me', 'esse', 'eles', 'estão', 'você', 'tinha', 'foram', 'essa', 'num', 'nem', 'suas', 'meu', 'às', 'minha', 'numa', 'pelos', 'elas', 'havia', 'seja', 'qual', 'será', 'nós', 'tenho', 'lhe', 'deles', 'essas', 'esses', 'pelas', 'este', 'fosse', 'dele', 'tu', 'te', 'vocês', 'vos', 'lhes', 'meus', 'minhas', 'teu', 'tua', 'teus', 'tuas', 'nosso', 'nossa', 'nossos', 'nossas', 'dela', 'delas', 'esta', 'estes', 'estas', 'aquele', 'aquela', 'aqueles', 'aquelas', 'isto', 'aquilo', 'estou', 'está', 'estamos', 'estão', 'estive', 'esteve', 'estivemos', 'estiveram', 'estava', 'estávamos', 'estavam', 'estivera', 'estivéramos', 'esteja', 'estejamos', 'estejam', 'estivesse', 'estivéssemos', 'estivessem', 'estiver', 'estivermos', 'estiverem', 'hei', 'há', 'havemos', 'hão', 'houve', 'houvemos', 'houveram', 'houvera', 'houvéramos', 'haja', 'hajamos', 'hajam', 'houvesse', 'houvéssemos', 'houvessem', 'houver', 'houvermos', 'houverem', 'houverei', 'houverá', 'houveremos', 'houverão', 'houveria', 'houveríamos', 'houveriam', 'sou', 'somos', 'são', 'era', 'éramos', 'eram', 'fui', 'foi', 'fomos', 'foram', 'fora', 'fôramos', 'seja', 'sejamos', 'sejam', 'fosse', 'fôssemos', 'fossem', 'for', 'formos', 'forem', 'serei', 'será', 'seremos', 'serão', 'seria', 'seríamos', 'seriam', 'tenho', 'tem', 'temos', 'tém', 'tinha', 'tínhamos', 'tinham', 'tive', 'teve', 'tivemos', 'tiveram', 'tivera', 'tivéramos', 'tenha', 'tenhamos', 'tenham', 'tivesse', 'tivéssemos', 'tivessem', 'tiver', 'tivermos', 'tiverem', 'terei', 'terá', 'teremos', 'terão', 'teria', 'teríamos', 'teriam', 'pc', 'computador']);
  const categories = {
    "Funcionamento do PC": ['desempenho', 'rápido', 'xmp', 'placa','placa-mãe','windows', 'lento', 'travando', 'temperatura', 'fps', 'funciona', 'funcionando', 'problema', 'defeito', 'liga', 'desliga', 'performance', 'jogos', 'roda', 'rodando'],
    "Qualidade de Montagem": ['montagem', 'cabo', 'cabos', 'organização', 'organizado', 'arrumado', 'encaixado', 'solto', 'peça', 'peças', 'cable', 'management'],
    "Visual do PC": ['visual', 'bonito', 'lindo', 'led', 'leds', 'rgb', 'aparência', 'gabinete', 'fan', 'fans', 'cooler', 'bonita'],
    "Transporte": ['entrega', 'caixa', 'embalagem', 'chegou', 'transporte', 'danificado', 'amassada', 'amassado', 'veio', 'rápida', 'rápido', 'demorou'],
    "Atendimento": ['atendimento', 'atendente', 'suporte', 'demorou', 'rapido', 'problema', 'resolvido', 'atencioso', 'educado', 'mal', 'bom', 'otimo', 'pessimo', 'ruim']
  };
  const wordCounts = { "Funcionamento do PC": {}, "Qualidade de Montagem": {}, "Visual do PC": {}, "Transporte": {}, "Atendimento": {} };
  comments.forEach(comment => {
    if (!comment) return;
    const words = comment.toLowerCase().match(/\b(\w+)\b/g) || [];
    words.forEach(word => {
      if (stopWords.has(word) || word.length < 3) return;
      for (const category in categories) {
        if (categories[category].includes(word)) {
          wordCounts[category][word] = (wordCounts[category][word] || 0) + 1;
          break;
        }
      }
    });
  });
  const result = {};
  for (const category in wordCounts) {
    result[category] = Object.entries(wordCounts[category]).sort(([, a], [, b]) => b - a).slice(0, 10).map(([palavra, frequencia]) => ({ palavra, frequencia }));
  }
  return result;
}

function getUniqueValidRows(dados, idIndex, classIndex) {
  const pedidosProcessados = new Map();
  const validClassifications = ['promotor', 'detrator', 'neutro'];
  for (let i = dados.length - 1; i >= 0; i--) {
    const linha = dados[i];
    const pedidoId = linha[idIndex]?.toString().trim(); // CORREÇÃO APLICADA AQUI
    const classificacao = linha[classIndex]?.toString().toLowerCase();
    if (pedidoId && validClassifications.includes(classificacao)) {
      if (!pedidosProcessados.has(pedidoId)) {
        pedidosProcessados.set(pedidoId, linha);
      }
    }
  }
  return Array.from(pedidosProcessados.values());
}

function getEvolutionChartData(dateRange, groupBy) {
  const ss = SpreadsheetApp.openById(ID_PLANILHA_NPS);
  const abaDados = ss.getSheetByName(NOME_ABA_NPS);
  if (!abaDados) return [];
  const todosOsDados = abaDados.getRange(2, 1, abaDados.getLastRow() - 1, abaDados.getLastColumn()).getDisplayValues();
  const dadosFiltradosPorData = todosOsDados.filter(linha => {
    const dataDisparoTexto = linha[INDICES_NPS.DATA_AVALIACAO];
    if (!dataDisparoTexto) return false;
    const dataPlanilhaFormatada = dataDisparoTexto.substring(0, 10);
    return dataPlanilhaFormatada >= dateRange.start && dataPlanilhaFormatada <= dateRange.end;
  });
  const dadosUnicos = getUniqueValidRows(dadosFiltradosPorData, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO);
  if (groupBy === 'week') { return calculateWeeklyEvolution(dadosUnicos); }
  return calculateMonthlyEvolution(dadosUnicos);
}

function calculateMonthlyEvolution(dadosUnicos) {
    const monthlyMetrics = {};
    const monthNames = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
    dadosUnicos.forEach(linha => {
        const dateStr = linha[INDICES_NPS.DATA_AVALIACAO];
        if (!dateStr) return;
        const date = new Date(dateStr.substring(0, 10).replace(/-/g, '/'));
        if (isNaN(date.getTime())) return;
        const month = date.getMonth();
        const year = date.getFullYear();
        const key = `${year}-${String(month).padStart(2, '0')}`;
        const monthLabel = `${monthNames[month]}. de ${year}`;
        if (!monthlyMetrics[key]) { monthlyMetrics[key] = { label: monthLabel, promoters: 0, neutrals: 0, detractors: 0, firstDayOfMonth: new Date(year, month, 1) }; }
        const classificacao = linha[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
        if (classificacao === 'promotor') monthlyMetrics[key].promoters++;
        else if (classificacao === 'detrator') monthlyMetrics[key].detractors++;
        else if (classificacao === 'neutro') monthlyMetrics[key].neutrals++;
    });
    const sortedKeys = Object.keys(monthlyMetrics).sort();
    const dataTable = [['Mês', 'Total', 'Promotores', 'Detratores', 'Neutros', 'NPS', {type: 'string', role: 'annotation'}, {type: 'string', role: 'id'}]];
    sortedKeys.forEach(key => {
        const metrics = monthlyMetrics[key];
        const total = metrics.promoters + metrics.neutrals + metrics.detractors;
        const nps = total > 0 ? parseFloat((((metrics.promoters - metrics.detractors) / total) * 100).toFixed(1)) : 0;
        dataTable.push([metrics.label, total, metrics.promoters, metrics.detractors, metrics.neutrals, nps, nps.toFixed(1), metrics.firstDayOfMonth.toISOString()]);
    });
    return dataTable;
}

function calculateWeeklyEvolution(dadosUnicos) {
    const weeklyMetrics = {};
    dadosUnicos.forEach(linha => {
        const dateStr = linha[INDICES_NPS.DATA_AVALIACAO];
        if (!dateStr) return;
        const date = new Date(dateStr.substring(0, 10).replace(/-/g, '/'));
        if (isNaN(date.getTime())) return;
        const firstDayOfWeek = new Date(date.setDate(date.getDate() - date.getDay()));
        const key = firstDayOfWeek.toISOString().substring(0, 10);
        const weekLabel = `Semana de ${key.substring(8,10)}/${key.substring(5,7)}`;
        if (!weeklyMetrics[key]) { weeklyMetrics[key] = { label: weekLabel, promoters: 0, neutrals: 0, detractors: 0 }; }
        const classificacao = linha[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
        if (classificacao === 'promotor') weeklyMetrics[key].promoters++;
        else if (classificacao === 'detrator') weeklyMetrics[key].detractors++;
        else if (classificacao === 'neutro') weeklyMetrics[key].neutrals++;
    });
    const sortedKeys = Object.keys(weeklyMetrics).sort();
    const dataTable = [['Semana', 'Total', 'Promotores', 'Detratores', 'Neutros', 'NPS', {type: 'string', role: 'annotation'}]];
    sortedKeys.forEach(key => {
        const metrics = weeklyMetrics[key];
        const total = metrics.promoters + metrics.neutrals + metrics.detractors;
        const nps = total > 0 ? parseFloat((((metrics.promoters - metrics.detractors) / total) * 100).toFixed(1)) : 0;
        dataTable.push([metrics.label, total, metrics.promoters, metrics.detractors, metrics.neutrals, nps, nps.toFixed(1)]);
    });
    return dataTable;
}

function getDashboardData(dateRange) {
  const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
  const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);
  const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);
  const abaAtendimento = planilhaCalltech.getSheetByName(NOME_ABA_ATENDIMENTO);
  const abaOS = planilhaNPS.getSheetByName(NOME_ABA_OS);
  if (!abaNPS) throw new Error(`A aba "${NOME_ABA_NPS}" não foi encontrada na planilha de NPS.`);
  if (!abaAtendimento) throw new Error(`A aba "${NOME_ABA_ATENDIMENTO}" não foi encontrada na planilha de CallTech.`);
  if (!abaOS) throw new Error(`A aba "${NOME_ABA_OS}" não foi encontrada na planilha de NPS.`);
  const todosDadosNPS = abaNPS.getRange(2, 1, abaNPS.getLastRow() - 1, abaNPS.getLastColumn()).getDisplayValues();
  const dadosAtendimento = abaAtendimento.getRange(2, 1, abaAtendimento.getLastRow() - 1, abaAtendimento.getLastColumn()).getDisplayValues();
  const dadosOS = abaOS.getRange(2, 1, abaOS.getLastRow() - 1, abaOS.getLastColumn()).getDisplayValues();
  const osPorPedido = new Map();
  dadosOS.forEach(linha => {
    const pedidoId = linha[INDICES_OS.PEDIDO_ID]?.trim();
    const os = linha[INDICES_OS.OS]?.trim();
    if (pedidoId && os) { osPorPedido.set(pedidoId, os); }
  });
  const atendimentoPorPedido = new Map();
  dadosAtendimento.forEach(linha => {
      const pedidoId = linha[INDICES_ATENDIMENTO.PEDIDO_ID]?.trim();
      if (pedidoId) {
          if (!atendimentoPorPedido.has(pedidoId)) { atendimentoPorPedido.set(pedidoId, []); }
          atendimentoPorPedido.get(pedidoId).push(linha);
      }
  });
  const anoCorrente = new Date().getFullYear().toString();
  const dadosUnicosDoAno = getUniqueValidRows(todosDadosNPS.filter(linha => {
    const dataDisparoTexto = linha[INDICES_NPS.DATA_AVALIACAO];
    return dataDisparoTexto && dataDisparoTexto.substring(0, 4) === anoCorrente;
  }), INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO);
  const metricasAcumuladas = calcularMetricasBasicas(dadosUnicosDoAno);
  const dadosParaMeta = todosDadosNPS.filter(linha => {
    const dataDisparoTexto = linha[INDICES_NPS.DATA_AVALIACAO];
    if (!dataDisparoTexto) return false;
    const data = new Date(dataDisparoTexto.substring(0, 10).replace(/-/g, '/'));
    return data.getFullYear().toString() === anoCorrente && data.getMonth() >= 3;
  });
  const dadosUnicosMeta = getUniqueValidRows(dadosParaMeta, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO);
  const metricasMeta = calcularMetricasBasicas(dadosUnicosMeta);
  const dadosFiltradosPorData = todosDadosNPS.filter(linha => {
    const dataDisparoTexto = linha[INDICES_NPS.DATA_AVALIACAO];
    return dataDisparoTexto && dataDisparoTexto.substring(0, 10) >= dateRange.start && dataDisparoTexto.substring(0, 10) <= dateRange.end;
  });
  const dadosUnicosDoPeriodo = getUniqueValidRows(dadosFiltradosPorData, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO);
  const startDate = new Date(dateRange.start.replace(/-/g, '/'));
  const endDate = new Date(dateRange.end.replace(/-/g, '/'));
  const duration = endDate.getTime() - startDate.getTime();
  const previousEndDate = new Date(startDate.getTime() - 24 * 60 * 60 * 1000);
  const previousStartDate = new Date(previousEndDate.getTime() - duration);
  const previousStartStr = previousStartDate.toISOString().split('T')[0];
  const previousEndStr = previousEndDate.toISOString().split('T')[0];
  const dadosPeriodoAnterior = todosDadosNPS.filter(linha => {
    const dataDisparoTexto = linha[INDICES_NPS.DATA_AVALIACAO];
    return dataDisparoTexto && dataDisparoTexto.substring(0, 10) >= previousStartStr && dataDisparoTexto.substring(0, 10) <= previousEndStr;
  });
  const dadosUnicosAnterior = getUniqueValidRows(dadosPeriodoAnterior, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO);
  const metricasAnteriores = calcularMetricasBasicas(dadosUnicosAnterior);
  const metricasDetalhadas = processarMetricasDetalhadas(dadosUnicosDoPeriodo, atendimentoPorPedido);
  const rawResponses = dadosUnicosDoPeriodo.map(linha => {
    const dateStr = linha[INDICES_NPS.DATA_AVALIACAO].substring(0, 10);
    const [year, month, day] = dateStr.split('-');
    const formattedDate = `${day}/${month}/${year}`;
    const pedidoId = linha[INDICES_NPS.PEDIDO_ID]?.trim();
    const os = osPorPedido.get(pedidoId) || '';
    return [
      formattedDate, pedidoId, os, linha[INDICES_NPS.CLIENTE], linha[INDICES_NPS.CLASSIFICACAO], linha[INDICES_NPS.COMENTARIO],
      linha[INDICES_NPS.MOTIVO_FUNCIONAMENTO], linha[INDICES_NPS.MOTIVO_QUALIDADE_MONTAGEM], linha[INDICES_NPS.MOTIVO_VISUAL_PC], linha[INDICES_NPS.MOTIVO_TRANSPORTE]
    ];
  }).sort((a, b) => {
      const [dayA, monthA, yearA] = a[0].split('/');
      const [dayB, monthB, yearB] = b[0].split('/');
      return new Date(`${yearB}-${monthB}-${dayB}`) - new Date(`${yearA}-${monthA}-${dayA}`);
  });
  const metricasPeriodo = calcularMetricasBasicas(dadosUnicosDoPeriodo);
  return {
    nps: metricasPeriodo.nps, totalRespostas: metricasPeriodo.totalRespostas, promotores: metricasPeriodo.promotores, neutros: metricasPeriodo.neutros, detratores: metricasPeriodo.detratores,
    previousNps: metricasAnteriores.totalRespostas > 0 ? metricasAnteriores.nps : null, npsAcumulado: metricasAcumuladas.nps, totalRespostasAcumulado: metricasAcumuladas.totalRespostas,
    npsMeta: metricasMeta.nps, totalRespostasMeta: metricasMeta.totalRespostas, contagemMotivos: metricasDetalhadas.contagemMotivos, detratoresComChamado: metricasDetalhadas.detratoresComChamado,
    detractorSupportReasons: metricasDetalhadas.detractorSupportReasons, rawResponses: rawResponses
  };
}

function calcularMetricasBasicas(dadosUnicos) {
  const totalRespostas = dadosUnicos.length;
  if (totalRespostas === 0) { return { nps: 0, promotores: 0, neutros: 0, detratores: 0, totalRespostas: 0 }; }
  let promotores = 0, detratores = 0, neutros = 0;
  for (const linha of dadosUnicos) {
    const classificacao = linha[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
    if (classificacao === 'promotor') promotores++;
    else if (classificacao === 'detrator') detratores++;
    else if (classificacao === 'neutro') neutros++;
  }
  const nps = totalRespostas > 0 ? parseFloat((((promotores - detratores) / totalRespostas) * 100).toFixed(1)) : 0;
  return { nps, promotores, neutros, detratores, totalRespostas };
}

function processarMetricasDetalhadas(dadosUnicos, atendimentoMap) {
  const contagemMotivos = { 'Funcionamento do PC': {}, 'Qualidade de Montagem': {}, 'Visual do PC': {}, 'Transporte': {} };
  const detratoresComChamado = {};
  const detractorSupportReasons = new Set();
  if (!dadosUnicos || dadosUnicos.length === 0) return { contagemMotivos, detratoresComChamado, detractorSupportReasons: [] };
  const categorias = Object.keys(contagemMotivos);
  const indicesMotivos = [ INDICES_NPS.MOTIVO_FUNCIONAMENTO, INDICES_NPS.MOTIVO_QUALIDADE_MONTAGEM, INDICES_NPS.MOTIVO_VISUAL_PC, INDICES_NPS.MOTIVO_TRANSPORTE ];
  dadosUnicos.forEach(linha => {
    categorias.forEach((cat, i) => {
      const motivo = linha[indicesMotivos[i]]?.toString().trim();
      if (motivo) { contagemMotivos[cat][motivo] = (contagemMotivos[cat][motivo] || 0) + 1; }
    });
    const classificacao = linha[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
    if (classificacao === 'detrator') {
      const pedidoId = linha[INDICES_NPS.PEDIDO_ID]?.trim();
      if (atendimentoMap && atendimentoMap.has(pedidoId)) {
          atendimentoMap.get(pedidoId).forEach(chamadoLinha => {
              const motivoChamado = chamadoLinha[INDICES_ATENDIMENTO.RESOLUCAO] || "Não especificado";
              detratoresComChamado[motivoChamado] = (detratoresComChamado[motivoChamado] || 0) + 1;
              detractorSupportReasons.add(motivoChamado);
          });
      }
    }
  });
  return { contagemMotivos, detratoresComChamado, detractorSupportReasons: Array.from(detractorSupportReasons).sort() };
}

function getDetractorSupportDetails(dateRange, reasons) {
  const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
  const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);
  const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);
  const abaAtendimento = planilhaCalltech.getSheetByName(NOME_ABA_ATENDIMENTO);
  const abaOS = planilhaNPS.getSheetByName(NOME_ABA_OS);
  const todosDadosNPS = abaNPS.getRange(2, 1, abaNPS.getLastRow() - 1, abaNPS.getLastColumn()).getDisplayValues();
  const dadosAtendimento = abaAtendimento.getRange(2, 1, abaAtendimento.getLastRow() - 1, abaAtendimento.getLastColumn()).getDisplayValues();
  const dadosOS = abaOS.getRange(2, 1, abaOS.getLastRow() - 1, abaOS.getLastColumn()).getDisplayValues();
  const osPorPedido = new Map();
  dadosOS.forEach(linha => {
    const pedidoId = linha[INDICES_OS.PEDIDO_ID]?.trim();
    const os = linha[INDICES_OS.OS]?.trim();
    if (pedidoId && os) { osPorPedido.set(pedidoId, os); }
  });
  const dadosFiltradosPorData = todosDadosNPS.filter(linha => {
    const dataDisparoTexto = linha[INDICES_NPS.DATA_AVALIACAO];
    return dataDisparoTexto && dataDisparoTexto.substring(0, 10) >= dateRange.start && dataDisparoTexto.substring(0, 10) <= dateRange.end;
  });
  const detratores = getUniqueValidRows(dadosFiltradosPorData, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO)
                      .filter(linha => linha[INDICES_NPS.CLASSIFICACAO].toLowerCase() === 'detrator');
  const atendimentoPorPedido = new Map();
  dadosAtendimento.forEach(linha => {
      const pedidoId = linha[INDICES_ATENDIMENTO.PEDIDO_ID]?.trim();
      if (pedidoId) {
          if (!atendimentoPorPedido.has(pedidoId)) { atendimentoPorPedido.set(pedidoId, []); }
          atendimentoPorPedido.get(pedidoId).push(linha);
      }
  });
  const details = [];
  detratores.forEach(detratorRow => {
    const pedidoId = detratorRow[INDICES_NPS.PEDIDO_ID].trim();
    if (atendimentoPorPedido.has(pedidoId)) {
      atendimentoPorPedido.get(pedidoId).forEach(atendimentoRow => {
        const resolucao = atendimentoRow[INDICES_ATENDIMENTO.RESOLUCAO];
        if (reasons.includes(resolucao)) {
          const dateStr = detratorRow[INDICES_NPS.DATA_AVALIACAO].substring(0, 10);
          const [year, month, day] = dateStr.split('-');
          details.push({
            dataAvaliacao: `${day}/${month}/${year}`, pedido: pedidoId, os: osPorPedido.get(pedidoId) || '', cliente: detratorRow[INDICES_NPS.CLIENTE],
            classificacao: detratorRow[INDICES_NPS.CLASSIFICACAO], comentario: detratorRow[INDICES_NPS.COMENTARIO], resolucao: resolucao,
            defeito: atendimentoRow[INDICES_ATENDIMENTO.DEFEITO], relatoCliente: atendimentoRow[INDICES_ATENDIMENTO.RELATO_CLIENTE]
          });
        }
      });
    }
  });
  return details;
}

