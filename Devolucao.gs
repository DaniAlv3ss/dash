/**
 * Contém todas as funções do lado do servidor para a página do Dashboard de Devolução.
 */

/**
 * Busca e processa todos os dados necessários para os KPIs e gráficos do dashboard de devolução.
 */
function getDevolucaoPageData(dateRange) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilha.getSheetByName(NOME_ABA_DEVOLUCAO);
    if (!abaDevolucao) {
      throw new Error("Aba 'Base Devolução' não foi encontrada na planilha.");
    }

    const todosDados = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, abaDevolucao.getLastColumn()).getValues();

    const INDICES = {
      PEDIDO_ID: 0, CLIENTE: 1, NFE_NUMERO: 2, DATA_NFE: 3, CF_PRODUTO: 9, PRODUTO: 10,
      CATEGORIA_BUDGET: 12, QTD_DEVOLVIDA: 17, VALOR_DEVOLUCAO: 24, MOTIVO: 26, FABRICANTE: 28,
    };
    
    const hoje = new Date();
    const anoCorrente = hoje.getFullYear();
    const mesCorrente = hoje.getMonth();
    const devolucoesPorMes = {};
    const monthNames = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
    const nfsUnicasPorMes = {};

    todosDados.forEach(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (dataNfe instanceof Date && dataNfe.getFullYear() === anoCorrente) {
        const valorRaw = linha[INDICES.VALOR_DEVOLUCAO];
        const valor = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);
        const nfUnica = linha[INDICES.NFE_NUMERO] ? String(linha[INDICES.NFE_NUMERO]).trim() : null;
        const month = dataNfe.getMonth();
        const year = dataNfe.getFullYear();
        const key = `${year}-${String(month).padStart(2, '0')}`;
        if (!devolucoesPorMes[key]) {
          devolucoesPorMes[key] = { valor: 0 };
          nfsUnicasPorMes[key] = new Set();
        }
        devolucoesPorMes[key].valor += valor;
        if(nfUnica) nfsUnicasPorMes[key].add(nfUnica);
      }
    });

    const chartData = [['Mês', 'Valor Devolvido (R$)', 'Devoluções (NF-e)']];
    for(let i=0; i <= mesCorrente; i++) {
        const key = `${anoCorrente}-${String(i).padStart(2, '0')}`;
        const monthLabel = `${monthNames[i]}/${String(anoCorrente).slice(-2)}`;
        const valor = devolucoesPorMes[key] ? devolucoesPorMes[key].valor : 0;
        const quantidade = nfsUnicasPorMes[key] ? nfsUnicasPorMes[key].size : 0;
        chartData.push([monthLabel, valor, quantidade]);
    }

    const dadosFiltrados = todosDados.filter(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (!dataNfe || !(dataNfe instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });

    const startDate = new Date(dateRange.start + 'T05:00:00Z');
    const endDate = new Date(dateRange.end + 'T05:00:00Z');
    const duration = endDate.getTime() - startDate.getTime();
    const previousEndDate = new Date(startDate.getTime() - 24 * 60 * 60 * 1000);
    const previousStartDate = new Date(previousEndDate.getTime() - duration);
    const previousStartStr = previousStartDate.toISOString().split('T')[0];
    const previousEndStr = previousEndDate.toISOString().split('T')[0];

    const dadosPeriodoAnterior = todosDados.filter(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (!dataNfe || !(dataNfe instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= previousStartStr && dataFormatada <= previousEndStr;
    });
    
    let totalDevolvido = 0, valorCancelamento = 0, totalItensDevolvidos = 0;
    const nfsUnicas = new Set();
    const categorias = {}, fabricantes = {};
    const tabelaDetalhada = {}, todosMotivos = new Set();
    const pedidosComDevolucao = {};
    const motivos = {};

    dadosFiltrados.forEach(linha => {
      const qtdDevolvida = parseInt(linha[INDICES.QTD_DEVOLVIDA], 10) || 0;
      if (qtdDevolvida === 0) return;

      const motivo = (linha[INDICES.MOTIVO] || "Não especificado").trim().toLowerCase();
      const valorRaw = linha[INDICES.VALOR_DEVOLUCAO];
      const valor = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);
      const rawNf = linha[INDICES.NFE_NUMERO];
      const nfUnica = (rawNf !== null && rawNf !== undefined && rawNf !== '') ? String(rawNf).trim() : null;
      const dataNfe = linha[INDICES.DATA_NFE];

      if (!motivos[motivo]) motivos[motivo] = { nfs: new Set(), valor: 0 };
      if (nfUnica) motivos[motivo].nfs.add(nfUnica);
      motivos[motivo].valor += valor;
      totalDevolvido += valor;
      if (motivo.includes("cancelamento")) valorCancelamento += valor;
      if(nfUnica) nfsUnicas.add(nfUnica);
      totalItensDevolvidos += qtdDevolvida;

      const fabricante = (linha[INDICES.FABRICANTE] || "Não especificado").trim();
      const categoriaBudget = (linha[INDICES.CATEGORIA_BUDGET] || "Não especificado").trim();
      categorias[categoriaBudget] = (categorias[categoriaBudget] || 0) + qtdDevolvida;
      fabricantes[fabricante] = (fabricantes[fabricante] || 0) + qtdDevolvida;
      
      const displayMotivo = motivo.charAt(0).toUpperCase() + motivo.slice(1);
      todosMotivos.add(displayMotivo);

      const cfProduto = linha[INDICES.CF_PRODUTO] || 'N/A';
      const produtoNome = (linha[INDICES.PRODUTO] || "Não especificado").trim(); 
      const produtoKey = `${cfProduto}|${produtoNome}`;
      if(!tabelaDetalhada[produtoKey]) tabelaDetalhada[produtoKey] = { cf: cfProduto, nome: produtoNome, total: 0, motivos: {} };
      tabelaDetalhada[produtoKey].total += qtdDevolvida;
      tabelaDetalhada[produtoKey].motivos[displayMotivo] = (tabelaDetalhada[produtoKey].motivos[displayMotivo] || 0) + qtdDevolvida;
      
      const pedidoId = linha[INDICES.PEDIDO_ID] ? String(linha[INDICES.PEDIDO_ID]).trim() : null;
      if (pedidoId) {
        if (!pedidosComDevolucao[pedidoId]) {
          pedidosComDevolucao[pedidoId] = {
            pedidoId: pedidoId,
            dataNfe: dataNfe instanceof Date ? Utilities.formatDate(dataNfe, "GMT-3", "dd/MM/yyyy") : 'N/A',
            cliente: linha[INDICES.CLIENTE] || "Não especificado",
            totalItensDevolvidos: 0,
            totalValorDevolvido: 0
          };
        }
        pedidosComDevolucao[pedidoId].totalItensDevolvidos += qtdDevolvida;
        pedidosComDevolucao[pedidoId].totalValorDevolvido += valor;
      }
    });
    
    const motivosAnteriores = {};
    dadosPeriodoAnterior.forEach(linha => {
        const qtdDevolvidaAnterior = parseInt(linha[INDICES.QTD_DEVOLVIDA], 10) || 0;
        if (qtdDevolvidaAnterior > 0) {
            const motivo = (linha[INDICES.MOTIVO] || "Não especificado").trim().toLowerCase();
            const rawNfAnterior = linha[INDICES.NFE_NUMERO];
            const nfUnicaAnterior = (rawNfAnterior !== null && rawNfAnterior !== undefined && rawNfAnterior !== '') ? String(rawNfAnterior).trim() : null;
            if (!motivosAnteriores[motivo]) motivosAnteriores[motivo] = new Set();
            if (nfUnicaAnterior) motivosAnteriores[motivo].add(nfUnicaAnterior);
        }
    });

    const topMotivos = Object.entries(motivos).map(([motivo, data]) => {
        const displayMotivo = motivo.charAt(0).toUpperCase() + motivo.slice(1);
        const quantidadeAnterior = motivosAnteriores[motivo] ? motivosAnteriores[motivo].size : 0;
        return [displayMotivo, { quantidade: data.nfs.size, valor: data.valor, quantidadeAnterior: quantidadeAnterior }];
    }).sort(([, a], [, b]) => b.quantidade - a.quantidade).slice(0, 4);

    const topCategorias = Object.entries(categorias).sort(([,a],[,b]) => b-a).slice(0, 10);
    const topFabricantes = Object.entries(fabricantes).sort(([,a],[,b]) => b-a).slice(0, 10);
    
    const tabelaPedidos = Object.values(pedidosComDevolucao).sort((a, b) => {
      try {
        const [dayA, monthA, yearA] = a.dataNfe.split('/');
        const [dayB, monthB, yearB] = b.dataNfe.split('/');
        return new Date(`${yearB}-${monthB}-${dayB}`) - new Date(`${yearA}-${monthA}-${dayA}`);
      } catch(e) { return 0; }
    });

    const devolucoesPorDia = {};
    dadosFiltrados.forEach(linha => {
        const dataNfe = linha[INDICES.DATA_NFE];
        if (dataNfe instanceof Date) {
            const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
            const nfUnica = linha[INDICES.NFE_NUMERO] ? String(linha[INDICES.NFE_NUMERO]).trim() : null;
            const motivo = (linha[INDICES.MOTIVO] || "").trim().toLowerCase();
            if (!devolucoesPorDia[dataFormatada]) {
                devolucoesPorDia[dataFormatada] = { total: new Set(), cancelamentos: new Set() };
            }
            if (nfUnica) {
                devolucoesPorDia[dataFormatada].total.add(nfUnica);
                if (motivo.includes("cancelamento")) {
                    devolucoesPorDia[dataFormatada].cancelamentos.add(nfUnica);
                }
            }
        }
    });

    const dailyChartData = [['Dia', 'Total Devoluções', 'Cancelamentos', { role: 'id' }]];
    Object.keys(devolucoesPorDia).sort().forEach(day => {
      const [year, month, dayOfMonth] = day.split('-');
      const label = `${dayOfMonth}/${month}`;
      const totalCount = devolucoesPorDia[day].total.size;
      const cancelCount = devolucoesPorDia[day].cancelamentos.size;
      dailyChartData.push([label, totalCount, cancelCount, day]);
    });

    return {
      kpis: { totalDevolvido, valorCancelamento, devolucoesUnicas: nfsUnicas.size, totalItensDevolvidos },
      topMotivos, topCategorias, topFabricantes,
      devolucoesPorMes: chartData,
      devolucoesPorDia: dailyChartData,
      tabelaDetalhada: { motivos: Array.from(todosMotivos).sort(), produtos: Object.values(tabelaDetalhada).sort((a,b) => b.total - a.total) },
      tabelaPedidos: tabelaPedidos
    };
  } catch (e) {
    Logger.log(`Erro em getDevolucaoPageData: ${e.stack}`);
    return { error: e.message };
  }
}

function getDevolucaoDailyBreakdown(dateRange) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilha.getSheetByName(NOME_ABA_DEVOLUCAO);
    if (!abaDevolucao) throw new Error("Aba 'Base Devolução' não encontrada.");

    const todosDados = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, 27).getValues();
    const INDICES = { DATA_NFE: 3, MOTIVO: 26, VALOR_DEVOLUCAO: 24, NFE_NUMERO: 2 };
    
    const dadosFiltrados = todosDados.filter(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (!dataNfe || !(dataNfe instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });

    const dailyBreakdown = dadosFiltrados.map(linha => ({
        d: Utilities.formatDate(linha[INDICES.DATA_NFE], "GMT-3", "yyyy-MM-dd"),
        m: (linha[INDICES.MOTIVO] || "Não especificado").trim().toLowerCase(),
        v: (typeof linha[INDICES.VALOR_DEVOLUCAO] === 'number') ? linha[INDICES.VALOR_DEVOLUCAO] : (parseFloat(String(linha[INDICES.VALOR_DEVOLUCAO]).replace(/[R$\s.]/g, '').replace(',', '.')) || 0),
        nf: linha[INDICES.NFE_NUMERO] ? String(linha[INDICES.NFE_NUMERO]).trim() : null
    }));

    return dailyBreakdown;

  } catch (e) {
    Logger.log(`Erro em getDevolucaoDailyBreakdown: ${e.stack}`);
    return []; 
  }
}

function getItensDevolvidosPorPedido(pedidoId) {
  if (!pedidoId) return { error: "ID do Pedido não fornecido." };
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilha.getSheetByName(NOME_ABA_DEVOLUCAO);
    if (!abaDevolucao) throw new Error("Aba 'Base Devolução' não foi encontrada.");
    const todosDados = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, abaDevolucao.getLastColumn()).getValues();
    const INDICES = { PEDIDO_ID: 0, CF_PRODUTO: 9, PRODUTO: 10, QTD_DEVOLVIDA: 17, VALOR_DEVOLUCAO: 24 };
    return todosDados
      .filter(linha => String(linha[INDICES.PEDIDO_ID]).trim() === String(pedidoId).trim() && (parseInt(linha[INDICES.QTD_DEVOLVIDA], 10) || 0) > 0)
      .map(linha => {
        const valorRaw = linha[INDICES.VALOR_DEVOLUCAO];
        const valor = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);
        return {
          cf: linha[INDICES.CF_PRODUTO] || 'N/A',
          descricao: linha[INDICES.PRODUTO] || 'Descrição indisponível',
          quantidade: parseInt(linha[INDICES.QTD_DEVOLVIDA], 10) || 0,
          valor: valor
        };
      });
  } catch (e) {
    Logger.log(`Erro em getItensDevolvidosPorPedido para o pedido ${pedidoId}: ${e.stack}`);
    return { error: `Erro ao buscar detalhes do pedido: ${e.message}` };
  }
}

function getDevolucaoExportData(dateRange) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilha.getSheetByName(NOME_ABA_DEVOLUCAO);
    if (!abaDevolucao) throw new Error("Aba 'Base Devolução' não foi encontrada na planilha.");
    const todosDados = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, abaDevolucao.getLastColumn()).getValues();
    const INDICES = {
      PEDIDO_ID: 0, CLIENTE: 1, NFE_NUMERO: 2, DATA_NFE: 3, PRODUTO: 10,
      QTD_DEVOLVIDA: 17, VALOR_DEVOLUCAO: 24, MOTIVO: 26,
    };
    const dadosFiltrados = todosDados.filter(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (!dataNfe || !(dataNfe instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });
    return dadosFiltrados.map(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      const valorRaw = linha[INDICES.VALOR_DEVOLUCAO];
      const valor = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);
      return [
        dataNfe instanceof Date ? Utilities.formatDate(dataNfe, "GMT-3", "dd/MM/yyyy") : 'N/A',
        linha[INDICES.NFE_NUMERO] || '', linha[INDICES.PEDIDO_ID] || '', linha[INDICES.CLIENTE] || '',
        linha[INDICES.PRODUTO] || '', linha[INDICES.QTD_DEVOLVIDA] || 0, valor,
        (linha[INDICES.MOTIVO] || "Não especificado").trim()
      ];
    });
  } catch (e) {
    Logger.log(`Erro em getDevolucaoExportData: ${e.stack}`);
    return { error: e.message };
  }
}

/**
 * MODIFICADO: Aceita um 'specificDate' opcional para filtrar por um único dia.
 */
function getPedidosPorMotivo(dateRange, motivo, specificDate) {
  try {
    const motivoLower = motivo.toLowerCase();
    const planilhaDev = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilhaDev.getSheetByName(NOME_ABA_DEVOLUCAO);
    if (!abaDevolucao) throw new Error("Aba 'Base Devolução' não encontrada.");
    const todosDadosDevolucao = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, 27).getValues();
    
    const pedidosFiltrados = {};
    
    todosDadosDevolucao.forEach(linha => {
      const dataNfe = linha[3];
      if (!dataNfe || !(dataNfe instanceof Date)) return;
      
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      const motivoLinha = (linha[26] || "").toLowerCase().trim();
      
      const isDateMatch = specificDate 
        ? (dataFormatada === specificDate) 
        : (dataFormatada >= dateRange.start && dataFormatada <= dateRange.end);

      if (isDateMatch && motivoLinha === motivoLower) {
          const pedidoId = linha[0] ? String(linha[0]).trim() : null;
          if (!pedidoId) return;
          if (!pedidosFiltrados[pedidoId]) {
            pedidosFiltrados[pedidoId] = {
              pedidoId: pedidoId, dataNfe: Utilities.formatDate(dataNfe, "GMT-3", "dd/MM/yyyy"),
              cliente: linha[1] || 'N/A', produtos: new Set(), justificativa: new Set()
            };
          }
          pedidosFiltrados[pedidoId].produtos.add(linha[10] || 'Produto não especificado');
      }
    });

    if (motivoLower.includes("cancelamento")) {
        const pedidosParaBuscar = Object.keys(pedidosFiltrados);
        if (pedidosParaBuscar.length > 0) {
            try {
              const abaTrocas = SpreadsheetApp.openById(ID_PLANILHA_TROCAS).getSheetByName(NOME_ABA_TROCAS);
              if (abaTrocas) {
                  abaTrocas.getRange(2, 1, abaTrocas.getLastRow() - 1, 5).getValues().forEach(linha => {
                      const pedidoId = linha[0] ? String(linha[0]).trim() : null;
                      if (pedidosFiltrados[pedidoId] && linha[4]) pedidosFiltrados[pedidoId].justificativa.add(`Troca: ${String(linha[4]).trim()}`);
                  });
              }
            } catch(e) { Logger.log(`Não foi possível ler a planilha de Trocas: ${e.message}`); }
            try {
              const abaIncompat = SpreadsheetApp.openById(ID_PLANILHA_INCOMPATIBILIDADE).getSheetByName("Divergência");
              if (abaIncompat) {
                  abaIncompat.getRange(2, 1, abaIncompat.getLastRow() - 1, 11).getValues().forEach(linha => {
                      const pedidoId = linha[8] ? String(linha[8]).trim() : null;
                      if (pedidosFiltrados[pedidoId] && linha[10]) pedidosFiltrados[pedidoId].justificativa.add(`Incompat.: ${String(linha[10]).trim()}`);
                  });
              }
            } catch(e) { Logger.log(`Não foi possível ler a planilha de Incompatibilidade: ${e.message}`); }
        }
    }
    return Object.values(pedidosFiltrados).map(p => ({
        pedidoId: p.pedidoId, dataNfe: p.dataNfe, cliente: p.cliente,
        produtos: Array.from(p.produtos),
        justificativa: p.justificativa.size > 0 ? Array.from(p.justificativa).join('; ') : 'Não encontrada'
    })).sort((a, b) => {
        const [dayA, monthA, yearA] = a.dataNfe.split('/');
        const [dayB, monthB, yearB] = b.dataNfe.split('/');
        return new Date(`${yearB}-${monthB}-${dayB}`) - new Date(`${yearA}-${monthA}-${dayA}`);
    });
  } catch (e) {
    Logger.log(`Erro em getPedidosPorMotivo: ${e.stack}`);
    return { error: e.message };
  }
}

// --- VERSÕES COM CACHE PARA SEREM CHAMADAS PELO CLIENTE ---
function getDevolucaoPageDataWithCache(dateRange) {
  const cacheKey = `devolucao_pagedata_v4_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getDevolucaoPageData, [dateRange]);
}

function getDevolucaoDailyBreakdownWithCache(dateRange) {
  const cacheKey = `devolucao_breakdown_v4_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getDevolucaoDailyBreakdown, [dateRange]);
}

function getItensDevolvidosPorPedidoWithCache(pedidoId) {
  const cacheKey = `devolucao_itens_pedido_v2_${pedidoId}`;
  return getOrSetCache(cacheKey, getItensDevolvidosPorPedido, [pedidoId], 300);
}

function getDevolucaoExportDataWithCache(dateRange) {
  Logger.log(`Buscando dados de exportação (sem cache) para o período ${dateRange.start} a ${dateRange.end}`);
  return getDevolucaoExportData(dateRange);
}

// MODIFICADO: Inclui a data específica na chave do cache para evitar conflitos.
function getPedidosPorMotivoWithCache(dateRange, motivo, specificDate) {
  const specificDateKey = specificDate || 'all-dates';
  const cacheKey = `pedidos_por_motivo_v4_${dateRange.start}_${dateRange.end}_${motivo.replace(/\s+/g, '_')}_${specificDateKey}`;
  return getOrSetCache(cacheKey, getPedidosPorMotivo, [dateRange, motivo, specificDate]);
}
