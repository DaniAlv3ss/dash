/**
 * Contém todas as funções do lado do servidor para a página do Dashboard de Devolução.
 */

/**
 * Busca e processa todos os dados necessários para os KPIs e gráficos do dashboard de devolução.
 * MODIFICADO: Adicionada agregação de dados para a "Tabela de Pedidos com Devolução".
 */
function getDevolucaoData(dateRange) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilha.getSheetByName(NOME_ABA_DEVOLUCAO);
    if (!abaDevolucao) {
      throw new Error("Aba 'Base Devolução' não foi encontrada na planilha.");
    }

    const todosDados = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, abaDevolucao.getLastColumn()).getValues();

    const INDICES = {
      PEDIDO_ID: 0,       // Coluna A
      CLIENTE: 1,         // Coluna B
      NFE_NUMERO: 2,      // Coluna C
      DATA_NFE: 3,        // Coluna D
      CF_PRODUTO: 9,      // Coluna J
      PRODUTO: 10,        // Coluna K
      CATEGORIA_BUDGET: 12, // Coluna M
      QTD_DEVOLVIDA: 17,  // Coluna R
      VALOR_DEVOLUCAO: 24,// Coluna Y
      MOTIVO: 26,         // Coluna AA
      FABRICANTE: 28,     // Coluna AC
    };
    
    // 1. Processa dados do ano inteiro para o gráfico de evolução mensal
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
        if(nfUnica) {
          nfsUnicasPorMes[key].add(nfUnica);
        }
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

    // 2. Filtra dados para o período selecionado (dateRange)
    const dadosFiltrados = todosDados.filter(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (!dataNfe || !(dataNfe instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });

    // 3. Calcula dados para o período anterior (para o indicador de tendência)
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
    
    // 4. Processa os dados filtrados para KPIs, listas de Top 10 e tabelas
    let totalDevolvido = 0, valorCancelamento = 0, totalItensDevolvidos = 0;
    const nfsUnicas = new Set();
    const categorias = {}, fabricantes = {};
    const tabelaDetalhada = {}, todosMotivos = new Set();
    const pedidosComDevolucao = {};

    const motivos = {}; // Estrutura: { motivo: { nfs: Set(), valor: 0 } }
    dadosFiltrados.forEach(linha => {
      const qtdDevolvida = parseInt(linha[INDICES.QTD_DEVOLVIDA], 10) || 0;
      if (qtdDevolvida === 0) return;

      const motivo = (linha[INDICES.MOTIVO] || "Não especificado").trim().toLowerCase();
      const valorRaw = linha[INDICES.VALOR_DEVOLUCAO];
      const valor = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);
      const rawNf = linha[INDICES.NFE_NUMERO];
      const nfUnica = (rawNf !== null && rawNf !== undefined && rawNf !== '') ? String(rawNf).trim() : null;

      if (!motivos[motivo]) {
        motivos[motivo] = { nfs: new Set(), valor: 0 };
      }
      if (nfUnica) {
        motivos[motivo].nfs.add(nfUnica);
      }
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
          const dataNfe = linha[INDICES.DATA_NFE];
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
            
            if (!motivosAnteriores[motivo]) {
                motivosAnteriores[motivo] = new Set();
            }
            if (nfUnicaAnterior) {
                motivosAnteriores[motivo].add(nfUnicaAnterior);
            }
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
            if (!devolucoesPorDia[dataFormatada]) devolucoesPorDia[dataFormatada] = new Set();
            if (nfUnica) devolucoesPorDia[dataFormatada].add(nfUnica);
        }
    });
    const dailyChartData = [['Dia', 'Devoluções (NF-e)'], ...Object.keys(devolucoesPorDia).sort().map(day => {
        const [,, dayOfMonth] = day.split('-'); return [`${dayOfMonth}/${day.substring(5,7)}`, devolucoesPorDia[day].size];
    })];

    return {
      kpis: { totalDevolvido, valorCancelamento, devolucoesUnicas: nfsUnicas.size, totalItensDevolvidos },
      topMotivos, topCategorias, topFabricantes,
      devolucoesPorMes: chartData,
      devolucoesPorDia: dailyChartData,
      tabelaDetalhada: { motivos: Array.from(todosMotivos).sort(), produtos: Object.values(tabelaDetalhada).sort((a,b) => b.total - a.total) },
      tabelaPedidos: tabelaPedidos
    };
  } catch (e) {
    Logger.log(`Erro em getDevolucaoData: ${e.stack}`);
    return { error: e.message };
  }
}

/**
 * Busca os itens específicos devolvidos para um determinado número de pedido.
 * @param {string} pedidoId O ID do pedido a ser consultado.
 * @returns {Array<Object>} Uma lista de itens devolvidos com seus detalhes.
 */
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


/**
 * NOVA FUNÇÃO: Busca dados detalhados de devolução para exportação em CSV.
 * Retorna uma lista não agregada de itens devolvidos no período.
 */
function getDevolucaoExportData(dateRange) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilha.getSheetByName(NOME_ABA_DEVOLUCAO);
    if (!abaDevolucao) {
      throw new Error("Aba 'Base Devolução' não foi encontrada na planilha.");
    }

    const todosDados = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, abaDevolucao.getLastColumn()).getValues();

    const INDICES = {
      PEDIDO_ID: 0,
      CLIENTE: 1,
      NFE_NUMERO: 2,
      DATA_NFE: 3,
      PRODUTO: 10,
      QTD_DEVOLVIDA: 17,
      VALOR_DEVOLUCAO: 24,
      MOTIVO: 26,
    };

    const dadosFiltrados = todosDados.filter(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (!dataNfe || !(dataNfe instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });

    const exportData = dadosFiltrados.map(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      const valorRaw = linha[INDICES.VALOR_DEVOLUCAO];
      const valor = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);

      return [
        dataNfe instanceof Date ? Utilities.formatDate(dataNfe, "GMT-3", "dd/MM/yyyy") : 'N/A', // Data NFE
        linha[INDICES.NFE_NUMERO] || '', // NFe_Numero
        linha[INDICES.PEDIDO_ID] || '', // Pedido
        linha[INDICES.CLIENTE] || '', // Cliente
        linha[INDICES.PRODUTO] || '', // Produto
        linha[INDICES.QTD_DEVOLVIDA] || 0, // Qtd Devolvida
        valor, // Valor Devolução
        (linha[INDICES.MOTIVO] || "Não especificado").trim() // Motivo
      ];
    });

    return exportData;

  } catch (e) {
    Logger.log(`Erro em getDevolucaoExportData: ${e.stack}`);
    return { error: e.message };
  }
}

/**
 * NOVA FUNÇÃO: Busca os pedidos de um motivo específico e cruza com outras bases se for cancelamento.
 */
function getPedidosPorMotivo(dateRange, motivo) {
  try {
    const motivoLower = motivo.toLowerCase();
    
    // 1. Busca dados da base de Devolução
    const planilhaDev = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilhaDev.getSheetByName(NOME_ABA_DEVOLUCAO);
    if (!abaDevolucao) throw new Error("Aba 'Base Devolução' não encontrada.");
    const todosDadosDevolucao = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, 27).getValues();

    // 2. Filtra os pedidos da base de devolução pelo motivo e data
    const pedidosFiltrados = {};
    todosDadosDevolucao.forEach(linha => {
      const dataNfe = linha[3]; // DATA_NFE
      if (!dataNfe || !(dataNfe instanceof Date)) return;
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      
      const motivoLinha = (linha[26] || "").toLowerCase().trim(); // MOTIVO
      if (dataFormatada >= dateRange.start && dataFormatada <= dateRange.end && motivoLinha === motivoLower) {
          const pedidoId = linha[0] ? String(linha[0]).trim() : null; // PEDIDO_ID
          if (!pedidoId) return;

          if (!pedidosFiltrados[pedidoId]) {
            pedidosFiltrados[pedidoId] = {
              pedidoId: pedidoId,
              dataNfe: Utilities.formatDate(dataNfe, "GMT-3", "dd/MM/yyyy"),
              cliente: linha[1] || 'N/A', // CLIENTE
              produtos: new Set(),
              justificativa: new Set()
            };
          }
          const produto = linha[10] || 'Produto não especificado'; // PRODUTO
          pedidosFiltrados[pedidoId].produtos.add(produto);
      }
    });

    // 3. Se for 'Cancelamento', busca justificativas em outras bases
    if (motivoLower.includes("cancelamento")) {
        const pedidosParaBuscar = Object.keys(pedidosFiltrados);
        if (pedidosParaBuscar.length > 0) {
            // Busca na planilha de Trocas
            try {
              const planilhaTrocas = SpreadsheetApp.openById(ID_PLANILHA_TROCAS);
              const abaTrocas = planilhaTrocas.getSheetByName(NOME_ABA_TROCAS);
              if (abaTrocas) {
                  const dadosTrocas = abaTrocas.getRange(2, 1, abaTrocas.getLastRow() - 1, 5).getValues();
                  dadosTrocas.forEach(linha => {
                      const pedidoId = linha[0] ? String(linha[0]).trim() : null;
                      if (pedidosFiltrados[pedidoId]) {
                          const motivoTroca = linha[4] ? String(linha[4]).trim() : null; // MOTIVO
                          if (motivoTroca) {
                              pedidosFiltrados[pedidoId].justificativa.add(`Troca: ${motivoTroca}`);
                          }
                      }
                  });
              }
            } catch(e) { Logger.log(`Não foi possível ler a planilha de Trocas: ${e.message}`); }

            // Busca na planilha de Incompatibilidade
            try {
              const planilhaIncompat = SpreadsheetApp.openById(ID_PLANILHA_INCOMPATIBILIDADE);
              const abaIncompat = planilhaIncompat.getSheetByName("Divergência");
              if (abaIncompat) {
                  const dadosIncompat = abaIncompat.getRange(2, 1, abaIncompat.getLastRow() - 1, 11).getValues();
                  dadosIncompat.forEach(linha => {
                      const pedidoId = linha[8] ? String(linha[8]).trim() : null; // PEDIDO
                      if (pedidosFiltrados[pedidoId]) {
                          const tipoProblema = linha[10] ? String(linha[10]).trim() : null; // TIPO_PROBLEMA
                          if (tipoProblema) {
                              pedidosFiltrados[pedidoId].justificativa.add(`Incompat.: ${tipoProblema}`);
                          }
                      }
                  });
              }
            } catch(e) { Logger.log(`Não foi possível ler a planilha de Incompatibilidade: ${e.message}`); }
        }
    }

    // 4. Formata o resultado final
    return Object.values(pedidosFiltrados).map(p => {
        return {
            pedidoId: p.pedidoId,
            dataNfe: p.dataNfe,
            cliente: p.cliente,
            produtos: Array.from(p.produtos), // <-- ALTERADO AQUI
            justificativa: p.justificativa.size > 0 ? Array.from(p.justificativa).join('; ') : 'Não encontrada'
        };
    }).sort((a, b) => { // Ordena por data
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
function getDevolucaoDataWithCache(dateRange) {
  const cacheKey = `devolucao_data_v14_period_fix_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getDevolucaoData, [dateRange]);
}

function getItensDevolvidosPorPedidoWithCache(pedidoId) {
  const cacheKey = `devolucao_itens_pedido_v2_${pedidoId}`;
  return getOrSetCache(cacheKey, getItensDevolvidosPorPedido, [pedidoId], 300); // cache de 5 minutos
}

function getDevolucaoExportDataWithCache(dateRange) {
  // REMOVIDO: Cache para a função de exportação para evitar o erro "Argumento grande demais".
  // A exportação sempre buscará os dados mais recentes diretamente.
  Logger.log(`Buscando dados de exportação (sem cache) para o período ${dateRange.start} a ${dateRange.end}`);
  return getDevolucaoExportData(dateRange);
}

function getPedidosPorMotivoWithCache(dateRange, motivo) {
  const cacheKey = `pedidos_por_motivo_v3_${dateRange.start}_${dateRange.end}_${motivo.replace(/\s+/g, '_')}`;
  return getOrSetCache(cacheKey, getPedidosPorMotivo, [dateRange, motivo]);
}
