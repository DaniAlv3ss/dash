/**
 * Contém todas as funções do lado do servidor para a página do Dashboard de Devolução.
 */

/**
 * Busca e processa todos os dados necessários para os KPIs e gráficos do dashboard de devolução.
 * MODIFICADO: Lógica de contagem refeita para usar a coluna NFE_NUMERO (C) como chave primária,
 * garantindo a contagem correta de devoluções únicas.
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
      NFE_NUMERO: 2,      // Coluna C
      DATA_NFE: 3,        // Coluna D
      CF_PRODUTO: 9,      // Coluna J
      PRODUTO: 10,        // Coluna K
      CATEGORIA_BUDGET: 12, // Coluna M
      QTD_DEVOLVIDA: 17,  // Coluna R
      VALOR_DEVOLUCAO: 24,// Coluna Y
      MOTIVO: 26,         // Coluna AA
      FABRICANTE: 28,     // Coluna AC
      PEDIDO_NFD_UNICO: 38 // Coluna AM (Não utilizado para contagem de NFs)
    };
    
    // 1. Processa dados do ano inteiro para o gráfico de evolução mensal
    const anoCorrente = new Date().getFullYear();
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
    for(let i=0; i<12; i++) {
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
    const startDate = new Date(dateRange.start.replace(/-/g, '/'));
    const endDate = new Date(dateRange.end.replace(/-/g, '/'));
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
    
    // 4. Processa os dados filtrados para KPIs, listas de Top 10 e tabelas
    let totalDevolvido = 0;
    let valorCancelamento = 0;
    const nfsUnicas = new Set();
    const motivos = {}; // Estrutura: { motivo: { nfs: Set(), valor: 0 } }
    const categorias = {};
    const fabricantes = {};
    let totalItensDevolvidos = 0;
    const tabelaDetalhada = {};
    const todosMotivos = new Set();

    dadosFiltrados.forEach(linha => {
      const qtdDevolvida = parseInt(linha[INDICES.QTD_DEVOLVIDA], 10) || 0;
      if (qtdDevolvida === 0) return; // Pula a linha se nenhum item foi devolvido

      const motivo = (linha[INDICES.MOTIVO] || "Não especificado").trim().toLowerCase();
      const valorRaw = linha[INDICES.VALOR_DEVOLUCAO];
      const valor = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);
      
      const rawNf = linha[INDICES.NFE_NUMERO];
      const nfUnica = (rawNf !== null && rawNf !== undefined && rawNf !== '') ? String(rawNf).trim() : null;

      // --- Cálculos de KPIs ---
      totalDevolvido += valor;
      if (motivo.includes("cancelamento")) {
          valorCancelamento += valor;
      }
      if(nfUnica) {
        nfsUnicas.add(nfUnica);
      }
      totalItensDevolvidos += qtdDevolvida;

      // --- Agregação para Top Motivos (por NF-e) ---
      if (!motivos[motivo]) {
        motivos[motivo] = { nfs: new Set(), valor: 0 };
      }
      if (nfUnica) {
        motivos[motivo].nfs.add(nfUnica);
      }
      motivos[motivo].valor += valor;

      // --- Agregação para Top Categorias e Fabricantes (por item) ---
      const fabricante = (linha[INDICES.FABRICANTE] || "Não especificado").trim();
      const categoriaBudget = (linha[INDICES.CATEGORIA_BUDGET] || "Não especificado").trim();
      
      categorias[categoriaBudget] = (categorias[categoriaBudget] || 0) + qtdDevolvida;
      fabricantes[fabricante] = (fabricantes[fabricante] || 0) + qtdDevolvida;
      
      // --- Agregação para a tabela detalhada de produtos ---
      const displayMotivo = motivo.charAt(0).toUpperCase() + motivo.slice(1);
      todosMotivos.add(displayMotivo);

      const cfProduto = linha[INDICES.CF_PRODUTO] || 'N/A';
      const produtoNome = (linha[INDICES.PRODUTO] || "Não especificado").trim(); 
      const produtoKey = `${cfProduto}|${produtoNome}`;
      if(!tabelaDetalhada[produtoKey]){
        tabelaDetalhada[produtoKey] = { cf: cfProduto, nome: produtoNome, total: 0, motivos: {} };
      }
      tabelaDetalhada[produtoKey].total += qtdDevolvida;
      tabelaDetalhada[produtoKey].motivos[displayMotivo] = (tabelaDetalhada[produtoKey].motivos[displayMotivo] || 0) + qtdDevolvida;
    });
    
    // 5. Formata os dados agregados para o front-end
    const topMotivos = Object.entries(motivos)
        .map(([motivo, data]) => {
            const displayMotivo = motivo.charAt(0).toUpperCase() + motivo.slice(1);
            const quantidade = data.nfs.size;
            const valor = data.valor;
            const quantidadeAnterior = motivosAnteriores[motivo] ? motivosAnteriores[motivo].size : 0;
            return [displayMotivo, { quantidade, valor, quantidadeAnterior }];
        })
        .sort(([, a], [, b]) => b.quantidade - a.quantidade)
        .slice(0, 4);

    const topCategorias = Object.entries(categorias).sort(([,a],[,b]) => b-a).slice(0, 10);
    const topFabricantes = Object.entries(fabricantes).sort(([,a],[,b]) => b-a).slice(0, 10);

    return {
      kpis: {
        totalDevolvido: totalDevolvido,
        valorCancelamento: valorCancelamento,
        devolucoesUnicas: nfsUnicas.size,
        totalItensDevolvidos: totalItensDevolvidos
      },
      topMotivos: topMotivos,
      topCategorias: topCategorias,
      topFabricantes: topFabricantes,
      devolucoesPorMes: chartData,
      tabelaDetalhada: {
        motivos: Array.from(todosMotivos).sort(),
        produtos: Object.values(tabelaDetalhada).sort((a,b) => b.total - a.total)
      }
    };

  } catch (e) {
    Logger.log(`Erro em getDevolucaoData: ${e.stack}`);
    return { error: e.message };
  }
}

// --- VERSÃO COM CACHE PARA SER CHAMADA PELO CLIENTE ---
function getDevolucaoDataWithCache(dateRange) {
  // Versão do cache incrementada para invalidar dados antigos após a mudança de lógica
  const cacheKey = `devolucao_data_v8_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getDevolucaoData, [dateRange]);
}

