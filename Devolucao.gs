/**
 * Contém todas as funções do lado do servidor para a página do Dashboard de Devolução.
 */

/**
 * Busca e processa todos os dados necessários para os KPIs e gráficos do dashboard de devolução.
 */
function getDevolucaoData(dateRange) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_DEVOLUCAO);
    const abaDevolucao = planilha.getSheetByName(NOME_ABA_DEVOLUCAO);
    if (!abaDevolucao) {
      throw new Error("Aba 'Base Devolução' não foi encontrada na planilha.");
    }

    const todosDados = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, abaDevolucao.getLastColumn()).getValues();

    // Mapeamento de colunas (ajuste os números se a ordem das suas colunas for diferente)
    const INDICES = {
      PEDIDO_ID: 0,       // Coluna A
      NFE_NUMERO: 2,      // Coluna C
      DATA_NFE: 3,        // Coluna D
      CF_PRODUTO: 9,      // Coluna J
      PRODUTO: 10,        // Coluna K
      QTD_DEVOLVIDA: 17,  // Coluna R
      VALOR_DEVOLUCAO: 24,// Coluna Y
      MOTIVO: 26,         // Coluna AA
      FABRICANTE: 28,     // Coluna AC
      PEDIDO_NFD_UNICO: 38 // Coluna AM
    };
    
    // Processa dados do ano inteiro para o gráfico
    const anoCorrente = new Date().getFullYear();
    const devolucoesPorMes = {};
    const monthNames = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
    const pedidosUnicosPorMes = {};

    todosDados.forEach(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (dataNfe instanceof Date && dataNfe.getFullYear() === anoCorrente) {
        const valorRaw = linha[INDICES.VALOR_DEVOLUCAO];
        const valor = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);
        const pedidoUnico = linha[INDICES.PEDIDO_NFD_UNICO] || null;
        
        const month = dataNfe.getMonth();
        const year = dataNfe.getFullYear();
        const key = `${year}-${String(month).padStart(2, '0')}`;
        const monthLabel = `${monthNames[month]}/${String(year).slice(-2)}`;
        
        if (!devolucoesPorMes[key]) {
          devolucoesPorMes[key] = { label: monthLabel, valor: 0 };
          pedidosUnicosPorMes[key] = new Set();
        }
        
        devolucoesPorMes[key].valor += valor;
        if(pedidoUnico) {
          pedidosUnicosPorMes[key].add(pedidoUnico);
        }
      }
    });

    const chartData = [['Mês', 'Valor Devolvido (R$)', 'Pedidos Únicos']];
    for(let i=0; i<12; i++) {
        const key = `${anoCorrente}-${String(i).padStart(2, '0')}`;
        const monthLabel = `${monthNames[i]}/${String(anoCorrente).slice(-2)}`;
        const valor = devolucoesPorMes[key] ? devolucoesPorMes[key].valor : 0;
        const quantidade = pedidosUnicosPorMes[key] ? pedidosUnicosPorMes[key].size : 0;
        chartData.push([monthLabel, valor, quantidade]);
    }


    // Filtra dados para KPIs e listas com base no dateRange
    const dadosFiltrados = todosDados.filter(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (!dataNfe || !(dataNfe instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });

    let totalDevolvido = 0;
    let valorCancelamento = 0;
    const pedidosUnicos = new Set();
    const motivos = {};
    const produtos = {};
    const fabricantes = {};
    let totalItensDevolvidos = 0;
    const tabelaDetalhada = {};
    const todosMotivos = new Set();

    dadosFiltrados.forEach(linha => {
      const motivo = linha[INDICES.MOTIVO] || "Não especificado";
      const valorRaw = linha[INDICES.VALOR_DEVOLUCAO];
      const valor = (typeof valorRaw === 'number') ? valorRaw : (parseFloat(String(valorRaw).replace(/[R$\s.]/g, '').replace(',', '.')) || 0);

      if (motivo.toLowerCase().includes("cancelamento")) {
          valorCancelamento += valor;
      }
      totalDevolvido += valor;
      
      pedidosUnicos.add(linha[INDICES.PEDIDO_ID]);

      const produto = linha[INDICES.PRODUTO] || "Não especificado";
      const fabricante = linha[INDICES.FABRICANTE] || "Não especificado";
      const qtdDevolvida = parseInt(linha[INDICES.QTD_DEVOLVIDA], 10) || 0;
      const cfProduto = linha[INDICES.CF_PRODUTO] || 'N/A';

      totalItensDevolvidos += qtdDevolvida;
      motivos[motivo] = (motivos[motivo] || 0) + qtdDevolvida;
      produtos[produto] = (produtos[produto] || 0) + qtdDevolvida;
      fabricantes[fabricante] = (fabricantes[fabricante] || 0) + qtdDevolvida;

      if(qtdDevolvida > 0){
        todosMotivos.add(motivo);
        const produtoKey = `${cfProduto}|${produto}`;
        if(!tabelaDetalhada[produtoKey]){
          tabelaDetalhada[produtoKey] = {
            cf: cfProduto,
            nome: produto,
            total: 0,
            motivos: {}
          };
        }
        tabelaDetalhada[produtoKey].total += qtdDevolvida;
        tabelaDetalhada[produtoKey].motivos[motivo] = (tabelaDetalhada[produtoKey].motivos[motivo] || 0) + qtdDevolvida;
      }

    });
    
    const topMotivos = Object.entries(motivos).sort(([,a],[,b]) => b-a).slice(0, 10);
    const topProdutos = Object.entries(produtos).sort(([,a],[,b]) => b-a).slice(0, 10);
    const topFabricantes = Object.entries(fabricantes).sort(([,a],[,b]) => b-a).slice(0, 10);

    return {
      kpis: {
        totalDevolvido: totalDevolvido,
        valorCancelamento: valorCancelamento,
        pedidosComDevolucao: pedidosUnicos.size,
        totalItensDevolvidos: totalItensDevolvidos
      },
      topMotivos: topMotivos,
      topProdutos: topProdutos,
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
  const cacheKey = `devolucao_data_v3_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getDevolucaoData, [dateRange]);
}

