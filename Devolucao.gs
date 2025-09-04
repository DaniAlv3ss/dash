/**
 * Contém todas as funções do lado do servidor para a página do Dashboard de Devolução.
 */

// IMPORTANTE: Substitua "ID_DA_SUA_NOVA_PLANILHA" pelo ID da sua planilha do Google com os dados de devolução.
const ID_PLANILHA_DEVOLUCAO = "1m3tOvmSOJIvRZY9uZNf1idSTEnUFbHIWPNh5tiHkKe0";
const NOME_ABA_DEVOLUCAO = "Base Devolução"; // Confirme se este é o nome da aba na sua planilha

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
      BUDGET_ID: 12,      // Coluna M
      QTD_DEVOLVIDA: 17,  // Coluna R
      VALOR_DEVOLUCAO: 24,// Coluna Y
      MOTIVO: 26,         // Coluna AA
      FABRICANTE: 28      // Coluna AC
    };

    const dadosFiltrados = todosDados.filter(linha => {
      const dataNfe = linha[INDICES.DATA_NFE];
      if (!dataNfe || !(dataNfe instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataNfe, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });

    let totalDevolvido = 0;
    const pedidosUnicos = new Set();
    const motivos = {};
    const produtos = {};
    const fabricantes = {};
    const devolucoesPorMes = {};
    const monthNames = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];


    dadosFiltrados.forEach(linha => {
      const pedidoId = linha[INDICES.PEDIDO_ID];
      const valorStr = (linha[INDICES.VALOR_DEVOLUCAO] || '0').toString().replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
      const valor = parseFloat(valorStr);
      const motivo = linha[INDICES.MOTIVO] || "Não especificado";
      const produto = linha[INDICES.PRODUTO] || "Não especificado";
      const fabricante = linha[INDICES.FABRICANTE] || "Não especificado";
      const dataNfe = linha[INDICES.DATA_NFE];

      if (!isNaN(valor)) {
        totalDevolvido += valor;
      }
      pedidosUnicos.add(pedidoId);

      motivos[motivo] = (motivos[motivo] || 0) + 1;
      produtos[produto] = (produtos[produto] || 0) + 1;
      fabricantes[fabricante] = (fabricantes[fabricante] || 0) + 1;
      
      if (dataNfe instanceof Date) {
        const month = dataNfe.getMonth();
        const year = dataNfe.getFullYear();
        const key = `${year}-${String(month).padStart(2, '0')}`;
        const monthLabel = `${monthNames[month]}/${year}`;

        if (!devolucoesPorMes[key]) {
          devolucoesPorMes[key] = { label: monthLabel, valor: 0, quantidade: 0 };
        }
        devolucoesPorMes[key].valor += valor;
        devolucoesPorMes[key].quantidade++;
      }
    });

    const topMotivos = Object.entries(motivos).sort(([,a],[,b]) => b-a).slice(0, 10);
    const topProdutos = Object.entries(produtos).sort(([,a],[,b]) => b-a).slice(0, 10);
    const topFabricantes = Object.entries(fabricantes).sort(([,a],[,b]) => b-a).slice(0, 10);
    
    const sortedMeses = Object.keys(devolucoesPorMes).sort();
    const chartData = [['Mês', 'Valor Devolvido (R$)', 'Quantidade']];
    sortedMeses.forEach(key => {
        const item = devolucoesPorMes[key];
        chartData.push([item.label, item.valor, item.quantidade]);
    });

    return {
      kpis: {
        totalDevolvido: totalDevolvido,
        pedidosComDevolucao: pedidosUnicos.size,
        totalItensDevolvidos: dadosFiltrados.length
      },
      topMotivos: topMotivos,
      topProdutos: topProdutos,
      topFabricantes: topFabricantes,
      devolucoesPorMes: chartData
    };

  } catch (e) {
    Logger.log(`Erro em getDevolucaoData: ${e.stack}`);
    return { error: e.message };
  }
}
