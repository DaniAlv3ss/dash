/**
 * Contém todas as funções do lado do servidor para a página do Dashboard de Devolução.
 */

// IMPORTANTE: Substitua "ID_DA_SUA_NOVA_PLANILHA" pelo ID da sua planilha do Google com os dados de devolução.
const ID_PLANILHA_DEVOLUCAO = "1m3tOvmSOJIvRZY9uZNf1idSTEnUFbHIWPNh5tiHkKe0";
const NOME_ABA_DEVOLUCAO = "Base Devolução"; // Confirme se este é o nome da aba na sua planilha

/**
 * Helper para converter com segurança um valor da planilha para um objeto Date.
 * @param {*} value O valor da célula.
 * @returns {Date|null} Um objeto Date se a conversão for bem-sucedida, caso contrário null.
 */
function safeParseDate(value) {
  if (value instanceof Date && !isNaN(value)) {
    return value;
  }
  if (typeof value === 'string' || typeof value === 'number') {
    const d = new Date(value);
    if (d instanceof Date && !isNaN(d)) {
      return d;
    }
  }
  return null;
}

/**
 * Helper para converter com segurança um valor da planilha para um número (float).
 * Lida com formatos numéricos e monetários em string (ex: "R$ 1.234,56").
 * @param {*} value O valor da célula.
 * @returns {number} O valor numérico convertido. Retorna 0 se a conversão falhar.
 */
function safeParseFloat(value) {
    if (typeof value === 'number' && !isNaN(value)) {
        return value;
    }
    if (typeof value === 'string' && value.trim() !== '') {
        const cleanedString = value.replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
        const number = parseFloat(cleanedString);
        return isNaN(number) ? 0 : number;
    }
    return 0;
}


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

    // Usar getValues() para obter os objetos Date corretamente
    const todosDados = abaDevolucao.getRange(2, 1, abaDevolucao.getLastRow() - 1, abaDevolucao.getLastColumn()).getValues();

    // Mapeamento de colunas
    const INDICES = {
      CF_PRODUTO: 9,        // Coluna J
      PRODUTO: 10,        // Coluna K
      VALOR_DEVOLUCAO: 24,// Coluna Y
      MOTIVO: 26,         // Coluna AA
      FABRICANTE: 28,     // Coluna AC
      PEDIDO_UNICO_DEVOLUCAO: 38 // Coluna AM
    };

    const dadosFiltrados = todosDados.filter(linha => {
      // A data de referência para o filtro de período continua sendo a da NFe da Coluna D
      const dataReferenciaFiltro = linha[3];
      if (!dataReferenciaFiltro || !(dataReferenciaFiltro instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataReferenciaFiltro, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });

    let totalDevolvido = 0;
    let totalCancelado = 0;
    const motivos = {};
    const produtos = {};
    const fabricantes = {};
    const pedidosUnicosFiltrados = new Set();
    
    // Para o gráfico anual
    const devolucoesPorMes = {};
    const pedidosUnicosMes = {}; 
    const monthNames = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

    // Para a nova tabela detalhada
    const produtosDetalhado = {};
    const todosOsMotivos = new Set();


    // Processa os dados filtrados pelo período para os KPIs
    dadosFiltrados.forEach(linha => {
      pedidosUnicosFiltrados.add(linha[INDICES.PEDIDO_UNICO_DEVOLUCAO]);
      
      const valor = safeParseFloat(linha[INDICES.VALOR_DEVOLUCAO]);
      totalDevolvido += valor;
      
      const motivo = linha[INDICES.MOTIVO] || "Não especificado";
      const motivoLowerCase = motivo.toLowerCase();
      if (motivoLowerCase.includes('cancelamento') || motivoLowerCase.includes('arrependimento')) {
          totalCancelado += valor;
      }

      motivos[motivo] = (motivos[motivo] || 0) + 1;

      const produto = linha[INDICES.PRODUTO] || "Não especificado";
      produtos[produto] = (produtos[produto] || 0) + 1;

      const fabricante = linha[INDICES.FABRICANTE] || "Não especificado";
      fabricantes[fabricante] = (fabricantes[fabricante] || 0) + 1;

      // Lógica para a nova tabela detalhada
      const cf = linha[INDICES.CF_PRODUTO] || 'N/A';
      const nomeProduto = linha[INDICES.PRODUTO] || 'Não especificado';
      
      todosOsMotivos.add(motivo);

      if (!produtosDetalhado[cf]) {
        produtosDetalhado[cf] = {
          cf: cf,
          nome: nomeProduto,
          total: 0,
          motivos: {}
        };
      }
      produtosDetalhado[cf].total++;
      produtosDetalhado[cf].motivos[motivo] = (produtosDetalhado[cf].motivos[motivo] || 0) + 1;
    });

    // Processa todos os dados do ano para o gráfico
    const anoCorrente = new Date().getFullYear();
    todosDados.forEach(linha => {
      const dataReferenciaGrafico = linha[3]; // Coluna D para o eixo do tempo do gráfico
      if (dataReferenciaGrafico instanceof Date && dataReferenciaGrafico.getFullYear() === anoCorrente) {
        const month = dataReferenciaGrafico.getMonth();
        const year = dataReferenciaGrafico.getFullYear();
        const key = `${year}-${String(month).padStart(2, '0')}`;
        const monthLabel = `${monthNames[month]}/${year}`;

        const valor = safeParseFloat(linha[INDICES.VALOR_DEVOLUCAO]);
        const pedidoUnicoDevolucao = linha[INDICES.PEDIDO_UNICO_DEVOLUCAO];

        if (!devolucoesPorMes[key]) {
          devolucoesPorMes[key] = { label: monthLabel, valor: 0 };
          pedidosUnicosMes[key] = new Set();
        }
        
        devolucoesPorMes[key].valor += valor;

        if(pedidoUnicoDevolucao) {
            pedidosUnicosMes[key].add(pedidoUnicoDevolucao);
        }
      }
    });

    // Monta os dados para o gráfico
    const chartData = [['Mês', 'Valor Devolvido (R$)', 'Pedidos Únicos']];
    monthNames.forEach((name, index) => {
        const year = new Date().getFullYear();
        const key = `${year}-${String(index).padStart(2, '0')}`;
        if(devolucoesPorMes[key]){
            const item = devolucoesPorMes[key];
            const qtdPedidos = pedidosUnicosMes[key] ? pedidosUnicosMes[key].size : 0;
            chartData.push([item.label, item.valor, qtdPedidos]);
        } else {
            chartData.push([`${name}/${year}`, 0, 0]);
        }
    });

    const topMotivos = Object.entries(motivos).sort(([,a],[,b]) => b-a).slice(0, 10);
    const topProdutos = Object.entries(produtos).sort(([,a],[,b]) => b-a).slice(0, 10);
    const topFabricantes = Object.entries(fabricantes).sort(([,a],[,b]) => b-a).slice(0, 10);

    // Prepara dados da tabela detalhada para o retorno
    const cabecalhosMotivos = Array.from(todosOsMotivos).sort();
    const linhasProdutos = Object.values(produtosDetalhado).sort((a, b) => b.total - a.total);

    const devolucaoDetalhada = {
        cabecalhos: cabecalhosMotivos,
        linhas: linhasProdutos
    };

    return {
      kpis: {
        totalDevolvido: totalDevolvido,
        totalCancelado: totalCancelado,
        pedidosComDevolucao: pedidosUnicosFiltrados.size,
        totalItensDevolvidos: dadosFiltrados.length
      },
      topMotivos: topMotivos,
      topProdutos: topProdutos,
      topFabricantes: topFabricantes,
      devolucoesPorMes: chartData,
      devolucaoDetalhada: devolucaoDetalhada
    };

  } catch (e) {
    Logger.log(`Erro em getDevolucaoData: ${e.stack}`);
    return { error: e.message };
  }
}

