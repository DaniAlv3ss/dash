/**
 * Contém todas as funções do lado do servidor para a página de Incompatibilidades.
 */

/**
 * Busca e processa dados usando a API Google Visualization para performance otimizada.
 * A query filtra os dados por data diretamente na planilha, transferindo apenas
 * os dados relevantes para o script, o que acelera drasticamente o carregamento.
 */
function getIncompatibilidadeData(dateRange) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_INCOMPATIBILIDADE);
    const aba = planilha.getSheetByName("Divergência");
    if (!aba) throw new Error("Aba 'Divergência' não foi encontrada na planilha.");

    // INDICES REVISADOS CONFORME SOLICITAÇÃO DE MUDANÇA
    const INDICES = {
      CARIMBO_DATA: 0, NOME_TECNICO: 1, PROCESSO: 2, QUAL_PROBLEMA: 3,
      DETALHE_PROBLEMA: 4, OS: 5, INICIO_TRATATIVA: 6, FIM_TRATATIVA: 7,
      PEDIDO: 8, PRODUTO_DIVERGENTE: 9, TIPO_PROBLEMA: 10, AVARIA_CUSTOMIZA: 11,
      TIPO_TROCA: 12, DATA_LIBERACAO: 13, ETIQUETA_PRODUTO_ANTIGO: 14,
      CF_PRODUTO_ANTIGO: 15, CFS_INCOMPATIVEIS: 16, CF_PRODUTO_NOVO: 17,
      ETIQUETA_PRODUTO_NOVO: 18, STATUS_MONTAGEM: 19, STATUS_SISTEMA: 20,
      OBSERVACAO: 21, PEDIDO_COMPLEMENTAR: 22, CUSTO_CF_ANTIGO: 23,
      CUSTO_CF_NOVO: 24, CUSTO_TROCAS: 25, FABRICANTE: 26, NOME_RESPONSAVEL: 27
    };
    
    // 1. Construir a Query
    // Seleciona todas as colunas (A-AB) onde a data na coluna A está dentro do intervalo.
    // Ordena pela data de forma decrescente para facilitar a lógica de unicidade.
    const query = `SELECT * WHERE A >= date '${dateRange.start}' AND A <= date '${dateRange.end}' ORDER BY A DESC`;
    
    // 2. Fazer a Requisição via API de Visualização
    const url = `https://docs.google.com/spreadsheets/d/${ID_PLANILHA_INCOMPATIBILIDADE}/gviz/tq?tq=${encodeURIComponent(query)}&gid=${aba.getSheetId()}`;
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
    });

    // A API retorna um JSONP. Precisamos limpá-lo para obter o JSON puro.
    const jsonText = response.getContentText().replace("/*O_o*/\ngoogle.visualization.Query.setResponse(", "").replace(");", "");
    const queryResult = JSON.parse(jsonText);
    
    if (queryResult.status === 'error') {
      throw new Error(`Erro na query: ${queryResult.errors[0].detailed_message}`);
    }
    
    // 3. Parsear a Resposta da API para um formato de array simples
    const dadosBrutos = (queryResult.table.rows || []).map(row => 
      (row.c || []).map(cell => {
        if (!cell) return null;
        if (cell.v instanceof Date || (typeof cell.v === 'string' && cell.v.startsWith('Date('))) {
          // Extrai os valores numéricos da string "Date(y,m,d,h,m,s)"
          const params = String(cell.v).match(/\d+/g);
          if (params) {
            // new Date(year, month(0-11), day, hours, minutes, seconds)
            return new Date(params[0], params[1], params[2], params[3] || 0, params[4] || 0, params[5] || 0);
          }
        }
        return cell.v; // Retorna o valor (string, number, boolean)
      })
    );

    // 4. Processar os dados já filtrados e ordenados
    const osProcessadas = new Set();
    let custoTotalTrocas = 0;
    let tempoTotalTratativa = 0;
    let tratativasCompletas = 0;
    const contagemTecnicos = {};
    const contagemProblemas = {};
    const contagemFabricantes = {};
    const tabelaDetalhada = [];

    dadosBrutos.forEach(linha => {
      const os = linha[INDICES.OS] ? String(linha[INDICES.OS]).trim() : null;
      // Como os dados já estão ordenados do mais novo para o mais antigo, a primeira vez
      // que vemos uma OS, ela é a mais recente.
      if (!os || os === "" || osProcessadas.has(os)) return;
      
      osProcessadas.add(os);

      // CORREÇÃO: Lógica de conversão de moeda ajustada para o formato brasileiro (R$ 1.234,56)
      const custoRaw = linha[INDICES.CUSTO_TROCAS];
      let custo = 0;
      if (typeof custoRaw === 'number') {
          custo = custoRaw;
      } else if (typeof custoRaw === 'string' && custoRaw.trim() !== '') {
          const custoLimpo = custoRaw.replace("R$", "").trim().replace(/\./g, "").replace(",", ".");
          custo = parseFloat(custoLimpo) || 0;
      }
      custoTotalTrocas += custo;

      const inicio = linha[INDICES.INICIO_TRATATIVA];
      const fim = linha[INDICES.FIM_TRATATIVA];
      if (inicio instanceof Date && fim instanceof Date) {
        const diffHoras = (fim.getTime() - inicio.getTime()) / (1000 * 60 * 60);
        if (diffHoras >= 0) {
            tempoTotalTratativa += diffHoras;
            tratativasCompletas++;
        }
      }

      const tecnico = (linha[INDICES.NOME_TECNICO] || "Não especificado").trim();
      const problema = (linha[INDICES.TIPO_PROBLEMA] || "Não especificado").trim();
      const fabricante = (linha[INDICES.FABRICANTE] || "Não especificado").trim();
      
      if(tecnico && tecnico !== "Não especificado") contagemTecnicos[tecnico] = (contagemTecnicos[tecnico] || 0) + 1;
      if(problema && problema !== "Não especificado") contagemProblemas[problema] = (contagemProblemas[problema] || 0) + 1;
      if(fabricante && fabricante !== "Não especificado") contagemFabricantes[fabricante] = (contagemFabricantes[fabricante] || 0) + 1;
      
      tabelaDetalhada.push({
        data: inicio instanceof Date ? Utilities.formatDate(inicio, "GMT-3", "dd/MM/yyyy HH:mm") : 'N/A',
        tecnico: tecnico, os: linha[INDICES.OS], pedido: linha[INDICES.PEDIDO],
        produto: linha[INDICES.PRODUTO_DIVERGENTE], tipoProblema: problema,
        detalheProblema: linha[INDICES.DETALHE_PROBLEMA]
      });
    });

    const tempoMedioTratativa = tratativasCompletas > 0 ? (tempoTotalTratativa / tratativasCompletas) : 0;
    const topTecnicos = Object.entries(contagemTecnicos).sort(([,a],[,b]) => b-a).slice(0, 5);
    const topProblemas = Object.entries(contagemProblemas).sort(([,a],[,b]) => b-a).slice(0, 5);
    const topFabricantes = Object.entries(contagemFabricantes).sort(([,a],[,b]) => b-a).slice(0, 10);
    
    // A tabela já está ordenada pela query, mas se precisar reordenar por data de início:
    tabelaDetalhada.sort((a, b) => {
        const dateA = a.data !== 'N/A' ? new Date(a.data.split(' ')[0].split('/').reverse().join('-') + ' ' + a.data.split(' ')[1]) : new Date(0);
        const dateB = b.data !== 'N/A' ? new Date(b.data.split(' ')[0].split('/').reverse().join('-') + ' ' + b.data.split(' ')[1]) : new Date(0);
        return dateB - dateA;
    });

    return {
      kpis: { totalIncompatibilidades: osProcessadas.size, custoTotal: custoTotalTrocas, tempoMedioHoras: tempoMedioTratativa },
      topTecnicos: topTecnicos, topProblemas: topProblemas, topFabricantes: topFabricantes, tabela: tabelaDetalhada
    };

  } catch (e) {
    Logger.log(`Erro em getIncompatibilidadeData: ${e.stack}`);
    return { error: e.message };
  }
}

/**
 * Versão com cache da função de busca de dados.
 */
function getIncompatibilidadeDataWithCache(dateRange) {
  const cacheKey = `incompatibilidade_data_v6_query_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getIncompatibilidadeData, [dateRange]);
}

