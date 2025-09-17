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
 * HELPER: Executa uma query na API de Visualização e retorna a tabela de dados.
 */
function executeQuery_(sheet, query) {
  const url = `https://docs.google.com/spreadsheets/d/${ID_PLANILHA_INCOMPATIBILIDADE}/gviz/tq?tq=${encodeURIComponent(query)}&gid=${sheet.getSheetId()}`;
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  const jsonText = response.getContentText().replace("/*O_o*/\ngoogle.visualization.Query.setResponse(", "").replace(");", "");
  const queryResult = JSON.parse(jsonText);

  if (queryResult.status === 'error') {
    throw new Error(`Erro na query: ${queryResult.errors[0].detailed_message}`);
  }

  return (queryResult.table.rows || []).map(row => 
    (row.c || []).map(cell => {
      if (!cell) return null;
      if (cell.v instanceof Date || (typeof cell.v === 'string' && cell.v.startsWith('Date('))) {
        const params = String(cell.v).match(/\d+/g);
        if (params) return new Date(params[0], params[1], params[2], params[3] || 0, params[4] || 0, params[5] || 0);
      }
      return cell.f || cell.v; // Prefere valor formatado se existir
    })
  );
}


/**
 * OTIMIZADO: Busca todos os dados de um técnico com uma única query e processa no servidor para garantir consistência.
 */
function getIncompatibilidadeDataForTechnician(dateRange, tecnicoNome) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_INCOMPATIBILIDADE);
    const aba = planilha.getSheetByName("Divergência");
    if (!aba) throw new Error("Aba 'Divergência' não foi encontrada.");
    
    // MODIFICAÇÃO: Define o range para o ano inteiro, do início do ano até a data final do período selecionado.
    const currentYear = new Date(dateRange.end).getFullYear();
    const yearStartDate = `${currentYear}-01-01`;

    const baseWhereClause = `WHERE B = '${tecnicoNome}' AND A IS NOT NULL AND A >= date '${yearStartDate}' AND A <= date '${dateRange.end}'`;
    const queryCasos = `SELECT A, E, F, I, J, K ${baseWhereClause} ORDER BY A DESC`;
    const dadosBrutosDoAno = executeQuery_(aba, queryCasos);

    // --- Processamento ÚNICO para criar uma fonte de dados limpa e desduplicada ---
    const osProcessadas = new Set();
    const casosDetalhadosAno = []; // Armazena todos os casos únicos do ano
    
    dadosBrutosDoAno.forEach(linha => {
      const os = linha[2] ? String(linha[2]).trim() : null;
      if (!os || osProcessadas.has(os)) return; // Desduplicação CRÍTICA
      osProcessadas.add(os);

      const data = linha[0];
      if (data instanceof Date) {
        casosDetalhadosAno.push({
          data: data, // Mantém como objeto Date para facilitar o processamento
          detalheProblema: linha[1],
          os: os,
          pedido: linha[3],
          produto: linha[4],
          tipoProblema: (linha[5] || "Não especificado").trim()
        });
      }
    });

    // --- Agora, cria as agregações a partir dos dados limpos do ano ---
    const monthNames = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
    const frequenciaMensalAnual = {};
    const problemasPorMes = {};
    const contagemProblemasPeriodo = {};
    const frequenciaDiariaPeriodo = {};
    const casosTabelaPeriodo = [];
    let totalCasosPeriodo = 0;

    const startDatePeriodo = new Date(dateRange.start + 'T00:00:00');
    const endDatePeriodo = new Date(dateRange.end + 'T23:59:59');

    casosDetalhadosAno.forEach(caso => {
      const dataCaso = caso.data;
      
      // Agregação Anual
      const ano = dataCaso.getFullYear();
      const mesIndex = dataCaso.getMonth();
      const mesLabel = `${monthNames[mesIndex]}/${String(ano).slice(-2)}`;
      
      frequenciaMensalAnual[mesLabel] = (frequenciaMensalAnual[mesLabel] || 0) + 1;

      if (!problemasPorMes[mesLabel]) {
        problemasPorMes[mesLabel] = {};
      }
      problemasPorMes[mesLabel][caso.tipoProblema] = (problemasPorMes[mesLabel][caso.tipoProblema] || 0) + 1;
      
      // Agregação do Período Selecionado
      if (dataCaso >= startDatePeriodo && dataCaso <= endDatePeriodo) {
        totalCasosPeriodo++;
        contagemProblemasPeriodo[caso.tipoProblema] = (contagemProblemasPeriodo[caso.tipoProblema] || 0) + 1;
        
        const diaKey = Utilities.formatDate(dataCaso, "GMT-3", "dd/MM");
        frequenciaDiariaPeriodo[diaKey] = (frequenciaDiariaPeriodo[diaKey] || 0) + 1;

        casosTabelaPeriodo.push({
          data: Utilities.formatDate(dataCaso, "GMT-3", "dd/MM/yyyy HH:mm"),
          os: caso.os,
          pedido: caso.pedido,
          produto: caso.produto,
          tipoProblema: caso.tipoProblema,
          detalheProblema: caso.detalheProblema
        });
      }
    });

    // --- Formatação final dos dados para o front-end ---
    const dataTableProblemasPeriodo = [['Problema', 'Quantidade'], ...Object.entries(contagemProblemasPeriodo).sort(([, a], [, b]) => b - a)];
    
    // Garante a ordem cronológica dos meses
    const mesesOrdenados = Object.keys(frequenciaMensalAnual).sort((a, b) => {
        const [mesA, anoA] = a.split('/');
        const [mesB, anoB] = b.split('/');
        const indexA = monthNames.indexOf(mesA);
        const indexB = monthNames.indexOf(mesB);
        if (anoA !== anoB) return parseInt(anoA) - parseInt(anoB);
        return indexA - indexB;
    });
    const dataTableMensalAnual = [['Mês', 'Casos'], ...mesesOrdenados.map(mes => [mes, frequenciaMensalAnual[mes]])];
    
    // Transforma o `problemasPorMes` para o formato de datatable
    const dataTableProblemasPorMes = {};
    for (const mes in problemasPorMes) {
      dataTableProblemasPorMes[mes] = [['Problema', 'Quantidade'], ...Object.entries(problemasPorMes[mes]).sort(([, a], [, b]) => b - a)];
    }

    const dataTableDiariaPeriodo = [['Dia', 'Casos'], ...Object.entries(frequenciaDiariaPeriodo).sort(([keyA], [keyB]) => {
      const [diaA, mesA] = keyA.split('/');
      const [diaB, mesB] = keyB.split('/');
      if (mesA !== mesB) return parseInt(mesA) - parseInt(mesB);
      return parseInt(diaA) - parseInt(diaB);
    })];
    
    return {
      totalCasos: totalCasosPeriodo,
      problemas: dataTableProblemasPeriodo.length > 1 ? dataTableProblemasPeriodo : [],
      frequenciaMensal: dataTableMensalAnual.length > 1 ? dataTableMensalAnual : [],
      frequenciaDiaria: dataTableDiariaPeriodo.length > 1 ? dataTableDiariaPeriodo : [],
      casos: casosTabelaPeriodo.sort((a, b) => new Date(b.data.split(' ')[0].split('/').reverse().join('-') + ' ' + b.data.split(' ')[1]) - new Date(a.data.split(' ')[0].split('/').reverse().join('-') + ' ' + a.data.split(' ')[1])),
      problemasPorMes: dataTableProblemasPorMes
    };

  } catch (e) {
    Logger.log(`Erro em getIncompatibilidadeDataForTechnician: ${e.stack}`);
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

/**
 * Versão com cache da função de busca de dados de um técnico.
 */
function getIncompatibilidadeDataForTechnicianWithCache(dateRange, tecnicoNome) {
  const cacheKey = `incompat_tecnico_v9_annual_interactive_${dateRange.start}_${dateRange.end}_${tecnicoNome.replace(/\s+/g, '_')}`;
  return getOrSetCache(cacheKey, getIncompatibilidadeDataForTechnician, [dateRange, tecnicoNome]);
}

