/**
 * Contém todas as funções do lado do servidor para a página de Incompatibilidades.
 */

/**
 * MODIFICADO: Refina o cálculo do tempo médio de montagem para usar as colunas de data/hora de início e fim,
 * em vez de uma coluna pré-calculada, para maior precisão.
 */
function getIncompatibilidadeData(dateRange) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_INCOMPATIBILIDADE);
    const abaDivergencia = planilha.getSheetByName("Divergência");
    const abaCustomiza = planilha.getSheetByName("Base Customiza");
    if (!abaDivergencia) throw new Error("Aba 'Divergência' não foi encontrada na planilha.");
    if (!abaCustomiza) throw new Error("Aba 'Base Customiza' não foi encontrada na planilha.");

    const INDICES_DIVERGENCIA = {
      CARIMBO_DATA: 0, NOME_TECNICO: 1, PROCESSO: 2, QUAL_PROBLEMA: 3,
      DETALHE_PROBLEMA: 4, OS: 5, INICIO_TRATATIVA: 6, FIM_TRATATIVA: 7,
      PEDIDO: 8, PRODUTO_DIVERGENTE: 9, TIPO_PROBLEMA: 10, AVARIA_CUSTOMIZA: 11,
      TIPO_TROCA: 12, DATA_LIBERACAO: 13, ETIQUETA_PRODUTO_ANTIGO: 14,
      CF_PRODUTO_ANTIGO: 15, CFS_INCOMPATIVEIS: 16, CF_PRODUTO_NOVO: 17,
      ETIQUETA_PRODUTO_NOVO: 18, STATUS_MONTAGEM: 19, STATUS_SISTEMA: 20,
      OBSERVACAO: 21, PEDIDO_COMPLEMENTAR: 22, CUSTO_CF_ANTIGO: 23,
      CUSTO_CF_NOVO: 24, CUSTO_TROCAS: 25, FABRICANTE: 26, NOME_RESPONSAVEL: 27
    };
    
    const INDICES_CUSTOMIZA = {
      NUMERO_OS: 0, DATA_INTEGRACAO: 1, STATUS: 2, DATA_RECEBIMENTO: 3, HORA_RECEBIMENTO: 4,
      RECEBIDO_POR: 5, DATA_INICIO_MONTAGEM: 6, HORA_INICIO_MONTAGEM: 7, TECNICO_MONTAGEM: 8,
      DATA_FIM_MONTAGEM: 9, HORA_FIM_MONTAGEM: 10, MONTAGEM_DURACAO_HR: 11, DATA_INICIO_QUALIDADE: 12,
      HORA_INICIO_QUALIDADE: 13, TECNICO_QUALIDADE: 14, DATA_FIM_QUALIDADE: 15, HORA_FIM_QUALIDADE: 16,
      QUALIDADE_DURACAO_HR: 17, DATA_EMISSAO_NF: 18, HORA_EMISSAO_NF: 19, NUMERO_NF_OS: 20,
      NUMERO_NF_KABUM: 21, TEM_QUARENTENA: 22
    };

    const queryDivergencia = `SELECT * WHERE G >= date '${dateRange.start}' AND G <= date '${dateRange.end}'`;
    const dadosBrutosDivergencia = executeQuery_(abaDivergencia, queryDivergencia);

    const osProcessadas = new Set();
    let custoTotalTrocas = 0, tempoTotalTratativa = 0, tratativasCompletas = 0;
    const contagemTecnicosProblemasOS = {}, contagemProblemas = {}, contagemFabricantes = {};
    const tabelaDetalhada = [], tiposDeProblemaUnicos = new Set(), dailyProblems = {};

    dadosBrutosDivergencia.forEach(linha => {
      const tecnico = (linha[INDICES_DIVERGENCIA.NOME_TECNICO] || "Não especificado").trim();
      const os = linha[INDICES_DIVERGENCIA.OS] ? String(linha[INDICES_DIVERGENCIA.OS]).trim() : null;
      const inicio = linha[INDICES_DIVERGENCIA.INICIO_TRATATIVA];

      if (tecnico && tecnico !== "Não especificado" && os) {
        if (!contagemTecnicosProblemasOS[tecnico]) contagemTecnicosProblemasOS[tecnico] = new Set();
        contagemTecnicosProblemasOS[tecnico].add(os);
      }
      
      if (inicio instanceof Date && os) {
          const dayKey = Utilities.formatDate(inicio, "GMT-3", "yyyy-MM-dd");
          if (!dailyProblems[dayKey]) dailyProblems[dayKey] = new Set();
          dailyProblems[dayKey].add(os);
      }
      
      const problema = (linha[INDICES_DIVERGENCIA.TIPO_PROBLEMA] || "Não especificado").trim();
      if(problema && problema !== "Não especificado") {
        contagemProblemas[problema] = (contagemProblemas[problema] || 0) + 1;
        tiposDeProblemaUnicos.add(problema);
      }
      
      const fabricante = (linha[INDICES_DIVERGENCIA.FABRICANTE] || "Não especificado").trim();
      if(fabricante && fabricante !== "Não especificado") contagemFabricantes[fabricante] = (contagemFabricantes[fabricante] || 0) + 1;
      
      tabelaDetalhada.push({
        data: inicio instanceof Date ? Utilities.formatDate(inicio, "GMT-3", "dd/MM/yyyy HH:mm") : 'N/A',
        tecnico: tecnico, os: os, pedido: linha[INDICES_DIVERGENCIA.PEDIDO],
        produto: linha[INDICES_DIVERGENCIA.PRODUTO_DIVERGENTE], tipoProblema: problema,
        detalheProblema: linha[INDICES_DIVERGENCIA.DETALHE_PROBLEMA]
      });

      if (os && !osProcessadas.has(os)) {
        osProcessadas.add(os);
        let custo = 0;
        const custoRaw = linha[INDICES_DIVERGENCIA.CUSTO_TROCAS];
        if (typeof custoRaw === 'number') custo = custoRaw;
        else if (typeof custoRaw === 'string' && custoRaw.trim() !== '') {
          custo = parseFloat(custoRaw.replace("R$", "").trim().replace(/\./g, "").replace(",", ".")) || 0;
        }
        custoTotalTrocas += custo;

        const fim = linha[INDICES_DIVERGENCIA.FIM_TRATATIVA];
        if (inicio instanceof Date && fim instanceof Date) {
          const diffHoras = (fim.getTime() - inicio.getTime()) / (1000 * 60 * 60);
          if (diffHoras >= 0) {
              tempoTotalTratativa += diffHoras;
              tratativasCompletas++;
          }
        }
      }
    });

    const tempoMedioTratativa = tratativasCompletas > 0 ? (tempoTotalTratativa / tratativasCompletas) : 0;
    
    // Query para buscar todas as colunas necessárias para o cálculo do tempo
    const queryProducao = `SELECT A, G, H, I, J, K WHERE G >= date '${dateRange.start}' AND G <= date '${dateRange.end}'`;
    const dadosProducao = executeQuery_(abaCustomiza, queryProducao);
    
    const producaoPorTecnico = {}, osProducaoUnicas = new Set(), dailyProduction = {};
    const duracaoMontagemPorTecnico = {};
    let totalDuracaoMontagem = 0, maquinasComDuracao = 0;

    dadosProducao.forEach(row => {
        const os = row[0];
        const dataInicioMontagem = row[1];
        const horaInicioMontagem = row[2];
        const tecnicoMontagem = row[3] ? row[3].trim() : null;
        const dataFimMontagem = row[4];
        const horaFimMontagem = row[5];

        if (os) osProducaoUnicas.add(os);
        
        // Calcular duração da montagem em horas
        let duracao = null;
        if (dataInicioMontagem instanceof Date && horaInicioMontagem instanceof Date && dataFimMontagem instanceof Date && horaFimMontagem instanceof Date) {
          const startDateTime = new Date(dataInicioMontagem.getFullYear(), dataInicioMontagem.getMonth(), dataInicioMontagem.getDate(), horaInicioMontagem.getHours(), horaInicioMontagem.getMinutes(), horaInicioMontagem.getSeconds());
          const endDateTime = new Date(dataFimMontagem.getFullYear(), dataFimMontagem.getMonth(), dataFimMontagem.getDate(), horaFimMontagem.getHours(), horaFimMontagem.getMinutes(), horaFimMontagem.getSeconds());
          if (endDateTime > startDateTime) {
              duracao = (endDateTime.getTime() - startDateTime.getTime()) / (1000 * 60 * 60); // Duração em horas
          }
        }
        
        if (tecnicoMontagem) {
            if (!producaoPorTecnico[tecnicoMontagem]) producaoPorTecnico[tecnicoMontagem] = new Set();
            if(os) producaoPorTecnico[tecnicoMontagem].add(os);
            if (duracao !== null && duracao > 0) {
                if (!duracaoMontagemPorTecnico[tecnicoMontagem]) duracaoMontagemPorTecnico[tecnicoMontagem] = [];
                duracaoMontagemPorTecnico[tecnicoMontagem].push(duracao);
            }
        }
        
        if (dataInicioMontagem instanceof Date && os) {
            const dayKey = Utilities.formatDate(dataInicioMontagem, "GMT-3", "yyyy-MM-dd");
            if (!dailyProduction[dayKey]) dailyProduction[dayKey] = new Set();
            dailyProduction[dayKey].add(os);
        }
        
        if (duracao !== null && duracao > 0) {
            totalDuracaoMontagem += duracao;
            maquinasComDuracao++;
        }
    });

    const dailyChartData = [['Dia', 'Máquinas Montadas', 'Problemas']];
    const allDays = new Set([...Object.keys(dailyProduction), ...Object.keys(dailyProblems)]);
    const sortedDays = Array.from(allDays).sort();

    sortedDays.forEach(day => {
        const [,, dayOfMonth] = day.split('-'), month = day.substring(5, 7);
        const label = `${dayOfMonth}/${month}`;
        const productionCount = dailyProduction[day] ? dailyProduction[day].size : 0;
        const problemCount = dailyProblems[day] ? dailyProblems[day].size : 0;
        dailyChartData.push([label, productionCount, problemCount]);
    });

    const producaoContagemFinal = {};
    for (const tecnico in producaoPorTecnico) producaoContagemFinal[tecnico] = producaoPorTecnico[tecnico].size;

    const contagemTecnicosProblemasFinal = {};
    for (const tecnico in contagemTecnicosProblemasOS) contagemTecnicosProblemasFinal[tecnico] = contagemTecnicosProblemasOS[tecnico].size;

    const allTecnicosStats = [];
    const allKnownTecnicos = new Set([...Object.keys(contagemTecnicosProblemasFinal), ...Object.keys(producaoContagemFinal)]);
    allKnownTecnicos.forEach(tecnico => {
        if (tecnico === "Não especificado") return;
        const problemas = contagemTecnicosProblemasFinal[tecnico] || 0;
        const producao = producaoContagemFinal[tecnico] || 0;
        const porcentagem = producao > 0 ? (problemas / producao) * 100 : 0;
        const duracoes = duracaoMontagemPorTecnico[tecnico];
        let tempoMedioMontagem = 0;
        if (duracoes && duracoes.length > 0) {
            tempoMedioMontagem = duracoes.reduce((acc, val) => acc + val, 0) / duracoes.length;
        }
        allTecnicosStats.push({ tecnico, problemas, producao, porcentagem, tempoMedioMontagem });
    });
    
    const totalProducao = osProducaoUnicas.size, totalProblemas = osProcessadas.size;
    const taxaMediaProblemas = totalProducao > 0 ? (totalProblemas / totalProducao) * 100 : 0;
    const tempoMedioMontagemGeral = maquinasComDuracao > 0 ? (totalDuracaoMontagem / maquinasComDuracao) : 0;
    
    const topProblemas = Object.entries(contagemProblemas).sort(([,a],[,b]) => b-a).slice(0, 5);
    const topFabricantes = Object.entries(contagemFabricantes).sort(([,a],[,b]) => b-a).slice(0, 10);
    
    tabelaDetalhada.sort((a, b) => {
        const dateA = a.data !== 'N/A' ? new Date(a.data.split(' ')[0].split('/').reverse().join('-') + ' ' + a.data.split(' ')[1]) : new Date(0);
        const dateB = b.data !== 'N/A' ? new Date(b.data.split(' ')[0].split('/').reverse().join('-') + ' ' + b.data.split(' ')[1]) : new Date(0);
        return dateB - dateA;
    });

    return {
      kpis: { 
        totalIncompatibilidades: osProcessadas.size, custoTotal: custoTotalTrocas, tempoMedioHoras: tempoMedioTratativa,
        totalProducao, totalProblemas, taxaMediaProblemas, totalTiposDeProblema: tiposDeProblemaUnicos.size,
        tempoMedioMontagemGeral: tempoMedioMontagemGeral
      },
      allTecnicosStats, topProblemas, topFabricantes, tabela: tabelaDetalhada,
      uniqueProblemTypes: Array.from(tiposDeProblemaUnicos),
      dailyProductionVsProblems: dailyChartData
    };

  } catch (e) {
    Logger.log(`Erro em getIncompatibilidadeData: ${e.stack}`);
    return { error: e.message };
  }
}

function executeQuery_(sheet, query) {
  const url = `https://docs.google.com/spreadsheets/d/${ID_PLANILHA_INCOMPATIBILIDADE}/gviz/tq?tq=${encodeURIComponent(query)}&gid=${sheet.getSheetId()}`;
  const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } });
  const jsonText = response.getContentText().replace("/*O_o*/\ngoogle.visualization.Query.setResponse(", "").replace(");", "");
  const queryResult = JSON.parse(jsonText);

  if (queryResult.status === 'error') throw new Error(`Erro na query: ${queryResult.errors[0].detailed_message}`);

  return (queryResult.table.rows || []).map(row => 
    (row.c || []).map(cell => {
      if (!cell) return null;
      if (cell.v instanceof Date || (typeof cell.v === 'string' && cell.v.startsWith('Date('))) {
        const params = String(cell.v).match(/\d+/g);
        if (params) return new Date(params[0], params[1], params[2], params[3] || 0, params[4] || 0, params[5] || 0);
      }
      return cell.v;
    })
  );
}

function getIncompatibilidadeDataForTechnician(dateRange, tecnicoNome) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_INCOMPATIBILIDADE);
    const abaDivergencia = planilha.getSheetByName("Divergência");
    const abaCustomiza = planilha.getSheetByName("Base Customiza");

    if (!abaDivergencia) throw new Error("Aba 'Divergência' não foi encontrada.");
    if (!abaCustomiza) throw new Error("Aba 'Base Customiza' não foi encontrada.");
    
    const currentYear = new Date(dateRange.end).getFullYear();
    const yearStartDate = `${currentYear}-01-01`;

    const baseWhereClause = `WHERE B = '${tecnicoNome}' AND G IS NOT NULL AND G >= date '${yearStartDate}' AND G <= date '${dateRange.end}'`;
    const queryCasos = `SELECT G, E, F, I, J, K ${baseWhereClause} ORDER BY G DESC`;
    const dadosBrutosDoAno = executeQuery_(abaDivergencia, queryCasos);

    const osProcessadas = new Set();
    const casosDetalhadosAno = [];
    
    dadosBrutosDoAno.forEach(linha => {
      const os = linha[2] ? String(linha[2]).trim() : null;
      if (!os || osProcessadas.has(os)) return;
      osProcessadas.add(os);

      const data = linha[0];
      if (data instanceof Date) {
        casosDetalhadosAno.push({
          data: data, detalheProblema: linha[1], os: os,
          pedido: linha[3], produto: linha[4], tipoProblema: (linha[5] || "Não especificado").trim()
        });
      }
    });

    const monthNames = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
    const frequenciaMensalAnual = {}, problemasPorMes = {}, contagemProblemasPeriodo = {}, frequenciaDiariaPeriodo = {};
    const casosTabelaPeriodo = [];
    let totalCasosPeriodo = 0;

    const startDatePeriodo = new Date(dateRange.start + 'T00:00:00');
    const endDatePeriodo = new Date(dateRange.end + 'T23:59:59');

    casosDetalhadosAno.forEach(caso => {
      const dataCaso = caso.data;
      const ano = dataCaso.getFullYear(), mesIndex = dataCaso.getMonth();
      const mesLabel = `${monthNames[mesIndex]}/${String(ano).slice(-2)}`;
      frequenciaMensalAnual[mesLabel] = (frequenciaMensalAnual[mesLabel] || 0) + 1;
      if (!problemasPorMes[mesLabel]) problemasPorMes[mesLabel] = {};
      problemasPorMes[mesLabel][caso.tipoProblema] = (problemasPorMes[mesLabel][caso.tipoProblema] || 0) + 1;
      
      if (dataCaso >= startDatePeriodo && dataCaso <= endDatePeriodo) {
        totalCasosPeriodo++;
        contagemProblemasPeriodo[caso.tipoProblema] = (contagemProblemasPeriodo[caso.tipoProblema] || 0) + 1;
        const diaKey = Utilities.formatDate(dataCaso, "GMT-3", "dd/MM");
        frequenciaDiariaPeriodo[diaKey] = (frequenciaDiariaPeriodo[diaKey] || 0) + 1;
        casosTabelaPeriodo.push({
          data: Utilities.formatDate(dataCaso, "GMT-3", "dd/MM/yyyy HH:mm"),
          os: caso.os, pedido: caso.pedido, produto: caso.produto,
          tipoProblema: caso.tipoProblema, detalheProblema: caso.detalheProblema
        });
      }
    });

    let producaoCustomiza = 0;
    const queryProducao = `SELECT A WHERE G >= date '${dateRange.start}' AND G <= date '${dateRange.end}' AND I = '${tecnicoNome}'`;
    
    try {
      const dadosProducaoBrutos = executeQuery_(abaCustomiza, queryProducao);
      producaoCustomiza = new Set(dadosProducaoBrutos.map(row => row[0])).size;
    } catch (e) {
      Logger.log(`Não foi possível buscar dados de produção para ${tecnicoNome}. Erro: ${e.message}`);
    }
    
    const dataTableProblemasPeriodo = [['Problema', 'Quantidade'], ...Object.entries(contagemProblemasPeriodo).sort(([, a], [, b]) => b - a)];
    
    const mesesOrdenados = Object.keys(frequenciaMensalAnual).sort((a, b) => {
        const [mesA, anoA] = a.split('/'), [mesB, anoB] = b.split('/');
        if (anoA !== anoB) return parseInt(anoA) - parseInt(anoB);
        return monthNames.indexOf(mesA) - monthNames.indexOf(mesB);
    });
    const dataTableMensalAnual = [['Mês', 'Casos'], ...mesesOrdenados.map(mes => [mes, frequenciaMensalAnual[mes]])];
    
    const dataTableProblemasPorMes = {};
    for (const mes in problemasPorMes) {
      dataTableProblemasPorMes[mes] = [['Problema', 'Quantidade'], ...Object.entries(problemasPorMes[mes]).sort(([, a], [, b]) => b - a)];
    }

    const dataTableDiariaPeriodo = [['Dia', 'Casos'], ...Object.entries(frequenciaDiariaPeriodo).sort(([keyA], [keyB]) => {
      const [diaA, mesA] = keyA.split('/'), [diaB, mesB] = keyB.split('/');
      if (mesA !== mesB) return parseInt(mesA) - parseInt(mesB);
      return parseInt(diaA) - parseInt(diaB);
    })];
    
    return {
      totalCasos: totalCasosPeriodo, producaoCustomiza,
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

function getIncompatibilidadeDataWithCache(dateRange) {
  const cacheKey = `incompatibilidade_data_v20_real_duration_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getIncompatibilidadeData, [dateRange]);
}

function getIncompatibilidadeDataForTechnicianWithCache(dateRange, tecnicoNome) {
  const cacheKey = `incompat_tecnico_v15_montagem_G_${dateRange.start}_${dateRange.end}_${tecnicoNome.replace(/\s+/g, '_')}`;
  return getOrSetCache(cacheKey, getIncompatibilidadeDataForTechnician, [dateRange, tecnicoNome]);
}
