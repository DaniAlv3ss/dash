/**
 * Contém todas as funções do lado do servidor para a página do Dashboard de NPS.
 */

/**
 * Busca os dados para o carregamento inicial do Dashboard de NPS, unificando chamadas.
 */
function getInitialDashboardAndEvolutionData() {
  const today = new Date();
  const firstDayMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const mainDateRange = {
    start: firstDayMonth.toISOString().split('T')[0],
    end: today.toISOString().split('T')[0]
  };

  const firstDayYear = new Date(today.getFullYear(), 0, 1);
  const evolutionDateRange = {
    start: firstDayYear.toISOString().split('T')[0],
    end: today.toISOString().split('T')[0]
  };

  const mainData = getDashboardData(mainDateRange);
  const evolutionData = getEvolutionChartData(evolutionDateRange, 'month');

  return {
    mainData: mainData,
    evolutionData: evolutionData
  };
}


/**
 * Busca e processa todos os dados necessários para os KPIs e tabelas do dashboard de NPS.
 */
function getDashboardData(dateRange) {
  const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
  const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);
  const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);
  const abaAtendimento = planilhaCalltech.getSheetByName(NOME_ABA_ATENDIMENTO);

  if (!abaNPS || !abaAtendimento) throw new Error("Uma ou mais abas necessárias não foram encontradas.");

  const todosDadosNPS = abaNPS.getRange(2, 1, abaNPS.getLastRow() - 1, abaNPS.getLastColumn()).getDisplayValues();
  const dadosAtendimento = abaAtendimento.getRange(2, 1, abaAtendimento.getLastRow() - 1, abaAtendimento.getLastColumn()).getDisplayValues();
  const osPorPedido = getOsMap();

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
    previousNps: metricasAnteriores.totalRespostas > 0 ? metricasAnteriores.nps : null, 
    npsAcumulado: metricasAcumuladas.nps, totalRespostasAcumulado: metricasAcumuladas.totalRespostas,
    npsMeta: metricasMeta.nps, totalRespostasMeta: metricasMeta.totalRespostas, 
    contagemMotivos: metricasDetalhadas.contagemMotivos, 
    detratoresComChamado: metricasDetalhadas.detratoresComChamado,
    detractorSupportReasons: metricasDetalhadas.detractorSupportReasons, 
    rawResponses: rawResponses
  };
}


/**
 * Calcula métricas básicas de NPS (total, promotores, neutros, detratores) a partir de um conjunto de dados.
 */
function calcularMetricasBasicas(dadosUnicos) {
  const totalRespostas = dadosUnicos.length;
  if (totalRespostas === 0) { 
    return { nps: 0, promotores: 0, neutros: 0, detratores: 0, totalRespostas: 0 };
  }
  let promotores = 0, detratores = 0, neutros = 0;
  for (const linha of dadosUnicos) {
    const classificacao = linha[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
    if (classificacao === 'promotor') promotores++;
    else if (classificacao === 'detrator') detratores++;
    else if (classificacao === 'neutro') neutros++;
  }
  const nps = parseFloat((((promotores - detratores) / totalRespostas) * 100).toFixed(1));
  return { nps, promotores, neutros, detratores, totalRespostas };
}

/**
 * Processa métricas detalhadas, como motivos e detratores com chamado.
 */
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
    
    if (linha[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase() === 'detrator') {
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

/**
 * Busca dados para o gráfico de evolução (mensal ou semanal).
 */
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

/**
 * Calcula a evolução mensal do NPS.
 */
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

        if (!monthlyMetrics[key]) { 
          monthlyMetrics[key] = { label: monthLabel, promoters: 0, neutrals: 0, detractors: 0, firstDayOfMonth: new Date(year, month, 1) }; 
        }
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

/**
 * Calcula a evolução semanal do NPS.
 */
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

        if (!weeklyMetrics[key]) { 
          weeklyMetrics[key] = { label: weekLabel, promoters: 0, neutrals: 0, detractors: 0 }; 
        }
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

/**
 * Busca dados para o gráfico de timeline de performance do NPS vs. Ações.
 */
function getNpsTimelineData(dateRange) {
  try {
    const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
    const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);
    const abaAcoes = planilhaNPS.getSheetByName(NOME_ABA_ACOES);

    if (!abaNPS || !abaAcoes) throw new Error("Aba de NPS ou de Ações não encontrada.");

    const todosDadosNPS = abaNPS.getRange(2, 1, abaNPS.getLastRow() - 1, abaNPS.getLastColumn()).getValues();
    const dadosAcoes = abaAcoes.getRange(2, 1, abaAcoes.getLastRow() - 1, 7).getValues();

    const dadosFiltrados = todosDadosNPS.filter(linha => {
      const dataAvaliacao = linha[INDICES_NPS.DATA_AVALIACAO];
      if (!dataAvaliacao || !(dataAvaliacao instanceof Date)) return false;
      const dataFormatada = Utilities.formatDate(dataAvaliacao, "GMT-3", "yyyy-MM-dd");
      return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
    });
    const dadosUnicos = getUniqueValidRows(dadosFiltrados, INDICES_NPS.PEDIDO_ID, INDICES_NPS.CLASSIFICACAO);

    const weeklyMetrics = {};
    dadosUnicos.forEach(linha => {
      const dataAvaliacao = linha[INDICES_NPS.DATA_AVALIACAO];
      if (!dataAvaliacao || !(dataAvaliacao instanceof Date)) return;

      const utcDate = new Date(Date.UTC(dataAvaliacao.getFullYear(), dataAvaliacao.getMonth(), dataAvaliacao.getDate()));
      const dayOfWeek = utcDate.getUTCDay();
      utcDate.setUTCDate(utcDate.getUTCDate() - dayOfWeek);
      const key = Utilities.formatDate(utcDate, "UTC", "yyyy-MM-dd");

      if (!weeklyMetrics[key]) {
        weeklyMetrics[key] = { promoters: 0, neutrals: 0, detractors: 0 };
      }

      const classificacao = linha[INDICES_NPS.CLASSIFICACAO]?.toString().toLowerCase();
      if (classificacao === 'promotor') weeklyMetrics[key].promoters++;
      else if (classificacao === 'neutro') weeklyMetrics[key].neutrals++;
      else if (classificacao === 'detrator') weeklyMetrics[key].detractors++;
    });

    const sortedKeys = Object.keys(weeklyMetrics).sort();
    const result = [];
    let previousWeekMetrics = null;

    sortedKeys.forEach(key => {
      const metrics = weeklyMetrics[key];
      const total = metrics.promoters + metrics.neutrals + metrics.detractors;
      const nps = total > 0 ? parseFloat((((metrics.promoters - metrics.detractors) / total) * 100).toFixed(1)) : 0;
      
      const weekData = {
        weekKey: key,
        nps: nps,
        detractors: metrics.detractors,
        neutrals: metrics.neutrals,
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

/**
 * Busca as 6 ações estratégicas mais recentes para a timeline.
 */
function getRecentActions() {
  try {
    const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
    const abaAcoes = planilhaNPS.getSheetByName(NOME_ABA_ACOES);
    if (!abaAcoes) throw new Error("Aba de Ações não encontrada.");

    const dadosAcoes = abaAcoes.getRange(2, 2, abaAcoes.getLastRow() - 1, 6).getValues();
    const acoesComData = dadosAcoes
      .map(linha => ({ acao: linha[0], data: linha[5] }))
      .filter(item => item.acao && item.data instanceof Date && !isNaN(item.data));
    
    acoesComData.sort((a, b) => b.data - a.data);
    
    const acoesRecentes = acoesComData.slice(0, 6).map(item => ({
      acao: item.acao,
      data: Utilities.formatDate(item.data, "GMT-3", "dd/MM/yyyy")
    }));
    
    return acoesRecentes.reverse();
  } catch (e) {
    Logger.log(`Erro em getRecentActions: ${e.stack}`);
    return { error: e.message };
  }
}

/**
 * Função de cache para o mapeamento de OS por Pedido, evitando leituras repetidas.
 */
function getOsMap() {
  const cache = CacheService.getScriptCache();
  const cachedOsMap = cache.get('os_map');

  if (cachedOsMap) {
    return new Map(JSON.parse(cachedOsMap));
  }

  const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
  const abaOS = planilhaNPS.getSheetByName(NOME_ABA_OS);
  if (!abaOS) return new Map();

  const dadosOS = abaOS.getRange(2, 1, abaOS.getLastRow() - 1, abaOS.getLastColumn()).getDisplayValues();
  const osPorPedido = new Map();
  dadosOS.forEach(linha => {
    const pedidoId = linha[INDICES_OS.PEDIDO_ID]?.trim();
    const os = linha[INDICES_OS.OS]?.trim();
    if (pedidoId && os) { osPorPedido.set(pedidoId, os); }
  });

  cache.put('os_map', JSON.stringify(Array.from(osPorPedido.entries())), 3600); // Cache por 1 hora
  return osPorPedido;
}

/**
 * Busca detalhes de detratores que abriram chamados no suporte.
 */
function getDetractorSupportDetails(dateRange, reasons) {
  const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
  const planilhaCalltech = SpreadsheetApp.openById(ID_PLANILHA_CALLTECH);
  const abaNPS = planilhaNPS.getSheetByName(NOME_ABA_NPS);
  const abaAtendimento = planilhaCalltech.getSheetByName(NOME_ABA_ATENDIMENTO);

  const todosDadosNPS = abaNPS.getRange(2, 1, abaNPS.getLastRow() - 1, abaNPS.getLastColumn()).getDisplayValues();
  const dadosAtendimento = abaAtendimento.getRange(2, 1, abaAtendimento.getLastRow() - 1, abaAtendimento.getLastColumn()).getDisplayValues();
  const osPorPedido = getOsMap();

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
            dataAvaliacao: `${day}/${month}/${year}`, 
            pedido: pedidoId, 
            os: osPorPedido.get(pedidoId) || '', 
            cliente: detratorRow[INDICES_NPS.CLIENTE],
            classificacao: detratorRow[INDICES_NPS.CLASSIFICACAO], 
            comentario: detratorRow[INDICES_NPS.COMENTARIO], 
            resolucao: resolucao,
            defeito: atendimentoRow[INDICES_ATENDIMENTO.DEFEITO], 
            relatoCliente: atendimentoRow[INDICES_ATENDIMENTO.RELATO_CLIENTE]
          });
        }
      });
    }
  });
  return details;
}

// ===========================================
// === SEÇÃO DE ANÁLISE COM IA (GEMINI) ======
// ===========================================

/**
 * Função central para chamar a API do Gemini.
 */
function callGeminiAPI(prompt, schema) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error("A chave de API do Google AI não foi configurada.");

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
  if (result.candidates && result.candidates[0]?.content.parts[0]?.text) {
    return JSON.parse(result.candidates[0].content.parts[0].text);
  } else {
    Logger.log("Resposta da API inválida: " + responseText);
    throw new Error("Resposta da API inválida.");
  }
}

/**
 * Analisa comentários para identificar problemas e sugerir ações.
 */
function getProblemAndSuggestionAnalysis(comments) {
  try {
    const prompt = `Aja como um analista de dados sênior, especialista em Customer Experience. Analise os seguintes comentários e identifique os 3 problemas mais críticos. Para cada um, sugira uma ação corretiva. Comentários: ${JSON.stringify(comments)} Retorne a resposta estritamente no formato JSON.`;
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

/**
 * Analisa comentários para identificar os tópicos mais comuns.
 */
function getTopicAnalysis(comments) {
  try {
    const prompt = `Aja como um analista de dados sênior. Analise os seguintes comentários e identifique os 4 tópicos mais comentados (ex: "Transporte", "Qualidade de Montagem"). Para cada tópico, inclua a frequência (quantos comentários o mencionam) e o sentimento médio ("Positivo", "Negativo", "Neutro"). Comentários: ${JSON.stringify(comments)} Retorne a resposta estritamente no formato JSON.`;
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

/**
 * Analisa comentários para identificar os pontos fortes e elogios.
 */
function getPraiseAnalysis(comments) {
  try {
    const prompt = `Aja como um analista de marketing. Analise os seguintes comentários e identifique os 3 pontos positivos mais elogiados pelos clientes. Comentários: ${JSON.stringify(comments)} Retorne a resposta estritamente no formato JSON.`;
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

/**
 * Analisa a frequência de palavras-chave nos comentários, categorizando-as.
 */
function getWordFrequencyAnalysis(comments) {
  const stopWords = new Set(['de', 'a', 'o', 'que', 'e', 'do', 'da', 'em', 'um', 'para', 'com', 'não', 'uma', 'os', 'no', 'na', 'por', 'mais', 'as', 'dos', 'como', 'mas', 'foi', 'ao', 'ele', 'das', 'tem', 'à', 'seu', 'sua', 'ou', 'ser', 'quando', 'muito', 'há', 'nos', 'já', 'está', 'eu', 'também', 'só', 'pelo', 'pela', 'até', 'isso', 'ela', 'entre', 'era', 'depois', 'sem', 'mesmo', 'aos', 'ter', 'seus', 'quem', 'nas', 'me', 'esse', 'eles', 'estão', 'você', 'tinha', 'foram', 'essa', 'num', 'nem', 'suas', 'meu', 'às', 'minha', 'numa', 'pelos', 'elas', 'havia', 'seja', 'qual', 'será', 'nós', 'tenho', 'pc', 'computador']);
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
    result[category] = Object.entries(wordCounts[category])
      .sort(([, a], [, b]) => b - a)
      .slice(0, 10)
      .map(([palavra, frequencia]) => ({ palavra, frequencia }));
  }
  return result;
}

// --- VERSÕES COM CACHE PARA SEREM CHAMADAS PELO CLIENTE ---

function getInitialDashboardAndEvolutionDataWithCache() {
  const today = new Date();
  const dateKey = today.toISOString().split('T')[0];
  // A chave de cache para os dados iniciais é baseada apenas na data atual para simplicidade.
  const cacheKey = `initial_dashboard_${dateKey}`;
  return getOrSetCache(cacheKey, getInitialDashboardAndEvolutionData, []);
}

function getDashboardDataWithCache(dateRange) {
  const cacheKey = `dashboard_data_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getDashboardData, [dateRange]);
}

function getEvolutionChartDataWithCache(dateRange, groupBy) {
  const cacheKey = `evolution_data_${dateRange.start}_${dateRange.end}_${groupBy}`;
  return getOrSetCache(cacheKey, getEvolutionChartData, [dateRange, groupBy]);
}

function getNpsTimelineDataWithCache(dateRange) {
  const cacheKey = `timeline_data_${dateRange.start}_${dateRange.end}`;
  return getOrSetCache(cacheKey, getNpsTimelineData, [dateRange]);
}

function getRecentActionsWithCache() {
  const cacheKey = 'recent_actions_v2'; // As ações mudam com menos frequência
  return getOrSetCache(cacheKey, getRecentActions, []);
}

function getDetractorSupportDetailsWithCache(dateRange, reasons) {
   const cacheKey = `detractor_support_${dateRange.start}_${dateRange.end}_${reasons.join('-')}`;
   return getOrSetCache(cacheKey, getDetractorSupportDetails, [dateRange, reasons]);
}

function getWordFrequencyAnalysisWithCache(comments) {
    // Cache para análise de palavras pode ser curto e baseado no hash dos comentários
    const cacheKey = `word_analysis_${Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(comments)).map(b => (b+256).toString(16).slice(-2)).join('')}`;
    return getOrSetCache(cacheKey, getWordFrequencyAnalysis, [comments]);
}


