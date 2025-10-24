/**
 * Contém todas as funções do lado do servidor para o modal de Calendário.
 * @version 1.3 - Verificação de sintaxe reforçada.
 */

// Constantes globais (garanta que estas estão definidas em Code.gs ou similar)
// const ID_PLANILHA_NPS = "SEU_ID_PLANILHA_NPS";
// const NOME_ABA_ACOES = "ações 2025";
// const NOME_ABA_CALENDARIO = "Calendário";

/**
 * Busca eventos das abas "ações 2025" e "Calendário" para um mês específico.
 * @param {number|string} year O ano (ex: 2025).
 * @param {number|string} month O mês (1-12).
 * @return {object} Um objeto onde as chaves são datas 'YYYY-MM-DD' e os valores são arrays de eventos, ou um objeto { error: string }.
 */
function getCalendarEvents(year, month) {
  try {
    // Validação de entrada
    const targetYear = parseInt(year, 10);
    const targetMonth = parseInt(month, 10); // month é 1-12
    if (isNaN(targetYear) || isNaN(targetMonth) || targetMonth < 1 || targetMonth > 12) {
      Logger.log(`Erro: Ano ou mês inválido recebido - Ano: ${year}, Mês: ${month}`);
      return { error: `Ano (${year}) ou mês (${month}) inválido.` };
    }

    const planilhaNPS = SpreadsheetApp.openById(ID_PLANILHA_NPS);
    const abaAcoes = planilhaNPS.getSheetByName(NOME_ABA_ACOES);
    const abaCalendario = planilhaNPS.getSheetByName(NOME_ABA_CALENDARIO);

    const allEvents = [];
    const timeZone = Session.getScriptTimeZone();
    Logger.log(`Buscando eventos para ${targetYear}-${targetMonth} na timezone ${timeZone}`);

    // --- Processa Aba Ações ---
    if (abaAcoes) {
      const lastRowAcoes = abaAcoes.getLastRow();
      if (lastRowAcoes >= 2) {
        const numRowsAcoes = lastRowAcoes - 1;
        // Colunas B até H (índices 2 a 8), 7 colunas no total
        const range = abaAcoes.getRange(2, 2, numRowsAcoes, 7);
        const dadosAcoes = range.getValues();
        Logger.log(`Lidas ${dadosAcoes.length} linhas da aba Ações.`);

        dadosAcoes.forEach((row, rowIndex) => {
          // Índices baseados no array 'row' (começa em 0)
          const titulo = row[1];       // Coluna C (índice 1 no array 'row')
          const descricao = row[2];    // Coluna D (índice 2 no array 'row')
          const dataAcaoRaw = row[6]; // Coluna H (índice 6 no array 'row')
          let dataAcao = null;

          try {
            dataAcao = parseDateValue_(dataAcaoRaw);
          } catch (e) {
            Logger.log(`Erro ao parsear data na linha ${rowIndex + 2} da aba Ações: ${dataAcaoRaw}. Erro: ${e.message}`);
            dataAcao = null;
          }

          if (dataAcao instanceof Date && !isNaN(dataAcao) && titulo) {
            if (dataAcao.getFullYear() === targetYear && dataAcao.getMonth() + 1 === targetMonth) {
              allEvents.push({
                type: 'acao',
                date: Utilities.formatDate(dataAcao, timeZone, "yyyy-MM-dd"),
                title: String(titulo).trim(),
                summary: descricao ? String(descricao).trim() : '',
                startTime: null,
                endTime: null,
                links: null
              });
            }
          }
        });
      } else {
        Logger.log(`Aba "${NOME_ABA_ACOES}" tem menos de 2 linhas.`);
      }
    } else {
      Logger.log(`Aba "${NOME_ABA_ACOES}" não encontrada.`);
    }

    // --- Processa Aba Calendário ---
    if (abaCalendario) {
      const lastRowCalendario = abaCalendario.getLastRow();
      if (lastRowCalendario >= 2) {
        const numRowsCalendario = lastRowCalendario - 1;
        // Colunas A até F (índices 1 a 6), 6 colunas no total
        const dadosCalendario = abaCalendario.getRange(2, 1, numRowsCalendario, 6).getValues();
        Logger.log(`Lidas ${dadosCalendario.length} linhas da aba Calendário.`);

        dadosCalendario.forEach((row, rowIndex) => {
          const dataEventoRaw = row[0]; // Coluna A
          const titulo = row[1];       // Coluna B
          const resumo = row[2];       // Coluna C
          const horaInicioRaw = row[3]; // Coluna D
          const horaTerminoRaw = row[4]; // Coluna E
          const links = row[5];        // Coluna F
          let dataEvento = null;

          try {
            dataEvento = parseDateValue_(dataEventoRaw);
          } catch (e) {
            Logger.log(`Erro ao parsear data na linha ${rowIndex + 2} da aba Calendário: ${dataEventoRaw}. Erro: ${e.message}`);
            dataEvento = null;
          }

          if (dataEvento instanceof Date && !isNaN(dataEvento) && titulo) {
            if (dataEvento.getFullYear() === targetYear && dataEvento.getMonth() + 1 === targetMonth) {
              const horaInicio = (horaInicioRaw instanceof Date && !isNaN(horaInicioRaw)) ? Utilities.formatDate(horaInicioRaw, timeZone, "HH:mm") : null;
              const horaTermino = (horaTerminoRaw instanceof Date && !isNaN(horaTerminoRaw)) ? Utilities.formatDate(horaTerminoRaw, timeZone, "HH:mm") : null;

              allEvents.push({
                type: 'calendario',
                date: Utilities.formatDate(dataEvento, timeZone, "yyyy-MM-dd"),
                title: String(titulo).trim(),
                summary: resumo ? String(resumo).trim() : '',
                startTime: horaInicio,
                endTime: horaTermino,
                links: links ? String(links).trim() : null
              });
            }
          }
        });
      } else {
        Logger.log(`Aba "${NOME_ABA_CALENDARIO}" tem menos de 2 linhas.`);
      }
    } else {
      Logger.log(`Aba "${NOME_ABA_CALENDARIO}" não encontrada.`);
    }

    // --- Agrupa e Ordena Eventos ---
    const eventsByDate = {};
    allEvents.forEach(event => {
      if (!eventsByDate[event.date]) {
        eventsByDate[event.date] = [];
      }
      eventsByDate[event.date].push(event);
    });

    // Ordena os eventos dentro de cada dia
    for (const date in eventsByDate) {
      eventsByDate[date].sort((a, b) => {
        const timeA = a.startTime ? a.startTime : "23:59:59"; // Eventos sem hora vão para o fim
        const timeB = b.startTime ? b.startTime : "23:59:59";
        if (timeA !== timeB) {
          return timeA.localeCompare(timeB);
        }
        if (a.type !== b.type) {
          return a.type === 'acao' ? -1 : 1; // Prioriza Ação
        }
        return a.title.localeCompare(b.title); // Ordena por título
      });
    }

    Logger.log(`Retornando ${Object.keys(eventsByDate).length} dias com eventos.`);
    return eventsByDate;

  } catch (e) {
    Logger.log(`ERRO GRAVE em getCalendarEvents: ${e.message}\n${e.stack}`);
    // Retorna um objeto de erro claro para o frontend
    return { error: `Erro no servidor ao buscar eventos: ${e.message}` };
  }
}

/**
 * Função auxiliar para tentar parsear diferentes formatos de data.
 * @param {any} dateValue O valor da célula.
 * @return {Date|null} Um objeto Date ou null se inválido.
 * @private
 */
function parseDateValue_(dateValue) {
  // Se já for uma data válida, retorna
  if (dateValue instanceof Date && !isNaN(dateValue)) {
    return dateValue;
  }

  // Se for número (provavelmente serial do Sheets)
  if (typeof dateValue === 'number') {
    try {
      // Ajuste para UTC para evitar problemas de fuso na conversão do serial
       const date = new Date(Date.UTC(1899, 11, 30 + dateValue));
      if (!isNaN(date)) return date;
    } catch(e){
       Logger.log(`Erro ao converter serial ${dateValue} para data: ${e.message}`);
    }
  }

  // Se for string, tenta diferentes formatos
  if (typeof dateValue === 'string' && dateValue.trim() !== '') {
    const trimmedDate = dateValue.trim();
    let parsedDate = null;
    try {
      // Tenta formato ISO ou similar YYYY-MM-DD ou YYYY/MM/DD
      if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(trimmedDate)) {
        const parts = trimmedDate.split(/[-/]/);
        // Usa UTC para consistência
        parsedDate = new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)));
      }
      // Tenta formato brasileiro DD/MM/YYYY ou DD-MM-YYYY
      else if (/^\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(trimmedDate)) {
        const parts = trimmedDate.split(/[-/]/);
         // Usa UTC para consistência
        parsedDate = new Date(Date.UTC(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10)));
      }
      // Última tentativa: deixar o construtor Date tentar (menos confiável)
      else {
        parsedDate = new Date(trimmedDate);
         // Se resultou em inválido, tenta forçar UTC/GMT
         if (isNaN(parsedDate)) {
             const timestamp = Date.parse(trimmedDate + ' GMT');
             if (!isNaN(timestamp)) {
                 parsedDate = new Date(timestamp);
             }
         }
      }

      // Verifica validade final
      if (parsedDate instanceof Date && !isNaN(parsedDate)) {
        return parsedDate;
      } else {
          Logger.log(`Falha ao parsear string de data: ${trimmedDate}`);
      }
    } catch (e) {
      Logger.log(`Exceção ao parsear string de data '${trimmedDate}': ${e.message}`);
    }
  }

  // Se chegou aqui, não conseguiu parsear
  return null;
}

// --- Versão com Cache ---
function getCalendarEventsWithCache(year, month) {
  const cacheKey = `calendar_events_v1.3_${year}_${month}`; // Versão incrementada
  // Descomente a linha abaixo para testar SEMPRE buscando dados frescos
  // return getCalendarEvents(year, month);
  return getOrSetCache(cacheKey, getCalendarEvents, [year, month]);
}

