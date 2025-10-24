/**
 * Contém todas as funções do lado do servidor para o modal de Calendário.
 * @version 1.6 - Reutiliza a lógica de leitura de data da aba Ações (similar a getRecentActions) e atualiza cache.
 */

// Constantes globais (garanta que estas estão definidas em Code.gs ou similar)
// const ID_PLANILHA_NPS = "1ewRARy4u4V0MJMoup0XbPlLLUrdPmR4EZwRwmy_ZECM"; // Exemplo
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

    // --- Processa Aba Ações --- (Lógica Adaptada de getRecentActions)
    if (abaAcoes) {
      const lastRowAcoes = abaAcoes.getLastRow();
      if (lastRowAcoes >= 2) {
        const numRowsAcoes = lastRowAcoes - 1;
        // Lê colunas B (Ação) até G (Data) - Coluna B é índice 0 no array retornado por getValues()
        // Coluna G (Data) será o índice 5 no array retornado
        const range = abaAcoes.getRange(2, 2, numRowsAcoes, 6); // B2:G<lastRow>
        const dadosAcoes = range.getValues();
        Logger.log(`Lidas ${dadosAcoes.length} linhas da aba Ações usando getValues() [B:G].`);

        dadosAcoes.forEach((row, rowIndex) => {
          const rowNumber = rowIndex + 2;
          const titulo = row[0]; // Coluna B (Índice 0)
          const dataAcao = row[5]; // Coluna G (Índice 5) - Deveria ser um objeto Date

          Logger.log(`Ações - Linha ${rowNumber}: Lendo Título='${titulo}', Data='${dataAcao}' (Tipo: ${typeof dataAcao})`);

          // Verifica se a data é um objeto Date válido e se há título
          if (dataAcao instanceof Date && !isNaN(dataAcao) && titulo) {
            // Compara ano e mês (getMonth() é 0-indexado)
            if (dataAcao.getFullYear() === targetYear && dataAcao.getMonth() + 1 === targetMonth) {
              allEvents.push({
                type: 'acao',
                date: Utilities.formatDate(dataAcao, timeZone, "yyyy-MM-dd"),
                title: String(titulo).trim(),
                summary: '', // Aba ações não tem descrição separada nesta leitura
                startTime: null,
                endTime: null,
                links: null
              });
              Logger.log(`Ações - Linha ${rowNumber}: Evento adicionado para ${Utilities.formatDate(dataAcao, timeZone, "yyyy-MM-dd")}.`);
            } else {
                 Logger.log(`Ações - Linha ${rowNumber}: Data ${Utilities.formatDate(dataAcao, timeZone, "yyyy-MM-dd")} fora do período alvo ${targetYear}-${String(targetMonth).padStart(2,'0')}.`);
            }
          } else if (!titulo) {
               Logger.log(`Ações - Linha ${rowNumber}: Ignorado por falta de título.`);
          } else {
               Logger.log(`Ações - Linha ${rowNumber}: Ignorado por data inválida ou ausente (Data: ${dataAcao}).`);
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
        const dadosCalendario = abaCalendario.getRange(2, 1, numRowsCalendario, 6).getValues();
        Logger.log(`Lidas ${dadosCalendario.length} linhas da aba Calendário usando getValues().`);

        dadosCalendario.forEach((row, rowIndex) => {
          const dataEventoRaw = row[0]; // Coluna A
          const titulo = row[1];       // Coluna B
          const resumo = row[2];       // Coluna C
          const horaInicioRaw = row[3]; // Coluna D
          const horaTerminoRaw = row[4]; // Coluna E
          const links = row[5];        // Coluna F
          let dataEvento = null;

           if (dataEventoRaw instanceof Date && !isNaN(dataEventoRaw)) {
              dataEvento = dataEventoRaw;
           } else {
              try {
                dataEvento = parseDateValue_(dataEventoRaw);
              } catch (e) {
                Logger.log(`Calendário - Linha ${rowIndex + 2}: Erro ao parsear data '${dataEventoRaw}'. Erro: ${e.message}`);
                dataEvento = null;
              }
           }

          if (dataEvento instanceof Date && !isNaN(dataEvento) && titulo) {
            if (dataEvento.getFullYear() === targetYear && dataEvento.getMonth() + 1 === targetMonth) {
              const horaInicio = (horaInicioRaw instanceof Date && !isNaN(horaInicioRaw))
                                  ? Utilities.formatDate(horaInicioRaw, timeZone, "HH:mm")
                                  : (typeof horaInicioRaw === 'string' && /^\d{1,2}:\d{2}$/.test(horaInicioRaw.trim())) ? horaInicioRaw.trim() : null;
              const horaTermino = (horaTerminoRaw instanceof Date && !isNaN(horaTerminoRaw))
                                  ? Utilities.formatDate(horaTerminoRaw, timeZone, "HH:mm")
                                  : (typeof horaTerminoRaw === 'string' && /^\d{1,2}:\d{2}$/.test(horaTerminoRaw.trim())) ? horaTerminoRaw.trim() : null;

              allEvents.push({
                type: 'calendario',
                date: Utilities.formatDate(dataEvento, timeZone, "yyyy-MM-dd"),
                title: String(titulo).trim(),
                summary: resumo ? String(resumo).trim() : '',
                startTime: horaInicio,
                endTime: horaTermino,
                links: links ? String(links).trim() : null
              });
              Logger.log(`Calendário - Linha ${rowIndex + 2}: Evento adicionado para ${Utilities.formatDate(dataEvento, timeZone, "yyyy-MM-dd")}.`);
            }
          } else {
             Logger.log(`Calendário - Linha ${rowIndex + 2}: Ignorado por data inválida ou falta de título (Data: ${dataEvento}, Título: ${titulo}).`);
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
        const timeA = a.startTime ? a.startTime : "23:59:59";
        const timeB = b.startTime ? b.startTime : "23:59:59";
        if (timeA !== timeB) {
          const [hA = 0, mA = 0] = (a.startTime || "").split(':').map(Number);
          const [hB = 0, mB = 0] = (b.startTime || "").split(':').map(Number);
          if (hA !== hB) return hA - hB;
          if (mA !== mB) return mA - mB;
        }
        if (a.type !== b.type) {
          return a.type === 'acao' ? -1 : 1;
        }
        return a.title.localeCompare(b.title);
      });
    }

    Logger.log(`Retornando ${Object.keys(eventsByDate).length} dias com eventos. Total de eventos: ${allEvents.length}`);
    return eventsByDate;

  } catch (e) {
    Logger.log(`ERRO GRAVE em getCalendarEvents: ${e.message}\n${e.stack}`);
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
       const date = new Date(Date.UTC(1899, 11, 30 + dateValue));
      if (!isNaN(date)) return date;
    } catch(e){
       Logger.log(`parseDateValue_: Erro ao converter serial ${dateValue} para data: ${e.message}`);
    }
  }

  // Se for string, tenta diferentes formatos
  if (typeof dateValue === 'string' && dateValue.trim() !== '') {
    const trimmedDate = dateValue.trim();
    let parsedDate = null;
    try {
      // Tenta formato ISO ou similar YYYY-MM-DD ou YYYY/MM/DD (com ou sem hora)
      if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(trimmedDate)) {
        const parts = trimmedDate.split(/[-/ T]/);
        if (parts.length >= 3) {
            parsedDate = new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)));
        }
      }
      // Tenta formato brasileiro DD/MM/YYYY ou DD-MM-YYYY (com ou sem hora)
      else if (/^\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(trimmedDate)) {
        const parts = trimmedDate.split(/[-/ T]/);
         if (parts.length >= 3) {
            parsedDate = new Date(Date.UTC(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10)));
         }
      }
      // Última tentativa: construtor Date
      else {
        parsedDate = new Date(trimmedDate);
         if (isNaN(parsedDate)) {
             const timestamp = Date.parse(trimmedDate + ' GMT');
             if (!isNaN(timestamp)) {
                 parsedDate = new Date(timestamp);
             }
         }
      }

      if (parsedDate instanceof Date && !isNaN(parsedDate)) {
        return parsedDate;
      } else {
          Logger.log(`parseDateValue_: Falha ao parsear string de data: ${trimmedDate}. Resultado: ${parsedDate}`);
      }
    } catch (e) {
      Logger.log(`parseDateValue_: Exceção ao parsear string de data '${trimmedDate}': ${e.message}`);
    }
  }
  return null;
}

// --- Versão com Cache ---
function getCalendarEventsWithCache(year, month) {
  // **VERSÃO DO CACHE ATUALIZADA**
  const cacheKey = `calendar_events_v1.6_${year}_${month}`;
  // return getCalendarEvents(year, month); // Descomente para testar SEMPRE buscando dados frescos
  return getOrSetCache(cacheKey, getCalendarEvents, [year, month]);
}

