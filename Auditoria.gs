/**
 * Contém todas as funções do lado do servidor para a página de Auditoria.
 * @version 1.5 - Alterado KPI "Total de NC" para contar auditorias com NC (coluna AA). Removidos KPIs e gráficos de NC por tipo (HW/SW).
 */

function getAuditoriaData(dateRange) {
  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA_AUDITORIA);
    const aba = planilha.getSheetByName(NOME_ABA_AUDITORIA);
    if (!aba) {
      throw new Error("Aba '" + NOME_ABA_AUDITORIA + "' não foi encontrada na planilha.");
    }

    // Verifica se a aba tem mais do que apenas a linha do cabeçalho
    if (aba.getLastRow() < 2) {
      Logger.log("Aba de Auditoria está vazia ou contém apenas o cabeçalho.");
      // Retorna uma estrutura vazia compatível para evitar erros no frontend
      return {
        kpis: { totalAuditorias: 0, totalNaoConformidades: 0, taxaConformidade: 100 }, // Removidos totalNcHardware, totalNcSoftware
        charts: { /* topNcHardware, topNcSoftware removidos */ performanceAuditores: [] },
        tabela: []
      };
    }

    // Leitura dos dados
    const todosDados = aba.getRange(2, 1, aba.getLastRow() - 1, 29).getValues();
    Logger.log(`Total de linhas lidas da aba Auditoria: ${todosDados.length}`);

    // Índices das colunas (base 0)
    const INDICES = {
      // Identificação
      TIMESTAMP: 0, EMAIL: 1, DATA_AUDITORIA: 2, AUDITOR: 3, OS: 4,
      TECNICO_MONTAGEM: 5, TECNICO_QUALIDADE: 6,
      // Hardware NC Checks (7-14, 27-28)
      NC_CABOS: 7, NC_FANS: 8, NC_CONEXOES: 9, NC_FIXACAO: 10, NC_GABINETE: 11,
      NC_FITA_RAM: 12, NC_RGB: 13, NC_CUSTOM_FISICA: 14, NC_ESPUMA: 27, NC_ARGB: 28,
      // Software NC Checks (15-20)
      NC_WINDOWS: 15, NC_XMP: 16, NC_DRIVERS: 17, NC_ARMAZENAMENTO: 18,
      NC_TEMPERATURA: 19, NC_CUSTOM_SOFTWARE: 20,
      // Detalhes NC (21-26)
      NC_DESCRICAO: 21, FOTO_EVIDENCIA: 22, CAUSA_RAIZ: 23, ACAO_CORRETIVA: 24,
      STATUS_FINAL: 25, HOUVE_NC: 26 // Coluna AA
    };

    // Filtragem por data
    const dadosFiltrados = todosDados.filter(linha => {
      const dataAuditoria = linha[INDICES.DATA_AUDITORIA];
      // Verifica se é uma data válida antes de formatar
      if (!dataAuditoria || !(dataAuditoria instanceof Date) || isNaN(dataAuditoria.getTime())) {
          Logger.log(`Data inválida ou ausente na linha: ${linha}`);
          return false;
      }
      try {
        const dataFormatada = Utilities.formatDate(dataAuditoria, Session.getScriptTimeZone(), "yyyy-MM-dd"); // Usar TimeZone do Script
        return dataFormatada >= dateRange.start && dataFormatada <= dateRange.end;
      } catch (e) {
        Logger.log(`Erro ao formatar data '${dataAuditoria}' na linha: ${linha}. Erro: ${e}`);
        return false;
      }
    });
    Logger.log(`Linhas após filtro de data (${dateRange.start} a ${dateRange.end}): ${dadosFiltrados.length}`);

    // Inicialização dos contadores e agregadores
    let totalAuditorias = dadosFiltrados.length;
    let totalAuditoriasComNC = 0; // Contará auditorias com pelo menos 1 NC (baseado na coluna AA)
    // let totalNcItensHardware = 0; // Removido
    // let totalNcItensSoftware = 0; // Removido
    // const contagemNcItensHardware = {}; // Removido
    // const contagemNcItensSoftware = {}; // Removido
    const performanceAuditores = {};
    const performanceTecnicosMontagem = {};
    const performanceTecnicosQualidade = {};
    const tabelaDetalhada = [];

    // Mapas para os itens de verificação (Nome amigável -> Índice) - Mantidos para a tabela
    const hardwareNcsMap = {
      "Organização Cabos": INDICES.NC_CABOS, "Conexão Fans": INDICES.NC_FANS,
      "Conexões Placa/Fonte": INDICES.NC_CONEXOES, "Fixação Componentes": INDICES.NC_FIXACAO,
      "Integridade Gabinete": INDICES.NC_GABINETE, "Fita RAM": INDICES.NC_FITA_RAM,
      "Controle RGB": INDICES.NC_RGB, "Custom Físicas": INDICES.NC_CUSTOM_FISICA,
      "Dano Espuma": INDICES.NC_ESPUMA, "Funcionamento ARGB": INDICES.NC_ARGB
    };
    const softwareNcsMap = {
      "Versão Windows": INDICES.NC_WINDOWS, "Perfil XMP": INDICES.NC_XMP,
      "Drivers Essenciais": INDICES.NC_DRIVERS, "Armazenamento": INDICES.NC_ARMAZENAMENTO,
      "Temperaturas": INDICES.NC_TEMPERATURA, "Custom Software": INDICES.NC_CUSTOM_SOFTWARE
    };

    // Processamento dos dados filtrados
    dadosFiltrados.forEach((linha, index) => {
      const dataAuditoria = linha[INDICES.DATA_AUDITORIA];
      const diaKey = dataAuditoria instanceof Date ? Utilities.formatDate(dataAuditoria, Session.getScriptTimeZone(), "dd/MM") : 'Inválida';

      const auditor = String(linha[INDICES.AUDITOR] || 'N/A').trim();
      const tecMontagem = String(linha[INDICES.TECNICO_MONTAGEM] || 'N/A').trim();
      const tecQualidade = String(linha[INDICES.TECNICO_QUALIDADE] || 'N/A').trim();

      // Inicializa performance se não existir
      if (!performanceAuditores[auditor]) performanceAuditores[auditor] = { auditadas: 0, comNC: 0 };
      performanceAuditores[auditor].auditadas++;
      if (!performanceTecnicosMontagem[tecMontagem]) performanceTecnicosMontagem[tecMontagem] = { auditadas: 0, comNC: 0 };
      performanceTecnicosMontagem[tecMontagem].auditadas++;
      if (!performanceTecnicosQualidade[tecQualidade]) performanceTecnicosQualidade[tecQualidade] = { auditadas: 0, comNC: 0 };
      performanceTecnicosQualidade[tecQualidade].auditadas++;

      // Verifica se houve NC geral na auditoria (baseado na coluna AA)
      const houveNCGeral = String(linha[INDICES.HOUVE_NC] || 'NÃO').trim().toUpperCase() === 'SIM';
      let ncItensNomes = []; // Para a tabela detalhada

      if (houveNCGeral) {
        totalAuditoriasComNC++; // Incrementa o contador de auditorias com NC
        // Atualiza performance
        performanceAuditores[auditor].comNC++;
        performanceTecnicosMontagem[tecMontagem].comNC++;
        performanceTecnicosQualidade[tecQualidade].comNC++;

        // Continua verificando itens individuais para preencher a tabela detalhada
        for (const [nome, indice] of Object.entries(hardwareNcsMap)) {
          const valorCelula = String(linha[indice] || 'Conforme').trim();
          if (valorCelula.toLowerCase() !== 'conforme') {
            ncItensNomes.push(nome);
            // contagemNcItensHardware[nome] = (contagemNcItensHardware[nome] || 0) + 1; // Contagem não é mais necessária para KPIs/Gráficos
            // totalNcItensHardware++; // Contagem não é mais necessária para KPIs/Gráficos
          }
        }
        for (const [nome, indice] of Object.entries(softwareNcsMap)) {
          const valorCelula = String(linha[indice] || 'Conforme').trim();
          if (valorCelula.toLowerCase() !== 'conforme') {
            ncItensNomes.push(nome);
            // contagemNcItensSoftware[nome] = (contagemNcItensSoftware[nome] || 0) + 1; // Contagem não é mais necessária para KPIs/Gráficos
            // totalNcItensSoftware++; // Contagem não é mais necessária para KPIs/Gráficos
          }
        }
      }
      
      // Adiciona à tabela detalhada
      tabelaDetalhada.push({
        dataCompleta: dataAuditoria instanceof Date ? dataAuditoria : new Date(0),
        dataExibicao: diaKey,
        os: linha[INDICES.OS],
        auditor: auditor,
        tecMontagem: tecMontagem,
        tecQualidade: tecQualidade,
        status: houveNCGeral ? "Não Conforme" : "Conforme",
        ncItens: ncItensNomes.join(', '), // Mantém a lista de itens para a tabela
        ncDescricao: linha[INDICES.NC_DESCRICAO] || ''
      });
    });
    Logger.log(`Processamento concluído. Total Auditorias: ${totalAuditorias}, Auditorias Com NC: ${totalAuditoriasComNC}`);

    // Cálculo final dos KPIs
    const taxaConformidade = totalAuditorias > 0 ? ((totalAuditorias - totalAuditoriasComNC) / totalAuditorias) * 100 : 100;

    // Formatação dos dados para gráficos
    // const topNcHardware = Object.entries(contagemNcItensHardware).sort(([, a], [, b]) => b - a).slice(0, 5); // Removido
    // const topNcSoftware = Object.entries(contagemNcItensSoftware).sort(([, a], [, b]) => b - a).slice(0, 5); // Removido

    const performanceAuditoresFormatada = Object.entries(performanceAuditores)
      .map(([nome, dados]) => ({ nome, ...dados, taxaNC: dados.auditadas > 0 ? (dados.comNC / dados.auditadas) * 100 : 0 }))
      .sort((a, b) => b.auditadas - a.auditadas);

    // Retorno dos dados processados
    return {
      kpis: {
        totalAuditorias,
        totalNaoConformidades: totalAuditoriasComNC, // KPI agora usa a contagem de auditorias com NC
        taxaConformidade,
        // totalNcHardware: totalNcItensHardware, // Removido
        // totalNcSoftware: totalNcItensSoftware  // Removido
      },
      charts: {
        // topNcHardware, // Removido
        // topNcSoftware, // Removido
        performanceAuditores: performanceAuditoresFormatada,
      },
      tabela: tabelaDetalhada.sort((a,b) => b.dataCompleta - a.dataCompleta)
                           .map(item => {
                             delete item.dataCompleta;
                             item.data = item.dataExibicao;
                             delete item.dataExibicao;
                             return item;
                           })
    };

  } catch (e) {
    Logger.log(`Erro em getAuditoriaData: ${e.stack}`);
    return { error: e.message };
  }
}

// Função de cache permanece a mesma, apenas incrementa a versão na chave
function getAuditoriaDataWithCache(dateRange) {
  const cacheKey = `auditoria_data_v1.5_${dateRange.start}_${dateRange.end}`; // Versão incrementada
  // return getAuditoriaData(dateRange); // Descomente para desativar cache para testes
  return getOrSetCache(cacheKey, getAuditoriaData, [dateRange]);
}
