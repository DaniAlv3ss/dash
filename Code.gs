/**
 * Script Principal para servir o aplicativo da web e conter funções/constantes globais.
 * @version 2.5 - Adicionada constante para a aba Calendário.
 */

// === CENTRAL DE CONFIGURAÇÕES GLOBAIS ===
const ID_PLANILHA_NPS = "1ewRARy4u4V0MJMoup0XbPlLLUrdPmR4EZwRwmy_ZECM";
const ID_PLANILHA_CALLTECH = "1bmHgGpAXAB4Sh95t7drXLImfNgAojCHv-o2CYS2d3-g";
const ID_PLANILHA_DEVOLUCAO = "1m3tOvmSOJIvRZY9uZNf1idSTEnUFbHIWPNh5tiHkKe0";
const ID_PLANILHA_INCOMPATIBILIDADE = "10l1w3d3HYSKFgSsnjOZ545efR-bdIECEkOR82IjV3TE";
const ID_PLANILHA_AUDITORIA = "11kNgPKKS_Ao7cpyI13DaDRzBaejvW67cKyUz_BrlGGc";


// Nomes das abas
const NOME_ABA_NPS = "Avaliações 2025";
const NOME_ABA_ACOES = "ações 2025";
const NOME_ABA_CALENDARIO = "Calendário"; // <-- Nova Aba
const NOME_ABA_ATENDIMENTO = "Forms";
const NOME_ABA_OS = "NPS Datas";
const NOME_ABA_MANAGER = "Pedidos Manager";
const NOME_ABA_DEVOLUCAO = "Base Devolução";
const NOME_ABA_DEFEITOS = "Dim_Defeitos";
const NOME_ABA_AUDITORIA = "Base de dados";

// Índices de colunas (mantidos aqui para referência global, se necessário)
const INDICES_NPS = {
  // A Coluna A (ID da avaliação) tem índice 0
  DATA_AVALIACAO: 2,  // Coluna C
  CLASSIFICACAO: 5,   // Coluna F
  CLIENTE: 12,        // Coluna M
  EMAIL: 15,          // Coluna P
  PEDIDO_ID: 32,      // Coluna AG
  COMENTARIO: 48,     // Coluna AW
  MOTIVO_FUNCIONAMENTO: 39,
  MOTIVO_QUALIDADE_MONTAGEM: 45,
  MOTIVO_VISUAL_PC: 43,
  MOTIVO_TRANSPORTE: 47
};

const INDICES_ATENDIMENTO = {
  PEDIDO_ID: 2,
  RESOLUCAO: 4,       // Coluna E
  STATUS_ATENDIMENTO: 5, // Coluna F
  VALOR_RETIDO: 6,       // Coluna G
  DEFEITO: 7,         // Coluna H
  RELATO_CLIENTE: 8,  // Coluna I
  DATA_ATENDIMENTO: 0,
  OS: 14
};

const INDICES_OS = {
  PEDIDO_ID: 2,
  OS: 3         // Coluna D
};

const INDICES_CALLTECH = {
  EMAIL: 0,               // Coluna A
  STATUS: 2,              // Coluna C
  CHAMADO_ID: 3,          // Coluna D
  DATA_ABERTURA: 6,       // Coluna G
  PEDIDO_ID: 12,            // Coluna M
  CLIENTE: 14,
  DATA_FINALIZACAO: 16    // Coluna Q
};
// =========================================

// OTIMIZAÇÃO: Duração do cache em segundos (900s = 15 minutos)
const CACHE_EXPIRATION_SECONDS = 900;

/**
 * UTILITY: Obtém dados do cache ou busca-os se não existirem, depois armazena no cache.
 * @param {string} cacheKey A chave única para os dados no cache.
 * @param {function} dataFetchFunction A função que busca os dados frescos.
 * @param {Array<any>} functionArgs Os argumentos para a dataFetchFunction.
 * @returns {Object} Os dados (do cache ou frescos).
 */
function getOrSetCache(cacheKey, dataFetchFunction, functionArgs) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);

  if (cachedData) {
    Logger.log(`CACHE HIT: Chave: ${cacheKey}`);
    return JSON.parse(cachedData);
  }

  Logger.log(`CACHE MISS: Chave: ${cacheKey}. Buscando dados frescos.`);
  const freshData = dataFetchFunction.apply(null, functionArgs);
  // Não armazena no cache se houver erro
  if (freshData && !freshData.error) {
    cache.put(cacheKey, JSON.stringify(freshData), CACHE_EXPIRATION_SECONDS);
  } else {
     Logger.log(`Dados com erro não foram cacheados para a chave: ${cacheKey}`);
  }
  return freshData;
}


/**
 * Função principal que serve o "casco" da aplicação (menu e área de conteúdo).
 * OTIMIZADO: Pré-carrega os dados do dashboard inicial.
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');

  // Otimização: Pré-carrega os dados do dashboard NPS (página inicial)
  try {
    const initialData = getInitialDashboardAndEvolutionDataWithCache(); // Usa a versão com cache
    template.initialData = JSON.stringify(initialData);
  } catch (err) {
    Logger.log("Erro ao pré-carregar dados iniciais: " + err);
    template.initialData = 'null'; // Evita que a página quebre se houver erro no pré-carregamento
  }


  return template.evaluate()
    .setTitle('Dashboard KaBuM! - Monte o Seu PC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Retorna o conteúdo HTML de uma página específica para ser carregado dinamicamente.
 */
function getPageHtml(pageName) {
  // Lista de páginas válidas
  const validPages = ['Dashboard', 'Calltech', 'Devolucao', 'Incompatibilidade', 'Auditoria'];
  if (validPages.includes(pageName)) {
    return HtmlService.createHtmlOutputFromFile('Page_' + pageName).getContent();
  }
  Logger.log('Tentativa de acesso a página inválida: ' + pageName);
  throw new Error('Página não encontrada.');
}


/**
 * Permite a inclusão de arquivos HTML (usados para os scripts JS) dentro de outro template HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * FUNÇÃO UTILITÁRIA GLOBAL: Filtra linhas de dados para manter apenas a mais recente por ID único (Pedido ou ID da avaliação).
 * Utilizada por ambos os dashboards para garantir dados únicos.
 * @param {Array<Array<any>>} dados - O conjunto de dados a ser filtrado.
 * @param {number} idIndex - O índice da coluna que contém o ID do pedido (identificador principal).
 * @param {number} classIndex - O índice da coluna de classificação (para validação).
 * @returns {Array<Array<any>>} Um array com as linhas únicas.
 */
function getUniqueValidRows(dados, idIndex, classIndex) {
  const processedIds = new Map();
  const validClassifications = ['promotor', 'detrator', 'neutro'];
  const fallbackIdIndex = 0; // Coluna A (ID da avaliação)

  // Itera de trás para frente para que a primeira ocorrência (a mais recente) seja a que fica.
  for (let i = dados.length - 1; i >= 0; i--) {
    const linha = dados[i];
    
    // Pula linhas que não têm o número esperado de colunas (evita erros com linhas vazias/malformadas)
    if (!linha || linha.length <= Math.max(idIndex, classIndex, fallbackIdIndex)) {
        continue;
    }

    const classificacaoValue = linha[classIndex];
    // Verifica se a classificação existe e é uma string antes de chamar toLowerCase()
    const classificacao = (typeof classificacaoValue === 'string') ? classificacaoValue.trim().toLowerCase() : '';

    // Ignora linhas sem uma classificação válida
    if (!validClassifications.includes(classificacao)) {
      continue;
    }

    // Tenta usar o ID do pedido primeiro
    let uniqueId = linha[idIndex];
    uniqueId = (typeof uniqueId === 'string' || typeof uniqueId === 'number') ? String(uniqueId).trim() : '';

    // Se não houver ID do pedido válido, usa o ID da coluna A como fallback
    if (!uniqueId) {
      const fallbackIdValue = linha[fallbackIdIndex];
      uniqueId = (typeof fallbackIdValue === 'string' || typeof fallbackIdValue === 'number') ? String(fallbackIdValue).trim() : '';
    }
    
    // Se um ID único foi encontrado (seja do pedido ou fallback) e ainda não foi processado...
    if (uniqueId && !processedIds.has(uniqueId)) {
        processedIds.set(uniqueId, linha);
    }
  }
  return Array.from(processedIds.values());
}

