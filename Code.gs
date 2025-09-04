/**
 * Script Principal para servir o aplicativo da web e conter funções/constantes globais.
 * @version 2.0 - Refatorado
 */

// === CENTRAL DE CONFIGURAÇÕES GLOBAIS ===
const ID_PLANILHA_NPS = "1ewRARy4u4V0MJMoup0XbPlLLUrdPmR4EZwRwmy_ZECM";
const ID_PLANILHA_CALLTECH = "1bmHgGpAXAB4Sh95t7drXLImfNgAojCHv-o2CYS2d3-g";

// Nomes das abas
const NOME_ABA_NPS = "Avaliações 2025";
const NOME_ABA_ACOES = "ações 2025";
const NOME_ABA_ATENDIMENTO = "Forms";
const NOME_ABA_OS = "NPS Datas";
const NOME_ABA_MANAGER = "Pedidos Manager";

// Índices de colunas (mantidos aqui para referência global, se necessário)
const INDICES_NPS = {
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


/**
 * Função principal que serve o "casco" da aplicação (menu e área de conteúdo).
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard KaBuM! - Monte o Seu PC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Retorna o conteúdo HTML de uma página específica para ser carregado dinamicamente.
 */
function getPageHtml(pageName) {
  // --- INÍCIO DA MODIFICAÇÃO ---
  // Adicionamos 'Devolucao' à lista de páginas válidas.
  if (pageName === 'Dashboard' || pageName === 'Calltech' || pageName === 'Devolucao') {
  // --- FIM DA MODIFICAÇÃO ---
    return HtmlService.createHtmlOutputFromFile('Page_' + pageName).getContent();
  }
  throw new Error('Página não encontrada.');
}

/**
 * Permite a inclusão de arquivos HTML (usados para os scripts JS) dentro de outro template HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * FUNÇÃO UTILITÁRIA GLOBAL: Filtra linhas de dados para manter apenas a mais recente por ID de pedido.
 * Utilizada por ambos os dashboards para garantir dados únicos.
 * @param {Array<Array<any>>} dados - O conjunto de dados a ser filtrado.
 * @param {number} idIndex - O índice da coluna que contém o ID do pedido.
 * @param {number} classIndex - O índice da coluna de classificação (para validação).
 * @returns {Array<Array<any>>} Um array com as linhas únicas.
 */
function getUniqueValidRows(dados, idIndex, classIndex) {
  const pedidosProcessados = new Map();
  const validClassifications = ['promotor', 'detrator', 'neutro'];
  
  // Itera de trás para frente para que a primeira ocorrência (a mais recente) seja a que fica.
  for (let i = dados.length - 1; i >= 0; i--) {
    const linha = dados[i];
    const pedidoId = linha[idIndex]?.toString().trim();
    const classificacao = linha[classIndex]?.toString().toLowerCase();

    if (pedidoId && validClassifications.includes(classificacao)) {
      if (!pedidosProcessados.has(pedidoId)) {
        pedidosProcessados.set(pedidoId, linha);
      }
    }
  }
  return Array.from(pedidosProcessados.values());
}
