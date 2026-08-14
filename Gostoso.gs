/**
 * ============================================================================
 * CONFIGURAÇÃO GERAL DO SISTEMA
 * 
 * SEGURANÇA: O token da Autentique DEVE ser salvo em:
 * Extensões > Apps Script > Configurações do Projeto > Propriedades do Script
 * Nome da Propriedade: AUTENTIQUE_TOKEN
 * ============================================================================
 */
const CONFIG = {
  get AUTENTIQUE_TOKEN() {
    const token = PropertiesService.getScriptProperties().getProperty('AUTENTIQUE_TOKEN');
    if (!token) {
      throw new Error('Token da Autentique não configurado em Propriedades do Script (AUTENTIQUE_TOKEN).');
    }
    return token.trim();
  },
  AUTENTIQUE_URL: 'https://api.autentique.com.br/v2/graphql',

  // ID DA PLANILHA DO GOOGLE SHEETS
  SPREADSHEET_ID: '1NuS3UHrRZGVrY1DKGYlK5sSgIL0FqjM6nZ1Tzbm_XtY',

  // IDs DAS PASTAS NO GOOGLE DRIVE
  FOLDER_ID_ORIGEM: '1FwXxNTC1nqYdhyl0mihTSM24DJl1tDLK',     // Pasta do PDF consolidado
  FOLDER_ID_DESTINO: '1jVTPUF44A1kXgTQdQXt6UDDE7f5juAt8',   // Pasta dos PDFs divididos

  // NOMES DAS ABAS NA PLANILHA
  ABA_DADOS_BASE: 'Dados Base',
  ABA_LOG: 'Log de Execução',

  // E-MAILS INSTITUCIONAIS A SEREM IGNORADOS DA CONTAGEM DE SIGNATÁRIOS
  EMAILS_IGNORADOS: ['contratos@institutoformar.org']
};

/**
 * 1. PONTO DE ENTRADA DA APLICAÇÃO WEB (Web App)
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Controle e Envio de Ponto')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. RECEBE O ARQUIVO DA INTERFACE WEB, SALVA NO DRIVE E DISPARA O PROCESSAMENTO
 */
function salvarPdfNoDriveEProcessar(arquivoData) {
  try {
    if (!arquivoData || !arquivoData.bytes) {
      throw new Error('Nenhum conteúdo de arquivo foi enviado.');
    }

    if (arquivoData.mimeType !== MimeType.PDF) {
      throw new Error('Apenas arquivos no formato PDF são aceitos.');
    }

    if (arquivoData.bytes.length > 35000000) {
      throw new Error('O arquivo excede o limite máximo permitido de envio.');
    }

    const pastaOrigem = DriveApp.getFolderById(CONFIG.FOLDER_ID_ORIGEM);

    // Limpa PDFs antigos na pasta de origem
    const arquivosAntigos = pastaOrigem.getFilesByType(MimeType.PDF);
    while (arquivosAntigos.hasNext()) {
      arquivosAntigos.next().setTrashed(true);
    }

    const bytes = Utilities.base64Decode(arquivoData.bytes);
    const nomeSanitizado = String(arquivoData.nome || 'Ponto_Consolidado.pdf').replace(/[\/\\]/g, '_');
    const blob = Utilities.newBlob(bytes, MimeType.PDF, nomeSanitizado);

    const novoArquivo = pastaOrigem.createFile(blob);
    Logger.log(`✅ Arquivo salvo na Origem: ${novoArquivo.getName()} (ID: ${novoArquivo.getId()})`);

    // Dispara a automação principal
    processarControleDePontoComLog();

    return {
      sucesso: true,
      mensagem: 'Arquivo processado com sucesso e espelhos encaminhados para a Autentique!'
    };

  } catch (e) {
    Logger.log(`❌ Erro em salvarPdfNoDriveEProcessar: ${e.toString()}`);
    throw new Error(`Falha no upload/processamento: ${e.message}`);
  }
}

/**
 * 3. RETORNA O HISTÓRICO DE LOGS PARA A TABELA HTML (Mapeia todos os formatos de URL)
 */
function obterHistoricoLogs(pasta) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abaLog = spreadsheet.getSheetByName(CONFIG.ABA_LOG);

    if (!abaLog) return [];

    const ultimaLinha = abaLog.getLastRow();
    if (ultimaLinha <= 1) return [];

    const pastaFiltro = pasta ? String(pasta).trim() : '';
    const inicio = pastaFiltro ? 2 : Math.max(2, ultimaLinha - 50);
    const quantidade = ultimaLinha - inicio + 1;

    const valores = abaLog.getRange(inicio, 1, quantidade, 11).getValues();

    const comLinha = valores.map((row, i) => ({ row: row, numeroLinha: inicio + i }));

    const filtrados = comLinha.filter(({ row }) => {
      const pastaAtual = String(row[10] || '').trim();
      return pastaFiltro ? pastaAtual === pastaFiltro : pastaAtual === '';
    });

    return filtrados.reverse().map(({ row, numeroLinha }) => {
      const docId = String(row[6] || '').trim();
      const linkPlanilha = String(row[7] || '').trim();

      const item = {
        linha: numeroLinha,
        data: row[0] instanceof Date
          ? Utilities.formatDate(row[0], "GMT-3", "dd/MM/yyyy HH:mm:ss")
          : String(row[0]),
        nome: String(row[2] || 'N/A'),
        link: linkPlanilha,
        urlAutentique: linkPlanilha,
        urlDocumento: linkPlanilha,
        urlDrive: linkPlanilha,
        status: String(row[8] || 'Desconhecido'),
        pasta: String(row[10] || '').trim(),
        signatarios: [],
        arquivoAssinadoUrl: ''
      };

      if (docId && /^[0-9a-f-]{20,}$/i.test(docId)) {
        try {
          const detalhes = obterDetalhesDocumento(docId);
          item.signatarios = detalhes.signatarios;
          item.arquivoAssinadoUrl = detalhes.arquivoAssinadoUrl;

          if (detalhes.arquivoAssinadoUrl) {
            item.urlAutentique = detalhes.arquivoAssinadoUrl;
            item.urlDocumento = detalhes.arquivoAssinadoUrl;
            item.link = detalhes.arquivoAssinadoUrl;
          }
        } catch (erroDetalhes) {
          Logger.log(`⚠️ Não foi possível obter detalhes do doc ${docId}: ${erroDetalhes.toString()}`);
        }
      }

      return item;
    });

  } catch (e) {
    Logger.log(`❌ Erro ao ler histórico: ${e.toString()}`);
    return [];
  }
}

/**
 * 3.1 BUSCA OS DETALHES DE UM DOCUMENTO NA AUTENTIQUE VIA GRAPHQL
 */
function obterDetalhesDocumento(docId) {
  const query = `
    query GetDocument($id: UUID!) {
      document(id: $id) {
        id
        name
        files {
          original
          signed
        }
        signatures {
          public_id
          name
          email
          link {
            short_link
          }
          signed {
            created_at
          }
        }
      }
    }
  `;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': `Bearer ${CONFIG.AUTENTIQUE_TOKEN}` },
    payload: JSON.stringify({ query: query, variables: { id: docId } }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(CONFIG.AUTENTIQUE_URL, options);
  if (response.getResponseCode() !== 200) throw new Error(`Autentique HTTP ${response.getResponseCode()}`);

  const resultado = JSON.parse(response.getContentText());
  if (resultado.errors) throw new Error(`Erro GraphQL: ${JSON.stringify(resultado.errors)}`);

  const documento = resultado.data?.document || null;
  const assinaturas = documento?.signatures || [];

  const signatariosValidos = assinaturas.filter(s => {
    const emailSignatario = String(s.email || '').toLowerCase().trim();
    return !CONFIG.EMAILS_IGNORADOS.some(ignorar => ignorar.toLowerCase() === emailSignatario);
  });

  const signatarios = signatariosValidos.map(s => ({
    nome: s.name || s.email || 'Signatário',
    assinado: !!(s.signed && s.signed.created_at),
    linkAssinatura: s.link?.short_link || ''
  }));

  const arquivoAssinadoUrl = documento?.files?.signed || '';

  return { signatarios, arquivoAssinadoUrl };
}

/**
 * 3.2 LISTA OS NOMES DE PASTAS DE ARQUIVO JÁ USADAS
 */
function obterPastasArquivadas() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abaLog = spreadsheet.getSheetByName(CONFIG.ABA_LOG);

    if (!abaLog) return [];

    const ultimaLinha = abaLog.getLastRow();
    if (ultimaLinha <= 1) return [];

    const valoresPasta = abaLog.getRange(2, 11, ultimaLinha - 1, 1).getValues();
    const pastas = new Set();

    valoresPasta.forEach(r => {
      const v = String(r[0] || '').trim();
      if (v) pastas.add(v);
    });

    return Array.from(pastas).sort((a, b) => a.localeCompare(b, 'pt-BR'));

  } catch (e) {
    Logger.log(`❌ Erro ao listar pastas: ${e.toString()}`);
    return [];
  }
}

/**
 * 3.3 ARQUIVA DOCUMENTOS SELECIONADOS
 */
function arquivarDocumentos(linhas, nomePasta) {
  if (!Array.isArray(linhas) || linhas.length === 0) throw new Error('Nenhum documento selecionado.');
  const pasta = String(nomePasta || '').trim();
  if (!pasta) throw new Error('Informe um nome para a pasta de arquivo.');

  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const abaLog = spreadsheet.getSheetByName(CONFIG.ABA_LOG);

  linhas.forEach(numeroLinha => {
    const linha = parseInt(numeroLinha, 10);
    if (!isNaN(linha) && linha >= 2) abaLog.getRange(linha, 11).setValue(pasta);
  });

  return { sucesso: true, mensagem: `${linhas.length} documento(s) arquivado(s) em "${pasta}".` };
}

/**
 * 3.4 RESTAURA DOCUMENTOS ARQUIVADOS
 */
function restaurarDocumentos(linhas) {
  if (!Array.isArray(linhas) || linhas.length === 0) throw new Error('Nenhum documento selecionado.');

  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const abaLog = spreadsheet.getSheetByName(CONFIG.ABA_LOG);

  linhas.forEach(numeroLinha => {
    const linha = parseInt(numeroLinha, 10);
    if (!isNaN(linha) && linha >= 2) abaLog.getRange(linha, 11).setValue('');
  });

  return { sucesso: true, mensagem: `${linhas.length} documento(s) restaurado(s) para Ativos.` };
}

/**
 * 3.5 EXCLUI DOCUMENTOS SELECIONADOS PERMANENTEMENTE
 */
function excluirDocumentos(linhas) {
  if (!Array.isArray(linhas) || linhas.length === 0) throw new Error('Nenhum documento selecionado.');

  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const abaLog = spreadsheet.getSheetByName(CONFIG.ABA_LOG);

  const linhasOrdenadas = linhas
    .map(n => parseInt(n, 10))
    .filter(n => !isNaN(n) && n >= 2)
    .sort((a, b) => b - a);

  linhasOrdenadas.forEach(numeroLinha => abaLog.deleteRow(numeroLinha));

  return { sucesso: true, mensagem: `${linhasOrdenadas.length} documento(s) excluído(s) permanentemente.` };
}

/**
 * 4. CARREGA A BIBLIOTECA PDF-LIB VIA CDN
 */
function carregarPdfLib() {
  if (typeof setTimeout === 'undefined') {
    this.setTimeout = function (callback, delay) {
      if (delay && delay > 0) {
        Utilities.sleep(delay);
      }
      try {
        callback();
      } catch (e) {
        Logger.log(`⚠️ Erro dentro do Callback do setTimeout: ${e.toString()}`);
      }
      return 0;
    };
  }

  if (typeof PDFLib === 'undefined') {
    const response = UrlFetchApp.fetch('https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js');
    if (response.getResponseCode() !== 200) {
      throw new Error('Não foi possível carregar a biblioteca PDF-Lib via CDN.');
    }
    const codePdfLib = response.getContentText();
    (0, eval)(codePdfLib);
  }
}

/**
 * 5. FUNÇÃO PRINCIPAL DE PROCESSAMENTO
 */
async function processarControleDePontoComLog() {
  Logger.log('🚀 Iniciando o processamento do controle de ponto...');
  carregarPdfLib();

  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (err) {
    Logger.log('❌ ERRO: Não foi possível abrir a planilha. Verifique o SPREADSHEET_ID.');
    return;
  }

  const abaDados = spreadsheet.getSheetByName(CONFIG.ABA_DADOS_BASE);
  const abaLog = spreadsheet.getSheetByName(CONFIG.ABA_LOG);

  if (!abaDados || !abaLog) {
    Logger.log('❌ ERRO: Verifique se as abas "Dados Base" e "Log de Execução" existem.');
    return;
  }

  const ultimaLinha = abaDados.getLastRow();
  if (ultimaLinha <= 1) {
    Logger.log('⚠️ AVISO: Nenhum funcionário encontrado na aba "Dados Base".');
    return;
  }

  const totalLinhasDados = ultimaLinha - 1;
  const dados = abaDados.getRange(2, 1, totalLinhasDados, 4).getValues();

  const funcionarios = dados
    .filter(row => row[0] !== '' && String(row[3]).trim().toLowerCase() === 'ativo')
    .map((row, index) => ({
      id: index + 1,
      nome: String(row[0]).trim(),
      whatsapp: String(row[1]).replace(/\D/g, '').trim(),
      email: String(row[2]).trim()
    }));

  if (funcionarios.length === 0) {
    Logger.log('⚠️ Nenhum colaborador com status "Ativo" foi encontrado.');
    return;
  }

  Logger.log(`👥 Total de colaboradores ativos encontrados: ${funcionarios.length}`);

  const pastaOrigem = DriveApp.getFolderById(CONFIG.FOLDER_ID_ORIGEM);
  const arquivos = pastaOrigem.getFilesByType(MimeType.PDF);

  if (!arquivos.hasNext()) {
    Logger.log('❌ ERRO: Nenhum arquivo PDF encontrado na pasta de origem.');
    registrarLog(abaLog, new Date(), '-', 'N/A', '-', '-', '-', '-', '-', 'Erro', 'Nenhum PDF encontrado na pasta de origem.');
    return;
  }

  const arquivoPdfGeral = arquivos.next();
  Logger.log(`📄 Arquivo PDF carregado: ${arquivoPdfGeral.getName()}`);

  const pdfBytesGeral = new Uint8Array(arquivoPdfGeral.getBlob().getBytes());

  Logger.log('🔍 Mapeando e lendo textos do PDF...');
  const textosPorPagina = extrairTextoNativoPDF(pdfBytesGeral);
  Logger.log(`📖 Total de páginas mapeadas no PDF: ${textosPorPagina.length}`);

  const pdfDocGeral = await PDFLib.PDFDocument.load(pdfBytesGeral);

  for (const func of funcionarios) {
    const dataHoraExecucao = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
    Logger.log(`\n--------------------------------------------------`);
    Logger.log(`🔄 Procurando página para: ${func.nome}`);

    if (!func.whatsapp || func.whatsapp.length < 10) {
      const msgErro = `Número de WhatsApp inválido (${func.whatsapp}).`;
      Logger.log(`⚠️ ERRO: ${msgErro}`);
      registrarLog(abaLog, dataHoraExecucao, func.id, func.nome, func.whatsapp, '-', '-', '-', '-', 'Erro WhatsApp', msgErro);
      continue;
    }

    const paginaEncontrada = buscarPaginaPorNome(textosPorPagina, func.nome);

    if (paginaEncontrada === -1) {
      const msgErro = `Nome "${func.nome}" não localizado em nenhuma página do PDF.`;
      Logger.log(`⚠️ ERRO: ${msgErro}`);
      registrarLog(abaLog, dataHoraExecucao, func.id, func.nome, func.whatsapp, 'Não Encontrada', '-', '-', '-', 'Erro Busca Nome', msgErro);
      continue;
    }

    Logger.log(`🎯 Nome localizado com sucesso na Página: ${paginaEncontrada}`);

    try {
      const pdfUnico = await PDFLib.PDFDocument.create();
      const [paginaCopiada] = await pdfUnico.copyPages(pdfDocGeral, [paginaEncontrada - 1]);
      pdfUnico.addPage(paginaCopiada);

      const pdfBytesSeparado = await pdfUnico.save();

      // Tenta extrair a competência (mês/ano) diretamente do texto da página do documento.
      // Só usa a data atual como último recurso, caso não seja encontrada nenhuma data no PDF.
      const textoDaPagina = textosPorPagina[paginaEncontrada - 1] || '';
      const competenciaDetectada = extrairCompetenciaDoTexto(textoDaPagina);
      const mesAnoCompetencia = competenciaDetectada || Utilities.formatDate(new Date(), "GMT-3", "MM/yyyy");

      if (!competenciaDetectada) {
        Logger.log(`⚠️ Competência não localizada no PDF de ${func.nome}. Usando o mês/ano atual (${mesAnoCompetencia}) como referência.`);
      } else {
        Logger.log(`📅 Competência detectada no documento de ${func.nome}: ${mesAnoCompetencia}`);
      }

      const [mesArquivo, anoArquivo] = mesAnoCompetencia.split('/');
      const mesAnoStr = `${anoArquivo}-${mesArquivo}`;
      const nomeArquivo = `Ponto_${func.nome.replace(/[^a-zA-Z0-9]/g, '_')}_${mesAnoStr}.pdf`;

      const pastaDestino = DriveApp.getFolderById(CONFIG.FOLDER_ID_DESTINO);
      const pdfBlob = Utilities.newBlob(Array.from(pdfBytesSeparado), MimeType.PDF, nomeArquivo);
      const arquivoCriado = pastaDestino.createFile(pdfBlob);

      Logger.log(`💾 PDF individual criado: ${arquivoCriado.getName()}`);

      Logger.log(`📡 Enviando para Autentique via WhatsApp (${func.whatsapp})...`);
      const resposta = enviarParaAutentiqueViaWhatsapp(pdfBlob, func, mesAnoCompetencia);

      if (resposta && resposta.data && resposta.data.createDocument) {
        const docId = resposta.data.createDocument.id;
        const linkGestao = `https://painel.autentique.com.br/documentos/${docId}`;

        Logger.log(`✅ Sucesso! Enviado com sucesso. ID: ${docId}`);

        registrarLog(
          abaLog,
          dataHoraExecucao,
          func.id,
          func.nome,
          func.whatsapp,
          paginaEncontrada,
          nomeArquivo,
          docId,
          linkGestao,
          'Sucesso',
          'Documento enviado para assinatura via WhatsApp na Autentique.'
        );

      } else {
        const detalheErro = JSON.stringify(resposta);
        Logger.log(`❌ Erro Autentique: ${detalheErro}`);
        registrarLog(abaLog, dataHoraExecucao, func.id, func.nome, func.whatsapp, paginaEncontrada, nomeArquivo, '-', '-', 'Erro Autentique', detalheErro);
      }

    } catch (err) {
      Logger.log(`❌ Erro Geral de Processamento: ${err.toString()}`);
      registrarLog(abaLog, dataHoraExecucao, func.id, func.nome, func.whatsapp, paginaEncontrada, '-', '-', '-', 'Erro Execução', err.toString());
    }
  }

  Logger.log(`\n==================================================`);
  Logger.log('🏁 Processamento concluído com sucesso!');
}

/**
 * 6.1 AUXILIAR: EXTRAI A COMPETÊNCIA (MÊS/ANO) DO TEXTO DA PÁGINA DO DOCUMENTO
 * Tenta, em ordem: "MM/AAAA" isolado, uma data completa "DD/MM/AAAA", ou o nome do mês por extenso + ano.
 * Retorna no formato "MM/AAAA" ou null se nada for encontrado.
 */
function extrairCompetenciaDoTexto(texto) {
  if (!texto) return null;

  // Formato "MM/AAAA" (o mais comum em espelhos de ponto, ex: "05/2026")
  const matchMesAno = texto.match(/\b(0[1-9]|1[0-2])\s*\/\s*(20\d{2})\b/);
  if (matchMesAno) {
    return `${matchMesAno[1]}/${matchMesAno[2]}`;
  }

  // Data completa "DD/MM/AAAA" — extrai apenas o mês/ano
  const matchDataCompleta = texto.match(/\b(0[1-9]|[12]\d|3[01])\s*\/\s*(0[1-9]|1[0-2])\s*\/\s*(20\d{2})\b/);
  if (matchDataCompleta) {
    return `${matchDataCompleta[2]}/${matchDataCompleta[3]}`;
  }

  // Nome do mês por extenso + ano (ex: "Maio de 2026", "Maio/2026")
  const meses = {
    'janeiro': '01', 'fevereiro': '02', 'marco': '03', 'abril': '04',
    'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08', 'setembro': '09',
    'outubro': '10', 'novembro': '11', 'dezembro': '12'
  };
  const textoNormalizado = texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const matchMesExtenso = textoNormalizado.match(
    /\b(janeiro|fevereiro|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\b[^\d]{0,10}(20\d{2})\b/
  );
  if (matchMesExtenso) {
    return `${meses[matchMesExtenso[1]]}/${matchMesExtenso[2]}`;
  }

  return null;
}


/**
 * 6. AUXILIAR: PARSER NATIVO DE TEXTO EM PDF
 */
function extrairTextoNativoPDF(bytes) {
  let str = '';
  for (let i = 0; i < bytes.length; i++) {
    str += String.fromCharCode(bytes[i]);
  }

  const paginasBlocos = str.split(/\/Type\s*\/Page\b/g);
  const textosPaginas = [];

  for (let i = 1; i < paginasBlocos.length; i++) {
    const bloco = paginasBlocos[i];
    let textoAcumulado = '';

    const regexTextos = /\(([^)]+)\)/g;
    let match;
    while ((match = regexTextos.exec(bloco)) !== null) {
      textoAcumulado += match[1] + ' ';
    }

    textosPaginas.push(textoAcumulado);
  }

  return textosPaginas;
}

/**
 * 7. AUXILIAR: BUSCA DE NOME NORMALIZADA
 */
function buscarPaginaPorNome(textosPorPagina, nome) {
  const normalizar = (texto) => texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const nomeNormalizado = normalizar(nome);

  for (let i = 0; i < textosPorPagina.length; i++) {
    const textoPagina = normalizar(textosPorPagina[i]);
    if (textoPagina.includes(nomeNormalizado)) {
      return i + 1;
    }
  }

  return -1;
}

/**
 * 8. AUXILIAR: REGISTRO DE LOGS NA PLANILHA
 */
function registrarLog(abaLog, dataHora, id, nome, whatsapp, pagina, nomeArquivo, docId, link, status, detalhes) {
  abaLog.appendRow([
    dataHora,
    id,
    nome,
    whatsapp,
    pagina,
    nomeArquivo,
    docId,
    link,
    status,
    detalhes,
    ''
  ]);
}

/**
 * 9. AUXILIAR: ENVIO PARA AUTENTIQUE VIA WHATSAPP (com mês e ano do documento no título)
 * @param {Blob} pdfBlob PDF individual do funcionário
 * @param {Object} funcionario Dados do funcionário
 * @param {string} mesAnoCompetencia Competência do documento no formato "MM/AAAA" (ex: "05/2026")
 */
function enviarParaAutentiqueViaWhatsapp(pdfBlob, funcionario, mesAnoCompetencia) {
  const query = `
    mutation CreateDocumentMutation(
      $document: DocumentInput!,
      $signers: [SignerInput!]!,
      $file: Upload!
    ) {
      createDocument(
        document: $document,
        signers: $signers,
        file: $file
      ) {
        id
        name
        signatures {
          public_id
          name
          email
        }
      }
    }
  `;

  let numTelefone = funcionario.whatsapp.replace(/\D/g, '');
  if (numTelefone.length === 10 || numTelefone.length === 11) {
    numTelefone = '55' + numTelefone;
  }
  if (!numTelefone.startsWith('+')) {
    numTelefone = '+' + numTelefone;
  }

  // Mês/ano da competência do documento (detectado no PDF, ou o mês atual como último recurso)
  if (!mesAnoCompetencia) {
    mesAnoCompetencia = Utilities.formatDate(new Date(), "GMT-3", "MM/yyyy");
  }

  const operations = {
    query: query,
    variables: {
      document: {
        name: `Espelho de Ponto - ${funcionario.nome} (${mesAnoCompetencia})`
      },
      signers: [
        {
          name: funcionario.nome,
          phone: numTelefone,
          delivery_method: 'DELIVERY_METHOD_WHATSAPP',
          action: 'SIGN'
        }
      ],
      file: null
    }
  };

  const map = { '0': ['variables.file'] };

  const payload = {
    'operations': JSON.stringify(operations),
    'map': JSON.stringify(map),
    '0': pdfBlob
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${CONFIG.AUTENTIQUE_TOKEN}`
    },
    payload: payload,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(CONFIG.AUTENTIQUE_URL, options);
    return JSON.parse(response.getContentText());
  } catch (error) {
    return { erro: error.toString() };
  }
}

/**
 * 10. GESTÃO DE FUNCIONÁRIOS (Aba "Dados Base")
 */

// Lista todos os funcionários para a interface de gestão
function obterListaFuncionarios() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abaDados = spreadsheet.getSheetByName(CONFIG.ABA_DADOS_BASE);
    if (!abaDados) return [];

    const ultimaLinha = abaDados.getLastRow();
    if (ultimaLinha <= 1) return [];

    const valores = abaDados.getRange(2, 1, ultimaLinha - 1, 4).getValues();

    return valores.map((row, index) => ({
      linha: index + 2,
      nome: String(row[0] || '').trim(),
      whatsapp: String(row[1] || '').trim(),
      email: String(row[2] || '').trim(),
      status: String(row[3] || 'Ativo').trim()
    })).filter(f => f.nome !== '');

  } catch (e) {
    Logger.log(`❌ Erro ao obter funcionários: ${e.toString()}`);
    return [];
  }
}

// Adiciona um novo funcionário na planilha
function adicionarFuncionario(dados) {
  if (!dados.nome || !dados.whatsapp) {
    throw new Error('Nome Completo e WhatsApp são obrigatórios.');
  }

  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const abaDados = spreadsheet.getSheetByName(CONFIG.ABA_DADOS_BASE);
  if (!abaDados) throw new Error('Aba "Dados Base" não encontrada.');

  const whatsappLimpo = String(dados.whatsapp).replace(/\D/g, '').trim();

  abaDados.appendRow([
    dados.nome.trim(),
    whatsappLimpo,
    (dados.email || '').trim(),
    'Ativo'
  ]);

  return { sucesso: true, mensagem: `Funcionário "${dados.nome}" cadastrado com sucesso!` };
}

// Remove o funcionário selecionado da planilha
function removerFuncionario(linha) {
  const numLinha = parseInt(linha, 10);
  if (isNaN(numLinha) || numLinha < 2) throw new Error('Linha inválida para remoção.');

  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const abaDados = spreadsheet.getSheetByName(CONFIG.ABA_DADOS_BASE);
  if (!abaDados) throw new Error('Aba "Dados Base" não encontrada.');

  abaDados.deleteRow(numLinha);

  return { sucesso: true, mensagem: 'Funcionário removido com sucesso!' };
}

/**
 * ============================================================================
 * 11. DASHBOARD MENSAL (Enviados / Assinados / Pendentes / Erros)
 * ============================================================================
 */

// Nomes dos meses em português, usados para montar os rótulos do dashboard
// (Utilities.formatDate depende do locale configurado no projeto Apps Script,
// que nem sempre está em pt-BR, por isso os nomes são montados manualmente).
const NOMES_MESES_PT = [
  'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
];

/**
 * 11.1 MONTA O RESUMO MENSAL A PARTIR DA ABA DE LOG, CONSULTANDO A AUTENTIQUE
 * PARA SABER O STATUS DE ASSINATURA DE CADA DOCUMENTO (COM CACHE).
 *
 * @param {boolean} forcarAtualizacao Se true, ignora o cache e consulta a Autentique
 *                                    novamente para cada documento (usado no botão "Atualizar").
 */
function obterDadosDashboard(forcarAtualizacao) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abaLog = spreadsheet.getSheetByName(CONFIG.ABA_LOG);
    if (!abaLog) throw new Error('Aba "Log de Execução" não encontrada.');

    const ultimaLinha = abaLog.getLastRow();
    if (ultimaLinha <= 1) {
      return { meses: [], totais: { enviados: 0, assinados: 0, pendentes: 0, erros: 0 } };
    }

    const valores = abaLog.getRange(2, 1, ultimaLinha - 1, 11).getValues();
    const mapaMeses = {};

    valores.forEach(row => {
      const dataRaw = row[0];
      const status = String(row[8] || '').trim();
      const docId = String(row[6] || '').trim();

      let dataObj;
      if (dataRaw instanceof Date) {
        dataObj = dataRaw;
      } else {
        const partes = String(dataRaw).match(/(\d{2})\/(\d{2})\/(\d{4})/);
        if (!partes) return; // linha sem data válida, ignora
        dataObj = new Date(Number(partes[3]), Number(partes[2]) - 1, Number(partes[1]));
      }

      // Mês/ano calculados no fuso GMT-3, mas o rótulo é montado manualmente em português
      const mesNumero = parseInt(Utilities.formatDate(dataObj, "GMT-3", "MM"), 10); // 1-12
      const anoNumero = Utilities.formatDate(dataObj, "GMT-3", "yyyy");
      const chave = `${anoNumero}-${String(mesNumero).padStart(2, '0')}`;
      const label = `${NOMES_MESES_PT[mesNumero - 1]} de ${anoNumero}`;

      if (!mapaMeses[chave]) {
        mapaMeses[chave] = { chave, label, enviados: 0, assinados: 0, pendentes: 0, erros: 0 };
      }

      // Não foi enviado com sucesso (erro no processamento/WhatsApp)
      if (status !== 'Sucesso' || !docId) {
        mapaMeses[chave].erros++;
        return;
      }

      mapaMeses[chave].enviados++;

      const statusAssinatura = obterStatusAssinaturaComCache(docId, forcarAtualizacao);
      if (statusAssinatura.todosAssinados) {
        mapaMeses[chave].assinados++;
      } else {
        mapaMeses[chave].pendentes++;
      }
    });

    const meses = Object.values(mapaMeses).sort((a, b) => a.chave.localeCompare(b.chave));

    const totais = meses.reduce((acc, m) => {
      acc.enviados += m.enviados;
      acc.assinados += m.assinados;
      acc.pendentes += m.pendentes;
      acc.erros += m.erros;
      return acc;
    }, { enviados: 0, assinados: 0, pendentes: 0, erros: 0 });

    return { meses, totais };

  } catch (e) {
    Logger.log(`❌ Erro ao montar dashboard: ${e.toString()}`);
    throw new Error(`Erro ao montar dashboard: ${e.message}`);
  }
}

/**
 * 11.2 CONSULTA STATUS DE ASSINATURA COM CACHE (evita sobrecarregar a API da Autentique)
 * Documentos já 100% assinados ficam em cache por mais tempo, pois o status é definitivo.
 *
 * @param {string} docId ID do documento na Autentique
 * @param {boolean} forcarAtualizacao Se true, ignora a leitura do cache e consulta direto na API
 */
function obterStatusAssinaturaComCache(docId, forcarAtualizacao) {
  const cache = CacheService.getScriptCache();
  const chaveCache = `assinatura_${docId}`;

  if (!forcarAtualizacao) {
    const cacheado = cache.get(chaveCache);
    if (cacheado) return JSON.parse(cacheado);
  }

  let resultado = { todosAssinados: false, algumAssinado: false };

  try {
    const detalhes = obterDetalhesDocumento(docId);
    const signatarios = detalhes.signatarios || [];

    if (signatarios.length > 0) {
      resultado.todosAssinados = signatarios.every(s => s.assinado);
      resultado.algumAssinado = signatarios.some(s => s.assinado);
    }
  } catch (e) {
    Logger.log(`⚠️ Não foi possível verificar assinatura do doc ${docId}: ${e.toString()}`);
  }

  const ttl = resultado.todosAssinados ? 21600 : 900; // 6h se assinado, 15min se pendente
  cache.put(chaveCache, JSON.stringify(resultado), ttl);

  return resultado;
}