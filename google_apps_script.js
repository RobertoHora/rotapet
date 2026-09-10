/**
 * RotaPet - Transporte Executivo de Pets
 * Google Apps Script para Registro de Pedidos e Calculo de Orcamento
 * 
 * Custos Operacionais Incluidos:
 * - Combustivel: 13,5 km/L a R$ 5,95/L (ida e volta)
 * - Pedagios: R$ 14,50 / 100 km (ida e volta)
 * - Aluguel de Carro: R$ 150,00 / dia
 * - Alimentacao Estrada: R$ 70,00 / dia
 * - Hotel Pernoite Pet-Friendly: R$ 180,00 / noite (viagens > 1 dia)
 * - Margem Liquida Minima do Roberto: R$ 300,00 / dia por cao
 */

function setupPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetOrcamentos = ss.getSheetByName('Pedidos de Orçamento');
  if (!sheetOrcamentos) {
    sheetOrcamentos = ss.insertSheet('Pedidos de Orçamento', 0);
  } else {
    ss.setActiveSheet(sheetOrcamentos);
    ss.moveActiveSheet(1);
  }
  
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== 'Pedidos de Orçamento' && sheets[i].getName() !== 'Prospecção de Clientes') {
      sheets[i].setName('Prospecção de Clientes');
      break;
    }
  }

  var headers = [
    'Data/Hora',
    'Nome do Cliente',
    'WhatsApp',
    'Perfil',
    'Origem (Coleta)',
    'Destino (Entrega)',
    'Espécie / Raça',
    'Porte',
    'Qtd Pets',
    'Modalidade',
    'Data Prevista',
    'Distância Estimada (km)',
    'Duração Estimada (dias)',
    'Combustível + Pedágio (R$)',
    'Aluguel Carro (R$)',
    'Alimentação (R$)',
    'Hotel / Pernoite (R$)',
    'Custo Operacional Total (R$)',
    'Sua Margem Líquida Mínima (R$)',
    'Valor Total Sugerido (R$)',
    'Mensagem Pronta para WhatsApp'
  ];

  sheetOrcamentos.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheetOrcamentos.getRange(1, 1, 1, headers.length)
    .setBackground('#0d121c')
    .setFontColor('#c59b4c')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setHorizontalAlignment('center');
  
  sheetOrcamentos.setFrozenRows(1);
  for (var col = 1; col <= headers.length; col++) {
    sheetOrcamentos.autoResizeColumn(col);
  }
}

function doPost(e) {
  try {
    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Pedidos de Orçamento');
    if (!sheet) {
      setupPlanilha();
      sheet = ss.getSheetByName('Pedidos de Orçamento');
    }

    var headers = [
      'Data/Hora',
      'Nome do Cliente',
      'WhatsApp',
      'Perfil',
      'Origem (Coleta)',
      'Destino (Entrega)',
      'Espécie / Raça',
      'Porte',
      'Qtd Pets',
      'Modalidade',
      'Data Prevista',
      'Distância Estimada (km)',
      'Duração Estimada (dias)',
      'Combustível + Pedágio (R$)',
      'Aluguel Carro (R$)',
      'Alimentação (R$)',
      'Hotel / Pernoite (R$)',
      'Custo Operacional Total (R$)',
      'Sua Margem Líquida Mínima (R$)',
      'Valor Total Sugerido (R$)',
      'Mensagem Pronta para WhatsApp'
    ];

    // Atualiza cabeçalhos automaticamente se houver novas colunas
    if (sheet.getLastColumn() < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#0d121c')
        .setFontColor('#c59b4c')
        .setFontWeight('bold')
        .setFontFamily('Arial')
        .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }

    var agora = new Date();
    var dataHora = Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
    
    var nome = data.nome || 'Cliente';
    var whatsapp = data.whatsapp || '';
    var perfil = data.perfil || 'Tutor';
    var origem = data.origem || 'São Paulo - SP';
    var destino = data.destino || 'Destino';
    var raca = data.raca || 'Pet';
    var porte = data.porte || 'Médio';
    var qtdPets = parseInt(data.qtdPets || '1', 10) || 1;
    var modalidade = data.modalidade || 'Vaga Executiva';
    var dataPrevista = data.dataPrevista || 'A combinar';

    // --- CÁLCULO LOGÍSTICO E FINANCEIRO (INTERNO) ---
    var estimativaKm = calcularDistanciaAproximada(origem, destino);
    
    // Dias de estrada (até 650km = 1 dia; até 1350km = 2 dias; acima = 3+ dias)
    var diasViagem = 1;
    if (estimativaKm > 1350) {
      diasViagem = Math.ceil(estimativaKm / 650);
    } else if (estimativaKm > 650) {
      diasViagem = 2;
    }

    // 1. Combustível + Pedágio (ida e volta)
    var combustivel = (estimativaKm * 2 / 13.5) * 5.95;
    var pedagios = (estimativaKm * 2 / 100) * 14.50;
    var combustivelPedagio = combustivel + pedagios;

    // 2. Aluguel do carro (R$ 150/dia)
    var aluguelCarro = diasViagem * 150.00;

    // 3. Alimentação na estrada (R$ 70/dia)
    var alimentacao = diasViagem * 70.00;

    // 4. Hotel / Pernoite Pet-Friendly (R$ 180/noite quando > 1 dia)
    var pernoiteHotel = (diasViagem > 1) ? (diasViagem - 1) * 180.00 : 0;

    // Custo Operacional Total
    var custoOperacionalTotal = combustivelPedagio + aluguelCarro + alimentacao + pernoiteHotel;

    // Margem líquida mínima do Roberto: R$ 300 por dia por cão
    var margemRoberto = diasViagem * qtdPets * 300.00;

    // Se modalidade for 100% Exclusivo, ajusta taxa de exclusividade (30%)
    var fatorModalidade = (modalidade.toLowerCase().indexOf('exclusivo') !== -1) ? 1.30 : 1.0;
    var valorSugerido = (custoOperacionalTotal + margemRoberto) * fatorModalidade;

    // Arredonda para múltiplo de 50 mais próximo
    valorSugerido = Math.ceil(valorSugerido / 50) * 50;

    // Mensagem pronta para copiar e colar no WhatsApp
    var msgPronta = 'Olá ' + nome + '! Aqui é o Roberto Hora do transporte executivo RotaPet. ' +
      'Recebi sua solicitação para o transporte de ' + origem + ' até ' + destino + ' (' + qtdPets + ' pet ' + raca + '). ' +
      'O valor para o transporte dedicado e climatizado fica em R$ ' + valorSugerido.toLocaleString('pt-BR') + 
      ' com paradas a cada 2h, vídeos ao vivo e acompanhamento por GPS. Podemos reservar para a data ' + dataPrevista + '?';

    // Linha completa com todos os custos discriminados
    var newRow = [
      dataHora,
      nome,
      whatsapp,
      perfil,
      origem,
      destino,
      raca,
      porte,
      qtdPets,
      modalidade,
      dataPrevista,
      estimativaKm,
      diasViagem,
      'R$ ' + combustivelPedagio.toFixed(2).replace('.', ','),
      'R$ ' + aluguelCarro.toFixed(2).replace('.', ','),
      'R$ ' + alimentacao.toFixed(2).replace('.', ','),
      'R$ ' + pernoiteHotel.toFixed(2).replace('.', ','),
      'R$ ' + custoOperacionalTotal.toFixed(2).replace('.', ','),
      'R$ ' + margemRoberto.toFixed(2).replace('.', ','),
      'R$ ' + valorSugerido.toFixed(2).replace('.', ','),
      msgPronta
    ];

    sheet.appendRow(newRow);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('RotaPet Webhook ativo.');
}

function calcularDistanciaAproximada(origem, destino) {
  var orig = origem.toLowerCase();
  var dest = destino.toLowerCase();

  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('rio de janeiro') !== -1 || dest.indexOf('rj') !== -1)) return 435;
  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('curitiba') !== -1 || dest.indexOf('pr') !== -1)) return 410;
  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('belo horizonte') !== -1 || dest.indexOf('mg') !== -1)) return 585;
  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('campinas') !== -1 || dest.indexOf('ribeir') !== -1)) return 240;
  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('santos') !== -1 || dest.indexOf('litoral') !== -1)) return 95;
  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('teresina') !== -1 || dest.indexOf('piauí') !== -1 || dest.indexOf('pi') !== -1)) return 2750;
  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('brasília') !== -1 || dest.indexOf('df') !== -1)) return 1010;
  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('florian') !== -1 || dest.indexOf('sc') !== -1)) return 705;
  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('porto alegre') !== -1 || dest.indexOf('rs') !== -1)) return 1120;
  if ((orig.indexOf('são paulo') !== -1 || orig.indexOf('sp') !== -1) && (dest.indexOf('salvador') !== -1 || dest.indexOf('ba') !== -1)) return 1950;
  
  return 450;
}
