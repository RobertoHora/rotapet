/**
 * RotaPet - Transporte Executivo de Pets
 * Google Apps Script para Registro de Pedidos e Calculo de Orcamento
 * 
 * Calculo Logistico:
 * - Distancia Real: Integrado nativamente com a API do Google Maps (Maps.newDirectionFinder)
 *   Calcula rota rodoviaria real para qualquer cidade do Brasil
 * 
 * Custos Operacionais:
 * - Combustivel: 13,5 km/L a R$ 5,95/L (ida e volta)
 * - Pedagios: R$ 14,50 / 100 km (ida e volta)
 * - Aluguel de Carro: R$ 150,00 / dia
 * - Alimentacao Estrada: R$ 70,00 / dia
 * - Hotel Pernoite Pet-Friendly: R$ 180,00 / noite (viagens > 1 dia)
 * - Margem Liquida Minima do Roberto: R$ 300,00 / dia por cao
 */

function setupPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetOrcamentos = ss.getSheetByName("Pedidos de Orçamento");
  if (!sheetOrcamentos) {
    sheetOrcamentos = ss.insertSheet("Pedidos de Orçamento", 0);
  } else {
    ss.setActiveSheet(sheetOrcamentos);
    ss.moveActiveSheet(1);
  }
  
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== "Pedidos de Orçamento" && sheets[i].getName() !== "Prospecção de Clientes") {
      sheets[i].setName("Prospecção de Clientes");
      break;
    }
  }

  var headers = [
    "Data/Hora",
    "Nome do Cliente",
    "WhatsApp",
    "Perfil",
    "Origem (Coleta)",
    "Destino (Entrega)",
    "Espécie / Raça",
    "Porte",
    "Qtd Pets",
    "Vacinas e Atestado",
    "Modalidade",
    "Data Prevista",
    "Distância Estimada (km)",
    "Duração Estimada (dias)",
    "Combustível + Pedágio (R$)",
    "Aluguel Carro (R$)",
    "Alimentação (R$)",
    "Hotel / Pernoite (R$)",
    "Custo Operacional Total (R$)",
    "Sua Margem Líquida Mínima (R$)",
    "Valor Total Sugerido (R$)",
    "Mensagem Pronta para WhatsApp"
  ];

  sheetOrcamentos.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheetOrcamentos.getRange(1, 1, 1, headers.length)
    .setBackground("#0d121c")
    .setFontColor("#c59b4c")
    .setFontWeight("bold")
    .setFontFamily("Arial")
    .setHorizontalAlignment("center");
  
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
    var sheet = ss.getSheetByName("Pedidos de Orçamento");
    if (!sheet) {
      setupPlanilha();
      sheet = ss.getSheetByName("Pedidos de Orçamento");
    }

    var headers = [
      "Data/Hora",
      "Nome do Cliente",
      "WhatsApp",
      "Perfil",
      "Origem (Coleta)",
      "Destino (Entrega)",
      "Espécie / Raça",
      "Porte",
      "Qtd Pets",
      "Vacinas e Atestado",
      "Modalidade",
      "Data Prevista",
      "Distância Estimada (km)",
      "Duração Estimada (dias)",
      "Combustível + Pedágio (R$)",
      "Aluguel Carro (R$)",
      "Alimentação (R$)",
      "Hotel / Pernoite (R$)",
      "Custo Operacional Total (R$)",
      "Sua Margem Líquida Mínima (R$)",
      "Valor Total Sugerido (R$)",
      "Mensagem Pronta para WhatsApp"
    ];

    if (sheet.getLastColumn() < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground("#0d121c")
        .setFontColor("#c59b4c")
        .setFontWeight("bold")
        .setFontFamily("Arial")
        .setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    var agora = new Date();
    var dataHora = Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy HH:mm");
    
    var nome = data.nome || "Cliente";
    var whatsapp = data.whatsapp || "";
    var perfil = data.perfil || "Tutor";
    var origem = data.origem || "São Paulo - SP";
    var destino = data.destino || "Destino";
    var raca = data.raca || "Pet";
    var porte = data.porte || "Médio";
    var qtdPets = parseInt(data.qtdPets || "1", 10) || 1;
    var vacinasDoc = data.vacinasDoc || "Sim (Tudo em dia)";
    var modalidade = data.modalidade || "Vaga Executiva";
    var dataPrevista = data.dataPrevista || "A combinar";

    // Distancia real via Google Maps
    var estimativaKm = calcularDistanciaReal(origem, destino);
    
    // Dias de estrada
    var diasViagem = 1;
    if (estimativaKm > 1350) {
      diasViagem = Math.ceil(estimativaKm / 650);
    } else if (estimativaKm > 650) {
      diasViagem = 2;
    }

    // Custos operacionais
    var combustivel = (estimativaKm * 2 / 13.5) * 5.95;
    var pedagios = (estimativaKm * 2 / 100) * 14.50;
    var combustivelPedagio = combustivel + pedagios;
    var aluguelCarro = diasViagem * 150.00;
    var alimentacao = diasViagem * 70.00;
    var pernoiteHotel = (diasViagem > 1) ? (diasViagem - 1) * 180.00 : 0;

    var custoOperacionalTotal = combustivelPedagio + aluguelCarro + alimentacao + pernoiteHotel;
    var margemRoberto = diasViagem * qtdPets * 300.00;

    var fatorModalidade = (modalidade.toLowerCase().indexOf("exclusivo") !== -1) ? 1.30 : 1.0;
    var valorSugerido = (custoOperacionalTotal + margemRoberto) * fatorModalidade;
    valorSugerido = Math.ceil(valorSugerido / 50) * 50;

    var msgPronta = "Olá " + nome + "! Aqui é o Roberto Hora do transporte executivo RotaPet. " +
      "Recebi sua solicitação para o transporte de " + origem + " até " + destino + " (" + qtdPets + " pet " + raca + "). " +
      "O valor para o transporte dedicado e climatizado fica em R$ " + valorSugerido.toLocaleString("pt-BR") + 
      " com paradas a cada 2h, vídeos ao vivo e acompanhamento por GPS. Podemos reservar para a data " + dataPrevista + "?";

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
      vacinasDoc,
      modalidade,
      dataPrevista,
      estimativaKm,
      diasViagem,
      "R$ " + combustivelPedagio.toFixed(2).replace(".", ","),
      "R$ " + aluguelCarro.toFixed(2).replace(".", ","),
      "R$ " + alimentacao.toFixed(2).replace(".", ","),
      "R$ " + pernoiteHotel.toFixed(2).replace(".", ","),
      "R$ " + custoOperacionalTotal.toFixed(2).replace(".", ","),
      "R$ " + margemRoberto.toFixed(2).replace(".", ","),
      "R$ " + valorSugerido.toFixed(2).replace(".", ","),
      msgPronta
    ];

    sheet.appendRow(newRow);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("RotaPet Webhook ativo.");
}

function calcularDistanciaReal(origem, destino) {
  try {
    var directions = Maps.newDirectionFinder()
      .setOrigin(origem + ", Brasil")
      .setDestination(destino + ", Brasil")
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (directions && directions.routes && directions.routes.length > 0) {
      var route = directions.routes[0];
      var totalMetros = 0;
      for (var i = 0; i < route.legs.length; i++) {
        totalMetros += route.legs[i].distance.value;
      }
      var km = Math.round(totalMetros / 1000);
      if (km > 0) return km;
    }
  } catch (e) {
  }

  return calcularDistanciaFallback(origem, destino);
}

function calcularDistanciaFallback(origem, destino) {
  var orig = removerAcentos(origem.toLowerCase());
  var dest = removerAcentos(destino.toLowerCase());

  if (orig.indexOf("sao paulo") !== -1 || orig.indexOf("sp") !== -1) {
    if (dest.indexOf("rio de janeiro") !== -1 || dest.indexOf("rj") !== -1) return 435;
    if (dest.indexOf("curitiba") !== -1 || dest.indexOf("pr") !== -1) return 410;
    if (dest.indexOf("belo horizonte") !== -1 || dest.indexOf("mg") !== -1) return 585;
    if (dest.indexOf("campinas") !== -1 || dest.indexOf("ribeir") !== -1) return 240;
    if (dest.indexOf("santos") !== -1 || dest.indexOf("litoral") !== -1) return 95;
    if (dest.indexOf("teresina") !== -1 || dest.indexOf("piaui") !== -1 || dest.indexOf("pi") !== -1) return 2750;
    if (dest.indexOf("brasilia") !== -1 || dest.indexOf("df") !== -1) return 1010;
    if (dest.indexOf("florian") !== -1 || dest.indexOf("sc") !== -1) return 705;
    if (dest.indexOf("porto alegre") !== -1 || dest.indexOf("rs") !== -1) return 1120;
    if (dest.indexOf("salvador") !== -1 || dest.indexOf("ba") !== -1) return 1950;
  }
  return 450;
}

function removerAcentos(texto) {
  if (!texto) return "";
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
