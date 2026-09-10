/**
 * RotaPet - Transporte Executivo de Pets
 * Google Apps Script para Registro de Pedidos,
 * Calculo de Orcamento e Notificacao Oracio
 * 
 * Calculo Logistico:
 * - Distancia Real: Integrado nativamente com a API do Google Maps
 *   (Maps.newDirectionFinder)
 *   Calcula rota rodoviaria real para qualquer cidade do Brasil
 * 
 * Custos Operacionais:
 * - Combustivel: 13,5 km/L a R$ 5,95/L (apenas ida)
 * - Pedagios: R$ 14,50 / 100 km (apenas ida)
 * - Aluguel de Carro: R$ 150,00 / dia
 * - Alimentacao Estrada: R$ 70,00 / dia
 * - Hotel Pernoite Pet-Friendly: R$ 180,00 / noite (viagens > 1 dia)
 * - Margem Liquida Minima do Roberto: R$ 400,00 / dia por cao
 */

// =========================================================================
// CONFIGURACOES DE NOTIFICACAO (ORACIO / WEBHOOK / TELEGRAM)
// =========================================================================
// 1. Webhook HTTP do Oracio (ou n8n, Make, endpoint customizado):
// Insira a URL do webhook do Oracio caso utilize webhook HTTP.
var ORACIO_WEBHOOK_URL = ""; 

// Se o webhook exigir token de autorizacao (Bearer Token):
var ORACIO_AUTH_TOKEN = "";

// 2. Telegram direto via Oracio (@Ooracio_bot -> Roberto):
var TELEGRAM_BOT_TOKEN = "8826892591:AAFddcPYwUVkif1RCaLvy-68iLmqin7JNwc";
var TELEGRAM_CHAT_ID = "1221813958";

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
    var eAbaOrc = nomeAba === "Pedidos de Orçamento";
    var eAbaPro = nomeAba === "Prospecção de Clientes";
    if (!eAbaOrc && !eAbaPro) {
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
    var tz = "America/Sao_Paulo";
    var dataHora = Utilities.formatDate(agora, tz, "dd/MM/yyyy HH:mm");
    
    var nome = data.nome || "Cliente";
    var whatsapp = data.whatsapp || "";
    var perfil = data.perfil || "Tutor";

    var cidOrig = (data.cidadeOrigem || "").trim();
    var ufOrig = (data.ufOrigem || "").trim().toUpperCase();
    var cidDest = (data.cidadeDestino || "").trim();
    var ufDest = (data.ufDestino || "").trim().toUpperCase();

    var padraoOrig = "São Paulo - SP";
    var padraoDest = "Rio de Janeiro - RJ";
    var textoOrig = cidOrig ? (cidOrig + " - " + (ufOrig || "SP")) : padraoOrig;
    var textoDest = cidDest ? (cidDest + " - " + (ufDest || "RJ")) : padraoDest;
    var origem = data.origem || textoOrig;
    var destino = data.destino || textoDest;

    if (!ufDest && destino.indexOf(" - ") !== -1) {
      ufDest = destino.split(" - ")[1].trim().toUpperCase();
    }

    var raca = data.raca || "Pet";
    var porte = data.porte || "Médio";
    var qtdPets = parseInt(data.qtdPets || "1", 10) || 1;
    var vacinasDoc = data.vacinasDoc || "Sim (Tudo em dia)";
    var modalidade = data.modalidade || "Vaga Executiva";
    var dataPrevista = data.dataPrevista || "A combinar";

    // Distancia real via Google Maps
    var estimativaKm = calcularDistanciaReal(origem, destino, ufDest);
    
    // Dias de estrada
    var diasViagem = 1;
    if (estimativaKm > 1350) {
      diasViagem = Math.ceil(estimativaKm / 650);
    } else if (estimativaKm > 650) {
      diasViagem = 2;
    }

    // Custos operacionais (calculados apenas para o trajeto de ida)
    var combustivel = (estimativaKm / 13.5) * 5.95;
    var pedagios = (estimativaKm / 100) * 14.50;
    var combustivelPedagio = combustivel + pedagios;
    var aluguelCarro = diasViagem * 150.00;
    var alimentacao = diasViagem * 70.00;
    var pernoiteHotel = (diasViagem > 1) ? (diasViagem - 1) * 180.00 : 0;

    var custoOperacionalTotal = combustivelPedagio + aluguelCarro +
      alimentacao + pernoiteHotel;
    var margemRoberto = diasViagem * qtdPets * 400.00;

    var ehExclusivo = modalidade.toLowerCase().indexOf("exclusivo") !== -1;
    var fatorModalidade = ehExclusivo ? 1.30 : 1.0;
    var totalBruto = custoOperacionalTotal + margemRoberto;
    var valorSugerido = totalBruto * fatorModalidade;
    valorSugerido = Math.ceil(valorSugerido / 50) * 50;

    var valorFormatado = formatarMoeda(valorSugerido);
    var sufixoPet = (qtdPets > 1) ? " filhotes " : " filhote ";
    var nomeRaca = raca ? raca.replace(/^filhotes?\s+/i, "") : "pet";
    var petDescricao = qtdPets + sufixoPet + nomeRaca;

    // Mensagem padrao de orcamento para envio ao cliente
    var parte1 = [
      "Olá, ",
      nome,
      "! Aqui é o Roberto Hora, da RotaPet ",
      "\u2014 transporte executivo de filhotes."
    ].join("");

    var parte2 = [
      "Recebi sua solicitação: transporte de ",
      origem,
      " até ",
      destino,
      ", para ",
      petDescricao,
      "."
    ].join("");

    var parte3 = [
      "O valor do transporte dedicado e climatizado é de ",
      valorFormatado,
      ", com paradas a cada 2h, ",
      "vídeos ao vivo e acompanhamento por GPS."
    ].join("");

    var parte4 = "Vamos falar sobre a data de retirada?";
    var msgPronta = [parte1, parte2, parte3, parte4].join("\n");

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
      formatarMoeda(combustivelPedagio),
      formatarMoeda(aluguelCarro),
      formatarMoeda(alimentacao),
      formatarMoeda(pernoiteHotel),
      formatarMoeda(custoOperacionalTotal),
      formatarMoeda(margemRoberto),
      valorFormatado,
      msgPronta
    ];

    sheet.appendRow(newRow);

    // Notificacao Oracio / Telegram / Webhook
    notificarOracio({
      dataHora: dataHora,
      nome: nome,
      whatsapp: whatsapp,
      perfil: perfil,
      origem: origem,
      destino: destino,
      raca: raca,
      petDescricao: petDescricao,
      porte: porte,
      qtdPets: qtdPets,
      vacinasDoc: vacinasDoc,
      modalidade: modalidade,
      dataPrevista: dataPrevista,
      estimativaKm: estimativaKm,
      diasViagem: diasViagem,
      valorFormatado: valorFormatado,
      msgPronta: msgPronta
    });

    var respOk = { status: "success", orcamento: valorSugerido };
    var jsonOk = JSON.stringify(respOk);
    return ContentService.createTextOutput(jsonOk)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    var respErr = { status: "error", message: error.toString() };
    var jsonErr = JSON.stringify(respErr);
    return ContentService.createTextOutput(jsonErr)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("RotaPet Webhook ativo.");
}

function notificarOracio(dados) {
  // 1. Webhook HTTP do Oracio
  if (ORACIO_WEBHOOK_URL && ORACIO_WEBHOOK_URL.indexOf("http") === 0) {
    try {
      var headers = {
        "Content-Type": "application/json"
      };
      if (ORACIO_AUTH_TOKEN) {
        headers["Authorization"] = "Bearer " + ORACIO_AUTH_TOKEN;
      }

      var payload = {
        titulo: "Nova Solicitação Taxi Dog",
        dataHora: dados.dataHora,
        cliente: dados.nome,
        whatsapp: dados.whatsapp,
        perfil: dados.perfil,
        origem: dados.origem,
        destino: dados.destino,
        pet: dados.petDescricao,
        porte: dados.porte,
        vacinasDoc: dados.vacinasDoc,
        modalidade: dados.modalidade,
        dataPrevista: dados.dataPrevista,
        distanciaKm: dados.estimativaKm,
        diasEstimados: dados.diasViagem,
        valor: dados.valorFormatado,
        mensagem: dados.msgPronta
      };

      UrlFetchApp.fetch(ORACIO_WEBHOOK_URL, {
        method: "post",
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
    } catch (errWebhook) {
      Logger.log("Erro ao enviar webhook Oracio: " + errWebhook.toString());
    }
  }

  // 2. Notificacao Direta no Telegram via Oracio
  if (TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) {
    try {
      var textoTelegram = "🚗 *Nova Solicitação Taxi Dog*\n\n" +
        "👤 *Cliente:* " + dados.nome + "\n" +
        "📱 *WhatsApp:* " + dados.whatsapp + "\n" +
        "📍 *Rota:* " + dados.origem + " ➔ " + dados.destino + "\n" +
        "📏 *Distância:* " + dados.estimativaKm + " km\n" +
        "🐾 *Pet:* " + dados.petDescricao + "\n" +
        "💉 *Vacinas/Doc:* " + dados.vacinasDoc + "\n" +
        "💰 *Valor Calculado:* " + dados.valorFormatado;

      var numTelefone = (dados.whatsapp || "").replace(/\D/g, "");
      if (numTelefone.indexOf("0") === 0) {
        numTelefone = numTelefone.substring(1);
      }
      if (numTelefone.length === 10 || numTelefone.length === 11) {
        numTelefone = "55" + numTelefone;
      }

      var payloadTelegram = {
        chat_id: TELEGRAM_CHAT_ID,
        text: textoTelegram,
        parse_mode: "Markdown"
      };

      if (numTelefone) {
        var txtZap = encodeURIComponent(dados.msgPronta);
        var zapBase = "https://api.whatsapp.com/send?phone=";
        var zapUrl = zapBase + numTelefone + "&text=" + txtZap;
        payloadTelegram.reply_markup = {
          inline_keyboard: [
            [
              {
                text: "📲 Enviar Orçamento no WhatsApp",
                url: zapUrl
              }
            ]
          ]
        };
      }

      var tgBase = "https://api.telegram.org/bot";
      var telegramUrl = tgBase + TELEGRAM_BOT_TOKEN + "/sendMessage";
      UrlFetchApp.fetch(telegramUrl, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payloadTelegram),
        muteHttpExceptions: true
      });
    } catch (errTelegram) {
      Logger.log("Erro ao enviar Telegram: " + errTelegram.toString());
    }
  }
}

function formatarMoeda(valor) {
  if (typeof valor !== "number") valor = Number(valor) || 0;
  var partes = valor.toFixed(2).split(".");
  partes[0] = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return "R$ " + partes.join(",");
}

function calcularDistanciaReal(origem, destino, ufDestino) {
  try {
    var origemQuery = formatarEnderecoMaps(origem);
    var destinoQuery = formatarEnderecoMaps(destino);

    var directions = Maps.newDirectionFinder()
      .setOrigin(origemQuery)
      .setDestination(destinoQuery)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (directions && directions.routes && directions.routes.length > 0) {
      var route = directions.routes[0];
      var totalMetros = 0;
      for (var i = 0; i < route.legs.length; i++) {
        totalMetros += route.legs[i].distance.value;
      }
      var km = Math.round(totalMetros / 1000);

      // Trava de seguranca de estado
      if (ufDestino && route.legs && route.legs.length > 0) {
        var ultLeg = route.legs[route.legs.length - 1];
        var endAddr = (ultLeg.end_address || "").toUpperCase();
        var ufAlvo = ufDestino.trim().toUpperCase();
        var estAlvo = (nomeDoEstado(ufAlvo) || "").toUpperCase();

        var temUf = endAddr.indexOf(ufAlvo) !== -1;
        var temEst = endAddr.indexOf(estAlvo) !== -1;

        if (!temUf && !temEst) {
          return calcularDistanciaFallback(origem, ufAlvo);
        }
      }

      if (km > 0) return km;
    }
  } catch (e) {
  }

  return calcularDistanciaFallback(origem, destino);
}

function formatarEnderecoMaps(texto) {
  if (!texto) return "Brasil";
  var t = texto.trim();
  if (t.indexOf(" - ") !== -1) {
    var partes = t.split(" - ");
    var cidade = partes[0].trim();
    var uf = partes[1].trim().toUpperCase();
    var estado = nomeDoEstado(uf);
    return cidade + ", Estado de " + estado + ", Brasil";
  }
  return t + ", Brasil";
}

function nomeDoEstado(uf) {
  var mapa = {
    "AC": "Acre", "AL": "Alagoas", "AP": "Amapá", "AM": "Amazonas",
    "BA": "Bahia", "CE": "Ceará", "DF": "Distrito Federal",
    "ES": "Espírito Santo", "GO": "Goiás", "MA": "Maranhão",
    "MT": "Mato Grosso", "MS": "Mato Grosso do Sul", "MG": "Minas Gerais",
    "PA": "Pará", "PB": "Paraíba", "PR": "Paraná",
    "PE": "Pernambuco", "PI": "Piauí", "RJ": "Rio de Janeiro",
    "RN": "Rio Grande do Norte", "RS": "Rio Grande do Sul",
    "RO": "Rondônia", "RR": "Roraima", "SC": "Santa Catarina",
    "SP": "São Paulo", "SE": "Sergipe", "TO": "Tocantins"
  };
  return mapa[uf] || uf;
}

function calcularDistanciaFallback(origem, destino) {
  var orig = removerAcentos(origem.toLowerCase());
  var dest = removerAcentos(destino.toLowerCase());

  if (orig.indexOf("sao paulo") !== -1 || orig.indexOf("sp") !== -1) {
    var eRj = dest.indexOf("rio de janeiro") !== -1 ||
      dest.indexOf("rj") !== -1;
    if (eRj) return 435;
    var ePr = dest.indexOf("curitiba") !== -1 || dest.indexOf("pr") !== -1;
    if (ePr) return 410;
    var eMg = dest.indexOf("belo horizonte") !== -1 ||
      dest.indexOf("mg") !== -1;
    if (eMg) return 585;
    var eInt = dest.indexOf("campinas") !== -1 || dest.indexOf("ribeir") !== -1;
    if (eInt) return 240;
    var eLit = dest.indexOf("santos") !== -1 || dest.indexOf("litoral") !== -1;
    if (eLit) return 95;
    var ePi = dest.indexOf("teresina") !== -1 || dest.indexOf("pi") !== -1;
    if (ePi) return 2750;
    var eDf = dest.indexOf("brasilia") !== -1 || dest.indexOf("df") !== -1;
    if (eDf) return 1010;
    var eSc = dest.indexOf("florian") !== -1 || dest.indexOf("sc") !== -1;
    if (eSc) return 705;
    var eRs = dest.indexOf("porto alegre") !== -1 || dest.indexOf("rs") !== -1;
    if (eRs) return 1120;
    var eBa = dest.indexOf("salvador") !== -1 || dest.indexOf("ba") !== -1;
    if (eBa) return 1950;
  }
  return 450;
}

function removerAcentos(texto) {
  if (!texto) return "";
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function testarEnvioTelegram() {
  var tgUrl = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN;
  var telegramUrl = tgUrl + "/sendMessage";
  var payload = {
    chat_id: TELEGRAM_CHAT_ID,
    text: "🚗 *Oracio conectado com sucesso à planilha RotaPet!*",
    parse_mode: "Markdown"
  };
  UrlFetchApp.fetch(telegramUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  });
}
