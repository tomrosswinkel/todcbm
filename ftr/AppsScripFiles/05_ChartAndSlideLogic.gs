 
/**
 * @file 05_ChartAndSlideLogic.gs
 * @description Logic for updating visuals and sending success notifications.
 * Updated to use bulletproof, high-reliability proxy-safe success icons and titles.
 */

function updateCharts(ss) {
  console.time("updateCharts");
  const currentDate = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd");
  const lookupSheet = ss.getSheetByName(LOOKUP_SHEET_NAME);
  const titleValues = lookupSheet.getRange("M1:M" + lookupSheet.getLastRow()).getValues();

  const chartConfigs = [
    ["Other Statistics", 51, "0"], ["Other Statistics", 121, "1"], ["Other Statistics", 123, "3"], ["Other Statistics", 124, "2"], ["Rack Burndown", 6, "0"], ["Rack Burndown", 10, "1"],
    ["Rack Burndown", 5, "2"], ["Rack Burndown", 9, "3"], ["Rack Burndown", 14, "4"], ["Rack Burndown", 53, "5"], ["Rack Burndown", 54, "6"], ["Rack Burndown", 55, "7"], ["Rack Burndown", 56, "8"],
    ["Rack Burndown", 12, "9"], ["Rack Burndown", 76, "10"], ["Rack Burndown", 77, "11"], ["Rack Burndown", 78, "12"], ["Rack Burndown", 79, "13"], ["PDU Burndown", 16, "0"], ["PDU Burndown", 20, "1"],
    ["PDU Burndown", 22, "2"], ["PDU Burndown", 18, "3"], ["PDU Burndown", 24, "4"], ["PDU Burndown", 26, "5"], ["PDU Burndown", 58, "6"], ["PDU Burndown", 59, "7"], ["PDU Burndown", 60, "8"],
    ["PDU Burndown", 61, "9"], ["PDU Burndown", 88, "10"], ["PDU Burndown", 89, "11"], ["PDU Burndown", 90, "12"], ["PDU Burndown", 91, "13"], ["Breaker Burndown", 30, "0"], ["Breaker Burndown", 34, "1"],
    ["Breaker Burndown", 29, "2"], ["Breaker Burndown", 33, "3"], ["Breaker Burndown", 100, "4"], ["Breaker Burndown", 62, "5"], ["Breaker Burndown", 63, "6"], ["Breaker Burndown", 66, "7"],
    ["Breaker Burndown", 65, "8"], ["Breaker Burndown", 64, "9"], ["Breaker Burndown", 36, "10"], ["Breaker Burndown", 38, "11"], ["Breaker Burndown", 101, "12"], ["Breaker Burndown", 102, "13"],
    ["Breaker Burndown", 103, "14"], ["Cabling Burndown", 46, "0"], ["Cabling Burndown", 42, "1"], ["Cabling Burndown", 40, "2"], ["Cabling Burndown", 44, "3"], ["Cabling Burndown", 48, "4"],
    ["Cabling Burndown", 67, "5"], ["Cabling Burndown", 68, "6"], ["Cabling Burndown", 71, "7"], ["Cabling Burndown", 70, "8"], ["Cabling Burndown", 69, "9"], ["Cabling Burndown", 112, "10"],
    ["Cabling Burndown", 113, "11"], ["Cabling Burndown", 114, "12"], ["Cabling Burndown", 115, "13"]
  ];

  const chartsBySheet = chartConfigs.reduce((acc, config) => {
    if (!acc[config[0]]) acc[config[0]] = [];
    acc[config[0]].push({ titleRow: config[1], chartIndex: config[2] });
    return acc;
  }, {});

  for (const sheetName in chartsBySheet) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const allCharts = sheet.getCharts();
    chartsBySheet[sheetName].forEach(config => {
      const chart = allCharts[config.chartIndex];
      if (chart) {
        const builder = chart.modify().setOption('title', titleValues[config.titleRow - 1][0])
          .setOption('hAxis.title', "Export Date: " + currentDate).setOption('hAxis.titleTextStyle', { color: "black", italic: false });
        sheet.updateChart(builder.build());
      }
    });
  }
  console.timeEnd("updateCharts");
}

function updateSlides(slidesId) {
  try {
    const charts = SlidesApp.openById(slidesId).getSlides().flatMap(s => s.getSheetsCharts());
    charts.forEach(c => c.refresh());
  } catch (e) { Logger.log(`Slides Error: ${e.message}`); }
}

/**
 * Sends a success notification, respecting the global toggle.
 * Uses high-reliability checkmark icon and native emoji header titles.
 */
function sendCardMessageToChat(sheetId, slidesId, context) {
  if (!ENABLE_CHAT_NOTIFICATIONS) {
    Logger.log("Chat notifications are disabled in Config.gs. Skipping success notification.");
    return;
  }

  // High-reliability checkmark icon hosted on Wikimedia CDN (proxy-safe PNG)
  const successIconUrl = "https://upload.wikimedia.org/wikipedia/commons/thumb/7/73/Flat_tick_icon.svg/240px-Flat_tick_icon.svg.png";

  const payload = JSON.stringify({
    'cardsV2': [{
      'cardId': 'FTR',
      'card': {
        'header': {
          'title': '🟢 FTR Import Complete',
          'subtitle': `Successful run in ${context}.`,
          'imageUrl': successIconUrl,
          'imageType': 'CIRCLE'
        },
        'sections': [{
          'widgets': [{
            'buttonList': {
              'buttons': [
                { 'text': 'View Sheet', 'onClick': { 'openLink': { 'url': `https://docs.google.com/spreadsheets/d/${sheetId}` } } },
                { 'text': 'View Slides', 'onClick': { 'openLink': { 'url': `https://docs.google.com/presentation/d/${slidesId}` } } }
              ]
            }
          }]
        }]
      }
    }]
  });
  try {
    UrlFetchApp.fetch(GOOGLE_CHAT_WEBHOOK_LINK, { method: 'POST', contentType: 'application/json', payload: payload });
  } catch (e) {
    Logger.log(`Failed to send Chat success notification: ${e.message}`);
    try {
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        '🟢 FTR Import Complete (Chat notification failed)',
        `The FTR workflow completed successfully in ${context}, but the Google Chat notification failed.\n\nError: ${e.message}\n\nView Sheet: https://docs.google.com/spreadsheets/d/${sheetId}\nView Slides: https://docs.google.com/presentation/d/${slidesId}`
      );
    } catch (mailError) {
      Logger.log(`Email fallback also failed: ${mailError.message}`);
    }
  }
}
