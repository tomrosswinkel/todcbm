function manualChatPost() {

  //Modify these 3 Variables every time script is run
  FTR_SHEET_ID = '1r32HIhlxySQnAb8yFJaRQclS3wo22ASEvH8ybBQWC80';
  FTR_SLIDES_ID = '1GiM0Thkylnefn6ft46x0eQGVS4w_wwM3-OBJadwx1Xc';
  CONTEXT = 'Testing';



  const sheetUrl = `https://docs.google.com/spreadsheets/d/${FTR_SHEET_ID}`;
  const slidesUrl = `https://docs.google.com/presentation/d/${FTR_SLIDES_ID}`;

  const card = {
    'cardsV2': [{
      'cardId': 'updateCompleteCard',
      'card': {
        'header': {
          'title': 'New FTR Data imported',
          'subtitle': `Successful run in ${CONTEXT}.`,
          'imageUrl': 'https://upload.wikimedia.org/wikipedia/commons/thumb/1/12/Google_Drive_icon_%282020%29.svg/240px-Google_Drive_icon_%282020%29.svg.png',
          'imageType': 'CIRCLE'
        },
        'sections': [{
          'widgets': [{
            'buttonList': {
              'buttons': [
                {
                  'text': 'View Sheet',
                  'onClick': { 'openLink': { 'url': sheetUrl } }
                },
                {
                  'text': 'View Slides',
                  'onClick': { 'openLink': { 'url': slidesUrl } }
                }
              ]
            }
          }]
        }]
      }
    }]
  };

  const payload = JSON.stringify(card);

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: payload,
  };

  UrlFetchApp.fetch(GOOGLE_CHAT_WEBHOOK_LINK, options);
  console.log("Successfully sent card message to Google Chat.");
}


