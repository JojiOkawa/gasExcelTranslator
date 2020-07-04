/*
　Excelをスプレッドシートに変換して指定したフォルダに保存する関数
*/
function excelTranslator(e) {
  let itemResponse;
  let file;
  let folder;
  const itemResponses = e.response.getItemResponses();
  
  //Q1 フォームからアップロードされたファイルidを取得する。
  file = DriveApp.getFileById(itemResponses[0].getResponse());
  file.getId();
  
  //Q2 get folder id by form's
  folder = DriveApp.getFolderById(itemResponses[1].getResponse());
  options = {
    title: file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{id: folder.getId()}]
  };
  
  // Excelをスプレッドシートに変換する。Drive APIへfileをPOSTする
  file = Drive.Files.insert(options, file.getBlob());

  //Q3 Language before translation
  const beforeLanguage = itemResponses[2].getResponse();
  //Q4 Language after translation
  const afterLanguage = itemResponses[3].getResponse();
  //Q5 mail
  const mailAddress = itemResponses[4].getResponse();

  //Q7 oneSheet or multisheet 
  //"シート1枚目のみ翻訳"かそれ以外かの2択の必須回答
  const multishtFlg = itemResponses[6].getResponse();
  if (multishtFlg === "シングルシート"){
    var targetSheet = SpreadsheetApp.openById(file.id).getSheets()[0];
    var targetSheetId = targetSheet.getSheetId(); 
    Logger.log(targetSheetId); 
    Logger.log("この上に表示されるよ。"); 

    TranslateToFrench(targetSheet ,beforeLanguage, afterLanguage);
  }else{
    let sheets = SpreadsheetApp.openById(file.id).getSheets();
    sheets.forEach( (sheet) => {
      var targetSheet = sheet;
      TranslateToFrench(targetSheet ,beforeLanguage, afterLanguage);
    });
    var targetSheetId = ""
  }

  const resultXlsx = ss2xlsx(file.id,folder);
  const resultPdf = ss2pdf(file.id, targetSheetId, folder);

  if (mailAddress != ""){
    //q6 mailsubject
    const mailsubject = itemResponses[5].getResponse();
    sendmail(mailAddress, resultPdf, mailsubject);
  };

}

/*
　翻訳関数
*/
function TranslateToFrench(targetSheet,beforeLanguage,afterLanguage) {
  const maxrow = targetSheet.getLastRow();
  const maxcol = targetSheet.getLastColumn();
  var text ;
  for (var j = 1; j <= maxcol; j++){
    for (var i = 1; i <= maxrow; i++) {
      text = targetSheet.getRange(i, j).getValue();
      // if (text != ""){
      if (text != "" && typeof(text) === "string" ){
        var sourceDoc = targetSheet.getRange(i, j).getValue();
        var translate = LanguageApp.translate(sourceDoc, language[beforeLanguage], language[afterLanguage]);
        targetSheet.getRange(i, j).setValue(translate);
      }
    }
  } 
}

/*
　スプレッドシートをExcelに変換してdriveに保存する関数
*/
function ss2xlsx(spreadsheet_id,folder) {
  let new_file;
  const url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id + "/export?format=xlsx";
  const options = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() == 200) {
    let ss = SpreadsheetApp.openById(spreadsheet_id);
    let filename = ss.getName();
    let pos = filename.indexOf(" - ");
    filename = filename.substring(0, pos);
    new_file = folder.createFile(res.getBlob()).setName("翻訳済" + filename +  ".xlsx");
  }
  return new_file;
}

/*
　スプレッドシートをpdfに変換してdriveに保存する関数
*/
function ss2pdf(spreadsheet_id,targetSheetId ,folder) {
  let new_file;
  let url;
  if (targetSheetId === ""){
    url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id + "/export?exportFormat=pdf";
  }else{
    url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id + "/export?exportFormat=pdf&gid=SID".replace("SID",targetSheetId)
  }
  
  const options = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() == 200) {
    const ss = SpreadsheetApp.openById(spreadsheet_id);
    let filename = ss.getName();
    const pos = filename.indexOf(" - ");
    filename = filename.substring(0, pos);
    new_file = folder.createFile(res.getBlob()).setName("翻訳済" + filename +  ".pdf");
  }
  return new_file;
}

/*
　メールを送信する関数
*/
function sendmail(address,file,mailsubject){
  const bodymsg1 = "こんにちは。\nこのメールはGoogle Appsのプログラムから自動送信で送信しています。"
  const bodymsg2 = "Hello,dear.\nThis email is automatically sent from Google Apps."
  GmailApp.sendEmail(
    address , 
    mailsubject,
    bodymsg1 + "\n\n" + bodymsg2,
    {attachments: [file]}
  );

}

const language = {
    "クメール語":"km",
    "キニヤルワンダ語":"rw",
    "ノルウェー語":"no",
    "アラビア文字":"ar",
    "スンダ語":"su",
    "ミャンマー語（ビルマ語）":"my",
    "リトアニア語":"lt",
    "エストニア語":"et",
    "ベラルーシ語":"be",
    "ブルガリア語":"bg",
    "アフリカーンス語":"af",
    "マルタ語":"mt",
    "タタール語":"tt",
    "フランス語":"fr",
    "マレー語":"ms",
    "ポルトガル語（ポルトガル、ブラジル）":"pt",
    "イディッシュ語":"yi",
    "アイルランド語":"ga",
    "モンゴル語":"mn",
    "セブ語":"ceb",
    "サモア語":"sm",
    "カンナダ語":"kn",
    "ボスニア語":"bs",
    "ラテン語":"la",
    "タミル語":"ta",
    "マラヤーラム文字":"ml",
    "オリヤ語":"or",
    "アムハラ語":"am",
    "マケドニア語":"mk",
    "スペイン語":"es",
    "クロアチア語":"hr",
    "インドネシア語":"id",
    "パンジャブ語":"pa",
    "ネパール語":"ne",
    "ショナ語":"sn",
    "エスペラント語":"eo",
    "パシュト語":"ps",
    "アイスランド語":"is",
    "モン語":"hmn",
    "マラガシ語":"mg",
    "タイ語":"th",
    "ヨルバ語":"yo",
    "フィンランド語":"fi",
    "チェコ語":"cs",
    "アルメニア語":"hy",
    "マオリ語":"mi",
    "フリジア語":"fy",
    "ヒンディー語":"hi",
    "ウクライナ語":"uk",
    "トルコ語":"tr",
    "ロシア語":"ru",
    "ベトナム語":"vi",
    "シンハラ語":"si",
    "テルグ語":"te",
    "ポーランド語":"pl",
    "ペルシャ語":"fa",
    "セソト語":"st",
    "タガログ語（フィリピン語）":"tl",
    "ウルドゥー語":"ur",
    "ウイグル語":"ug",
    "アゼルバイジャン語":"az",
    "セルビア語":"sr",
    "イボ語":"ig",
    "ルーマニア語":"ro",
    "スウェーデン語":"sv",
    "ヘブライ語":"he",
    "ラトビア語":"lv",
    "カザフ語":"kk",
    "スワヒリ語":"sw",
    "日本語":"ja",
    "デンマーク語":"da",
    "コルシカ語":"co",
    "ラオ語":"lo",
    "タジク語":"tg",
    "コーサ語":"xh",
    "韓国語":"ko",
    "オランダ語":"nl",
    "ハワイ語":"haw",
    "ルクセンブルク語":"lb",
    "スコットランド ゲール語":"gd",
    "ズールー語":"zu",
    "ガリシア語":"gl",
    "ベンガル文字":"bn",
    "シンド語":"sd",
    "キルギス語":"ky",
    "ハウサ語":"ha",
    "グジャラト語":"gu",
    "英語":"en",
    "クルド語":"ku",
    "ドイツ語":"de",
    "中国語（繁体）":"zh-TW",
    "バスク語":"eu",
    "クレオール語（ハイチ）":"ht",
    "ソマリ語":"so",
    "スロベニア語":"sl",
    "トルクメン語":"tk",
    "グルジア語":"ka",
    "ジャワ語":"jv",
    "カタロニア語":"ca",
    "イタリア語":"it",
    "ウェールズ語":"cy",
    "スロバキア語":"sk",
    "ウズベク語":"uz",
    "アルバニア語":"sq",
    "中国語（簡体）":"zh-CN",
    "ハンガリー語":"hu",
    "ギリシャ語":"el",
    "マラーティー語":"mr",
    "ニャンジャ語（チェワ語）":"ny"
};

