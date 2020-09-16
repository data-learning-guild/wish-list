//ドキュメントID
var id = PropertiesService.getScriptProperties().getProperty('SpreadSheet_id');

// IDからスプレッドシートを取得
var spreadsheet = SpreadsheetApp.openById(id);

// シートを取得
var sheet_name = PropertiesService.getScriptProperties().getProperty('Sheet_Name')
var sheet = spreadsheet.getSheetByName(sheet_name);
var token = PropertiesService.getScriptProperties().getProperty('OAuth_token');
var slackApp = SlackApp.create(token);

function doPost(e) {
  //GASのプロパティストアに登録したVerification Token
  var verified_token = PropertiesService.getScriptProperties().getProperty('verified_token');
  // wish-リストチャンネルのID
  var postChannel = PropertiesService.getScriptProperties().getProperty('Channel_id')
  var verificationToken = e.parameter.token || JSON.parse(e.parameter.payload).token || null;
  if (verificationToken !== verified_token) { // AppのVerification
    console.log(e);
    return ContentService.createTextOutput();
  }
  
  if (e.parameter.command === '/wish_help') {
    var rtnjson = {"response_type": "ephemeral","text": 'コマンドリスト\n`/wish` お願い事の追加\n`/wish_list` お願い事リストの表示'};
    return ContentService.createTextOutput(JSON.stringify(rtnjson)).setMimeType(ContentService.MimeType.JSON);
  } else if (e.parameter.command === '/wish_list') {
    //ウィッシュリストの一覧を取得する
    text = getWishList()
    var rtnjson = {"response_type": "ephemeral","text": '叶えてほしい願いはこれだ.....\n'+'```'+text+'\nスプレッドシート:https://docs.google.com/spreadsheets/d/1B6DDNtvUFOgcq3FNrkZ5ScJycRYK-9E990iuSGJZZaI/edit#gid=0'+'```'};
    return ContentService.createTextOutput(JSON.stringify(rtnjson)).setMimeType(ContentService.MimeType.JSON);
    
  } else if (e.parameter.command === '/wish'||e.parameter.command === '/wish_complete') {
    var createdDialog = createDialog(e);
    var options = {
    'method' : 'POST',
    'payload' : createdDialog,
    };
    var slackUrl = "https://slack.com/api/dialog.open";
    var response = UrlFetchApp.fetch(slackUrl, options);
    return ContentService.createTextOutput();
  } 
    var p = JSON.parse(e.parameter.payload)
    var s = p.submission;

  if (p.callback_id == 'irai_dialog') {
    var remark = s.remark || "なし"
    appendWishItem(p, remark)
    slackApp.postMessage(postChannel,
      "今回お願いした人by <@" + p.user.id + ">\n【タイトル】：" + s.title
      + "\n【お願いの種類】：" + s.type + "\n【説明】：" + s.description
      + "\n【期限】：" + s.date + "\n【備考】：" + remark);
   } else if (p.callback_id == 'complete_dialog'){
     moveItem(s.type.substring(0, 2), s.handler)
     slackApp.postMessage(postChannel,
                        " <@" + p.user.id + ">の願いが叶えられた\n" + s.type);
   }

    return ContentService.createTextOutput();//←これが無いとslackのダイアログが閉じない。	
}
function createDialog(e) {
  var trigger_id = e.parameter.trigger_id
  var token = PropertiesService.getScriptProperties().getProperty('OAuth_token')

  if (e.parameter.command === '/wish') {
    //期限を14日後に設定
    var date = new Date()
    date.setDate(date.getDate() + 14)
    var deadline = Utilities.formatDate(date, 'JST', 'yyyy/MM/dd')

    var dialog = {
      token: token, // OAuth_token
      trigger_id: trigger_id,
      dialog: JSON.stringify({
        callback_id: 'irai_dialog',
        title: 'お願いフォーム',
        submit_label: 'お願いする',
        elements: [
          {
            type: 'text',
            label: 'タイトル',
            name: 'title',
          },
          {
            type: 'select',
            label: 'お願いの種類',
            name: 'type',
            options: [
              {
                label: 'アイディア募集',
                value: 'アイディア募集',
              },
              {
                label: '助けて欲しい！',
                value: '助けて欲しい！',
              },
              {
                label: '教えて欲しい！',
                value: '教えて欲しい！',
              },
              {
                label: '紹介して欲しい！',
                value: '紹介して欲しい！',
              },
              {
                label: '手伝って欲しい！',
                value: '手伝って欲しい！',
              },
              {
                label: '相談に乗ってほしい！',
                value: '相談に乗ってほしい！',
              },
              {
                label: 'データ分析',
                value: 'データ分析',
              },
              {
                label: 'アンケート！',
                value: 'アンケート！',
              },
              {
                label: 'その他',
                value: 'その他',
              },
            ],
          },
          {
            type: 'textarea',
            label: '説明',
            name: 'description',
          },
          {
            type: 'text',
            label: '期限',
            name: 'date',
            value: deadline,
            placeholder: 'YYYY/MM/DD',
          },
          {
            type: 'textarea',
            label: '備考欄',
            name: 'remark',
            optional: true, //この指定がないとrequiredになる
            placeholder: '備考欄のみ入力必須ではありません',
          },
        ],
      }),
    }
  } else if (e.parameter.command === '/wish_complete') {
    var LastRow = sheet.getLastRow()
    var wishItems = sheet.getRange(2, 3, LastRow-1, 1).getValues()
    var options = []
    for (i = 0; i < wishItems.length; i++) {
      options.push({
        label: wishItems[i][0],
        value: (i + 2) + " " + wishItems[i][0]
      })
    }
    var dialog = {
      token: token, // OAuth_token
      trigger_id: trigger_id,
      dialog: JSON.stringify({
        callback_id: 'complete_dialog',
        title: '完了フォーム',
        submit_label: 'お願いする',
        elements: [
          {
            type: 'select',
            label: '叶った願いはどれだ...',
            name: 'type',
            options: options,
          },
          {
            type: 'text',
            label: '叶えてくれた人',
            name: 'handler'
          },
        ],
      }),
    }
  }
  return dialog
}

function appendWishItem(p, remark){

  var last_row = sheet.getLastRow();　// 最後の行を取得
  var soNo = 1

  for(var i = last_row; i >= 1; i--) {　
    if(sheet.getRange(i,　1).getValue() != '') {　
      soNo += sheet.getRange(i, 1).getValue();
      break;
    }
  }
  var s = p.submission;
  var user_name = getUser(p.user.id)

        // 1行を追加
  sheet.appendRow([soNo, user_name, s.title, s.type, s.description, s.date, remark]);
  return

}
function appendWishItem(p, remark){

  var last_row = sheet.getLastRow();　// 最後の行を取得
  var soNo = 1

  for(var i = last_row; i >= 1; i--) {　
    if(sheet.getRange(i,　1).getValue() != '') {　
      soNo += sheet.getRange(i, 1).getValue();
      break;
    }
  }
  var s = p.submission;
  var user_name = getUser(p.user.id)

        // 1行を追加
  sheet.appendRow([soNo, user_name, s.title, s.type, s.description, s.date, remark]);
  return

} 
                           
function getWishList(){
  var messageList = ''
  // レコード数
  var recordNum = sheet.getLastRow();
  //データ取得
  var records = sheet.getRange(1, 1, recordNum, 7).getValues();
  for(i=0;i<recordNum;i++){
    var wishItem = records[i]
    var message = String(wishItem[0])+ '    '+wishItem[2]+ '    '+wishItem[5]+ '\n'
    messageList += message
  }
  
  return messageList
}

function moveItem(num, handler) {
  num = num.replace(/[\s\t\n]/g,"");
  var comp_sheet_name = PropertiesService.getScriptProperties().getProperty('Comp_Sheet')
  comp_sheet = spreadsheet.getSheetByName(comp_sheet_name);
  
  var rowSpec = sheet.getRange("A"+num+":G"+num).getValues();
  //var rowSpec = sheet.getRange("A2:G2").getValues();
  comp_sheet.appendRow(rowSpec[0])
  comp_sheet.getRange("H"+String(comp_sheet.getLastRow())).setValue(handler)
  sheet.deleteRow(Number(num))
}

function getUser(user_id) {
  
  var url = "https://slack.com/api/users.info";
  var payload = {
    "token" : token,
    "user" : user_id
  };
  
  var params = {
    "method" : "get",
    "payload" : payload
  };
  
  // Slackにリクエストする
  var user = JSON.parse(UrlFetchApp.fetch(url, params));
  //display_nameが入ってなかったらreal_nameを入れる
  var user_name = user.user.profile.display_name
  if(user_name === ""){
    user_name = user.user.real_name
  }
  return user_name
  
}