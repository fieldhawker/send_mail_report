function doGet() {
  var html = HtmlService.createTemplateFromFile('input_monthly_report');
  html.data = JSON.stringify( getStaffs() );
  
  var htmlOutput = html.evaluate();
  htmlOutput
    .setTitle('[SEP][二課１G] ゲツジェネ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  return htmlOutput;
}

function doPost(e){

  var html = HtmlService.createTemplateFromFile('result_monthly_report');  
  
  var template = getTemplate();
  var text  = template[1][0];
  var title = template[1][1];
  
  var staffs = getStaffs();
  var staff_select = staffs[e.parameters.staff_select].name;
  var reader = staffs[e.parameters.staff_select].reader;
  var number = staffs[e.parameters.staff_select].number;
  var group_address = staffs[e.parameters.staff_select].group_address + ',' + staffs[e.parameters.staff_select].address;
  var affiliation = staffs[e.parameters.staff_select].affiliation;
  var start_date = e.parameters.start_date;
  var end_date = e.parameters.end_date;
  var q1 = e.parameters.q1;
  var q2 = e.parameters.q2;
  var q3 = e.parameters.q3;
  var q4 = e.parameters.q4;
  var q5 = e.parameters.q5;
  var q6 = e.parameters.q6;
  var q7 = e.parameters.q7;
  var q8 = e.parameters.q8;
  var q9 = e.parameters.q9;
  
  if (typeof q4 === "undefined") {
    q4 = '';
  }
  if (typeof q5 === "undefined") {
    q5 = '';
  }
  if (typeof q6 === "undefined") {
    q6 = '';
  }
  if (typeof q7 === "undefined") {
    q7 = '';
  }
  
  var mail_text  = text
    .replace(/__READER_NAME__/,reader)
    .replace(/__NAME__/g,staff_select)
    .replace(/__NUMBER__/,number)
    .replace(/__AFFILIATION__/,affiliation)
    .replace(/__START_DATE__/,start_date)
    .replace(/__END_DATE__/,end_date)
    .replace(/__Q1__/,q1)
    .replace(/__Q2__/,q2)
    .replace(/__Q3__/,q3)
    .replace(/__Q4__/,q4.toString().replace(/,/g,'\r\n'))
    .replace(/__Q5__/,q5.toString().replace(/,/g,'\r\n'))
    .replace(/__Q6__/,q6.toString().replace(/,/g,'\r\n'))
    .replace(/__Q7__/,q7.toString().replace(/,/g,'\r\n'))
    .replace(/__Q8__/,q8)
    .replace(/__Q9__/,q9);
    
  var mail_title  = title
    .replace(/__AFFILIATION__/,affiliation)
    .replace(/__NAME__/,staff_select);
    
  var array = ['mailto:', group_address, '?subject=', encodeURIComponent(mail_title), '&body=', encodeURIComponent(mail_text)];
  var mailto = array.join('');
  var array = ['mailto:', group_address, '?subject=', mail_title, '&body=', mail_text];
  var mailto2 = array.join('');
  
  Logger.log(mailto);
  
  html.mail_title = mail_title;
  html.mail_text = mail_text;
  html.mail_address = group_address;
  html.mailto = mailto;
  html.mailto2 = mailto2;
  html.staff_select = staff_select;
  html.start_date = start_date;
  html.end_date = end_date;
  html.q1 = q1;
  html.q2 = q2;
  html.q3 = q3;
  html.q4 = q4;
  html.q5 = q5;
  html.q6 = q6;
  html.q7 = q7;
  html.q8 = q8;
  html.q9 = q9;
  
  var htmlOutput = html.evaluate();
  htmlOutput
    .setTitle('[SEP][二課１G] ゲツジェネ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  return htmlOutput;
  
}

function getStaffs() {
  var id = '16hAj1SmKB94xHsB6wDUf7ewYnNLeEqF9L-6OhR3fv-o'; //Webアプリ化するとIDを自動で取れなかったので調べて記述しておきます
  var sheet = SpreadsheetApp.openById(id);
  var data = sheet.getDataRange().getValues();

  var json = [];

  for (var i = 0; i < data.length; i++) {
    json.push({
      "name": data[i][0], 
      "number": data[i][1], 
      "reader": data[i][2], 
      "group_address": data[i][3], 
      "affiliation": data[i][4],            
      "address": data[i][5]
    });
  }
//  Logger.log(json);
  return json;

}

function getTemplate() {
  var id = '1TnKm3klXYUAf6xizMLM0KkhSq5Qmyyh-dd8tzf1lEck'; //Webアプリ化するとIDを自動で取れなかったので調べて記述しておきます
  var sheet = SpreadsheetApp.openById(id);
  var data = sheet.getDataRange().getValues();
  
  return data;

}


function getScriptUrl() {
    var url = ScriptApp.getService().getUrl();
    return url;
}