var url_loginSheet = "https://docs.google.com/spreadsheets/d/1L0kdxaoabT0MCyPXn0tTmKUxu1WvlyKoXVTOcjuL7qM/edit?usp=sharing";

var url_research_table = "https://docs.google.com/spreadsheets/d/1vE8JJ2SftmPgEydzXYnberjQEd7VUFyJ2hDj3hbzKeg/edit?usp=sharing";

var folder_ID = "1lwfCB-2AAYglVLWOuvRtsn6ACaXQxU1G";


function doGet(request) {

  if(request.parameters.v){
    return HtmlService.createTemplateFromFile('PassReset').evaluate();
  
  }else if(request.parameters.n){
    return HtmlService.createTemplateFromFile('Activation').evaluate();

  }else if(request.parameters.m){
    return HtmlService.createTemplateFromFile('Copy of MainPage').evaluate();

  }else if(request.parameters.a){
    return HtmlService.createTemplateFromFile('AfterLoginPage').evaluate();

  }else if(request.parameters.r){
    return HtmlService.createTemplateFromFile('ResearchView').evaluate();

  }else if(request.parameters.d){
    return HtmlService.createTemplateFromFile('FrontPageAdmin').evaluate();

  }else if(request.parameters.e){
    return HtmlService.createTemplateFromFile('MainPageEdit').evaluate();  
  
  }else{
  return HtmlService.createTemplateFromFile('FrontPage').evaluate();
}

}

function include(File) {
    return HtmlService.createHtmlOutputFromFile(File).getContent();
};

function getUrl() {
    return ScriptApp.getService().getUrl();
}

function validateLogin(login_data){

  var result = {};
  result.email = "";
  result.password = "";
  result.activation = "";
  result.slmc_nu = "";

  var ss = SpreadsheetApp.openByUrl(url_loginSheet);
  var ws = ss.getSheetByName('LoginData');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:A").getValues();

    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==login_data.email){
        Logger.log(i)
        if(ws.getRange("D"+(i+1).toString()).getValue()==login_data.email && ws.getRange("E"+(i+1).toString()).getValue()==login_data.password){
        var resulted = ws.getRange("D"+(i+1).toString()+":J"+(i+1).toString()).getValues();
        result.email = resulted[0][0];
        result.password = resulted[0][1];
        result.activation = resulted[0][4];
        result.slmc_nu = resulted[0][6]
        }
      }
    }
  Logger.log('results_login_data : ', result);
  return result
}

function validate_forget_password(fp_data){

  // var fp_data= {};
  // fp_data.email = 'user2@gmail.com'
  var result = {};
  result.email = "";
  result.password = "";
  result.resetLink = "";

  var ss = SpreadsheetApp.openByUrl(url_loginSheet);
  var ws = ss.getSheetByName('LoginData');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:A").getValues();

    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==fp_data.email){
        Logger.log('dataIndex[i][0] : %s',dataIndex[i][0])
        if(ws.getRange("D"+(i+1).toString()).getValue()==fp_data.email){
        var resulted = ws.getRange("D"+(i+1).toString()+":F"+(i+1).toString()).getValues();
         Logger.log('resulted : %s',resulted)
        result.email = resulted[0][0];
        result.password = resulted[0][1];
        result.resetLink = resulted[0][2];
        }
      }
    }
  Logger.log('results_fp : %s', result.email);
  return result

}


function reset_password(reset_data){

  // var reset_data='huhjn67'

  var result = {};
  result.email = "";

  var ss = SpreadsheetApp.openByUrl(url_loginSheet);
  var ws = ss.getSheetByName('LoginData');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("B:B").getValues();

    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==reset_data){
        Logger.log('dataIndex[i][0] : %s',dataIndex[i][0])
        if(ws.getRange("F"+(i+1).toString()).getValue()==reset_data){
        var resulted = ws.getRange("D"+(i+1).toString()+":F"+(i+1).toString()).getValues();
         Logger.log('resulted : %s',resulted)
        result.email = resulted[0][0];
        }
      }
    }
  Logger.log('results : %s', result.email);
  return result

}

function activateUser(activationKey){


  var result = {};
  result.email = "";
  result.status = "";
  result.activationValue = "";

  var ss = SpreadsheetApp.openByUrl(url_loginSheet);
  var ws = ss.getSheetByName('LoginData');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("C:C").getValues();

    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==activationKey){
        Logger.log('dataIndex[i][0] : %s',dataIndex[i][0])
        if(ws.getRange("H"+(i+1).toString()).getValue()==activationKey){
        var resulted = ws.getRange("D"+(i+1).toString()+":H"+(i+1).toString()).getValues();
        ws.getRange("H"+(i+1).toString()).setValue("Activated");
         Logger.log('resulted : %s',resulted)
        result.email = resulted[0][0];
        result.activationValue = resulted[0][4];
        result.status = 'Success';
        }
      }
    }
  Logger.log('results : %s', result.status);
  return result

}

function reset_password_change_pass(change_pass){

  // var change_pass = {};
  // change_pass.reset_email = "user2@gmail.com";
  // change_pass.reset_password = "passg";
  // change_pass.reset_password_c = "passg";

  var result = {};
  result.status = "";

  var ss = SpreadsheetApp.openByUrl(url_loginSheet);
  var ws = ss.getSheetByName('LoginData');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:A").getValues();

    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==change_pass.reset_email){
        Logger.log('dataIndex[i][0] : %s',dataIndex[i][0])
        if(ws.getRange("D"+(i+1).toString()).getValue()==change_pass.reset_email && ws.getRange("F"+(i+1).toString()).getValue()!=""){
        ws.getRange("E"+(i+1).toString()).setValue(change_pass.reset_password);
        ws.getRange("F"+(i+1).toString()).setValue("");

        result.status = "Success";
        }
      }
    }
  Logger.log('results : %s', result.status);
  return result

}

function setResetKey(reset_data){

  // var reset_data = {};
  // reset_data.email = "user2@gmail.com";
  var keyData = {};
  keyData.status = '';
  // keyData.key = "huhkkhh";
  keyData.key = Math.random().toString(36).replace(/[^a-z]+/g, '').substr(0, 10);

  var ss = SpreadsheetApp.openByUrl(url_loginSheet);
  var ws = ss.getSheetByName('LoginData');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:A").getValues();

    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==reset_data.email){
        Logger.log('dataIndex[i][0] : %s',dataIndex[i][0])
        if(ws.getRange("D"+(i+1).toString()).getValue()==reset_data.email){
        ws.getRange("F"+(i+1).toString()).setValue(keyData.key);

        keyData.status = "Success";
        }
      }
    }
  Logger.log('keyData : %s', keyData.status);
  return keyData


}

function sendResetEmail(emailData){

// var emailData ={}
// emailData.email = "user1@gmail.com";
// emailData.link = "user1@edefefegmail.com";

  MailApp.sendEmail(
    emailData.email,
    "Login System password reset link",
    "Dear Sir/Madam," + "\n\n" +
    "Link :  "+emailData.link+ "\n" +
    "Thank you."+ "\n" +
    "Best regards,"+ "\n"+
    "Login Automation System.",
  {name: "GMOA in Collaboation with SHRI",
  }
  );
return "Success"
}

function sendNewUserActivationEmail(emailData){


  MailApp.sendEmail(
    emailData.email,
    "Login System New User Activation link",
    "Dear Sir/Madam," + "\n\n" +
    "Please click on or copy paste below link to activate your account," + "\n\n" +
    "Link :  "+emailData.link+ "\n" +
    "Thank you."+ "\n" +
    "Best regards,"+ "\n"+
    "ET&R Login Automation System.",
  {name: "ET&R Login System",
  }
  );
return "Success"
}

function randomKeyGen(){
  var key = Math.random().toString(36).replace(/[^a-z]+/g, '').substr(0, 20);
  Logger.log("key : %s", key)
}

function regNewUser(newUser_data){

  // var newUser_data = {};
  // newUser_data.email = "user5@gmail.com";
  // newUser_data.password = "dcdcdcdc";
  
  var newUser_reg = {};
  newUser_reg.status = '';
  newUser_reg.activationKey = newUser_data.activationKey;
  newUser_reg.email = newUser_data.email;
  newUser_reg.slmc_nu = newUser_data.slmc_nu;

  var ss = SpreadsheetApp.openByUrl(url_loginSheet);
  var ws = ss.getSheetByName('LoginData');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:A").getValues();

    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==newUser_data.email){
        Logger.log('dataIndex[i][0] : %s',dataIndex[i][0])
        if(ws.getRange("D"+(i+1).toString()).getValue()==newUser_data.email){
          newUser_reg.status ="Already Registered"
          break
        // ws.getRange("F"+(i+1).toString()).setValue(keyData.key);

        // keyData.status = "Success";
        }
      }
    }

    if(newUser_reg.status == ''){
      var lastRowIndex = ws.getLastRow();
      ws.getRange('D'+(lastRowIndex+1).toString()).setValue(newUser_data.email);
      ws.getRange('E'+(lastRowIndex+1).toString()).setValue(newUser_data.password);
      ws.getRange('H'+(lastRowIndex+1).toString()).setValue(newUser_data.activationKey);
      ws.getRange('I'+(lastRowIndex+1).toString()).setValue(newUser_data.initials_nu);
      ws.getRange('J'+(lastRowIndex+1).toString()).setValue(newUser_data.slmc_nu);
      ws.getRange('K'+(lastRowIndex+1).toString()).setValue(newUser_data.initials_nu);
      ws.getRange('L'+(lastRowIndex+1).toString()).setValue(newUser_data.last_name_nu);
      ws.getRange('M'+(lastRowIndex+1).toString()).setValue(newUser_data.full_name_nu);
      ws.getRange('N'+(lastRowIndex+1).toString()).setValue(newUser_data.sex_nu);
      ws.getRange('O'+(lastRowIndex+1).toString()).setValue(newUser_data.mobile_nu);
      ws.getRange('P'+(lastRowIndex+1).toString()).setValue(newUser_data.address_nu);


      // ws.appendRow(
      //   [,
      //   ,
      //   ,
      //   newUser_data.email,
      //   newUser_data.password,
      //   ,
      //   ,
      //   newUser_data.activationKey,
      //   ,
      //   newUser_reg.slmc_nu]
      // )

      newUser_reg.status = 'Success';

      

    }
  Logger.log('newUser_reg : %s', newUser_reg.status);
  return newUser_reg


}

function encodeBase64_url(value){
var encoded = Utilities.base64EncodeWebSafe("HHHHHHHH"+value.slmc_nu);
var url = ScriptApp.getService().getUrl();
var encoded_data = {};
encoded_data.url = url;
encoded_data.encoded = encoded;
Logger.log('encoded');
Logger.log(encoded);
// decodedVal = decodeBase64(encoded_data)
return encoded_data
}


function encodeBase64_researchID(value){
var encoded = Utilities.base64EncodeWebSafe("HHHHHHHH"+value);
var url = ScriptApp.getService().getUrl();
var encoded_data = {};
encoded_data.url = url;
encoded_data.encoded = encoded;
Logger.log('encoded');
Logger.log(encoded);
// decodedVal = decodeBase64(encoded_data)
return encoded_data
}

function decodeBase64(encoded_data){
  var decoded = Utilities.base64DecodeWebSafe(encoded_data);
  var decodedValStep1 = Utilities.newBlob(decoded).getDataAsString();
  // decodedVal = decodedValStep1.substring(8, 13);
  decodedVal = decodedValStep1.substring(8, decodedValStep1.length);
  Logger.log('decodedVal');
  Logger.log(decodedVal);
  return decodedVal;
}



function decodeBase64_for_view(encoded_data){
  var decoded = Utilities.base64DecodeWebSafe(encoded_data);
  var decodedValStep1 = Utilities.newBlob(decoded).getDataAsString();
  decodedVal = decodedValStep1.substring(8, decodedValStep1.length);
  Logger.log('decodedVal');
  Logger.log(decodedVal);
  return decodedVal;
}




function retrieveResearchData(slmc_raw="33333"){

  // var slmc="33333";

  var slmc = slmc_raw.toString();
  var ss = SpreadsheetApp.openByUrl(url_research_table);
  var ws = ss.getSheetByName('Research Table');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:G").getValues();

    var dataRetrieved = [];

    var i = 0;
    
    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][2]==slmc || dataIndex[i][3]==slmc|| dataIndex[i][4]==slmc||dataIndex[i][5]==slmc||dataIndex[i][6]==slmc){
        // dataRetrieved.push(dataIndex[i]);
        if(ws.getRange('F'+(i+1).toString()).getValue()==slmc ||
          ws.getRange('G'+(i+1).toString()).getValue()==slmc ||
          ws.getRange('H'+(i+1).toString()).getValue()==slmc ||
          ws.getRange('I'+(i+1).toString()).getValue()==slmc ||
          ws.getRange('J'+(i+1).toString()).getValue()==slmc){

            dataRetrieved.push(ws.getRange('A'+(i+1).toString()+":Q"+(i+1).toString()).getValues());
            //  dataRetrieved.push(ws.getRange('D'+(i+1).toString()).getValue());

        }
     }
    }
    Logger.log("array : %s",dataRetrieved)
    return dataRetrieved

}


function retrieveResearchsForRID(rid="ResearchID1"){

  var ss = SpreadsheetApp.openByUrl(url_research_table);
  var ws = ss.getSheetByName('Research Table');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:G").getValues();

    var dataRetrievedForRID = [];

    var i = 0;
    
    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==rid){
        // dataRetrieved.push(dataIndex[i]);
        if(ws.getRange('D'+(i+1).toString()).getValue()==rid){

            dataRetrievedForRID.push(ws.getRange('A'+(i+1).toString()+":AA"+(i+1).toString()).getValues());
            //  dataRetrieved.push(ws.getRange('D'+(i+1).toString()).getValue());

        }
     }
    }
    Logger.log("array : %s",dataRetrievedForRID)
    return dataRetrievedForRID

}


function retrieveLoginDataforSLMC(slmc="33333"){

  var ss = SpreadsheetApp.openByUrl(url_loginSheet);
  var ws = ss.getSheetByName('LoginData');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:D").getValues();

    var dataRetrievedForSLMC = [];

    var i = 0;
    
    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][3]==slmc){
        if(ws.getRange('J'+(i+1).toString()).getValue()==slmc){

            dataRetrievedForSLMC.push(ws.getRange('A'+(i+1).toString()+":Q"+(i+1).toString()).getValues());
            //  dataRetrieved.push(ws.getRange('D'+(i+1).toString()).getValue());

        }
     }
    }
    Logger.log("array : %s",dataRetrievedForSLMC)
    return dataRetrievedForSLMC

}



function retrieveUserData(){

  var ss = SpreadsheetApp.openByUrl(url_loginSheet);
  var wsUsers = ss.getSheetByName('Users');
  var dataUsers = wsUsers.getRange("A:C").getValues();

  Logger.log('dataUsers : %s',dataUsers[0])

  return dataUsers

}

function saveToSheet(saveData){

  // var saveData = {};
  // saveData.radio_pi_type = 'Yes'
  // saveData.slmc_pi = 'document.getElementById("slmc_pi").value;'
  // saveData.re_ID = 'ResearchID41'
  // saveData.slmc_pi = 'document.getElementById("slmc_pi").value;'
  // saveData.co_1 = 'document.getElementById("co_1").value;'
  // saveData.co_2 = 'document.getElementById("co_2").value;'
  // saveData.co_3 = 'document.getElementById("co_3").value;'
  // saveData.co_4 = 'document.getElementById("co_4").value;'
  // saveData.ra_type = 'document.getElementById("ra_type").value;'
  // saveData.ra_sub_type = 'document.getElementById("ra_sub_type").value;'
  // saveData.r_topic = 'document.getElementById("r_topic").value;'
  // saveData.file_raa = 'document.getElementById("file_raa").value;'


  var ss = SpreadsheetApp.openByUrl(url_research_table);
  var ws = ss.getSheetByName('Research Table');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:A").getValues();

    var saveStatusBack = {};
    saveStatusBack.status = 'Not found';

    var i = 0;
    
    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==saveData.re_ID){
        if(ws.getRange('D'+(i+1).toString()).getValue()==saveData.re_ID){
            saveStatusBack.status = 'Found';
            saveStatusBack.rowIndex = i;
        }
      }
    }

    Logger.log('saveStatusBack : %s', saveStatusBack)
    Logger.log('saveStatusBack.status : %s', saveStatusBack.status)
   
   if(saveStatusBack.status == 'Not found'){

    var date = Utilities.formatDate(new Date(), 'GMT+5:30', 'MMMM dd, yyyy HH:mm:ss Z')
    var today = new Date().toISOString().slice(0, 10)

      var lastRowIndex = ws.getLastRow();

      var date = Utilities.formatDate(new Date(), 'GMT+5:30', 'MMMM dd, yyyy HH:mm:ss Z')
      var today = new Date().toISOString().slice(0, 10)
      ws.getRange('B'+(lastRowIndex+1).toString()).setValue('1');
      
      ws.getRange('D'+(lastRowIndex+1).toString()).setValue(saveData.re_ID);
      ws.getRange('E'+(lastRowIndex+1).toString()).setValue(saveData.radio_pi_type);
      ws.getRange('F'+(lastRowIndex+1).toString()).setValue(saveData.slmc_pi);
      ws.getRange('G'+(lastRowIndex+1).toString()).setValue(saveData.co_1);
      ws.getRange('H'+(lastRowIndex+1).toString()).setValue(saveData.file_raa_ci_1);
      ws.getRange('I'+(lastRowIndex+1).toString()).setValue(saveData.co_2);
      ws.getRange('J'+(lastRowIndex+1).toString()).setValue(saveData.file_raa_ci_2);
      ws.getRange('K'+(lastRowIndex+1).toString()).setValue(saveData.co_3);
      ws.getRange('L'+(lastRowIndex+1).toString()).setValue(saveData.file_raa_ci_3);
      ws.getRange('M'+(lastRowIndex+1).toString()).setValue(saveData.co_4);
      ws.getRange('N'+(lastRowIndex+1).toString()).setValue(saveData.file_raa_ci_4);
      ws.getRange('O'+(lastRowIndex+1).toString()).setValue(saveData.ra_type);
      ws.getRange('P'+(lastRowIndex+1).toString()).setValue(saveData.ra_sub_type);
      ws.getRange('Q'+(lastRowIndex+1).toString()).setValue(saveData.r_topic);
      ws.getRange('R'+(lastRowIndex+1).toString()).setValue(saveData.file_raa1);
      ws.getRange('S'+(lastRowIndex+1).toString()).setValue(saveData.file_raa2);
      ws.getRange('T'+(lastRowIndex+1).toString()).setValue(saveData.file_raa3);
      ws.getRange('U'+(lastRowIndex+1).toString()).setValue(saveData.file_raa4);
      ws.getRange('V'+(lastRowIndex+1).toString()).setValue(saveData.file_raa5);
      ws.getRange('Z'+(lastRowIndex+1).toString()).setValue(date);
      ws.getRange('AA'+(lastRowIndex+1).toString()).setValue(today);
   

    }

    if(saveStatusBack.status == 'Found'){

    var date = Utilities.formatDate(new Date(), 'GMT+5:30', 'MMMM dd, yyyy HH:mm:ss Z')
    var today = new Date().toISOString().slice(0, 10)


      ws.getRange('B'+(saveStatusBack.rowIndex+1).toString()).setValue('1');
      
      ws.getRange('D'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.re_ID);
      ws.getRange('E'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.radio_pi_type);
      ws.getRange('F'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.slmc_pi);
      ws.getRange('G'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_1);
      ws.getRange('H'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_2);
      ws.getRange('I'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_3);
      ws.getRange('J'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_4);
      ws.getRange('K'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.ra_type);
      ws.getRange('L'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.ra_sub_type);
      ws.getRange('M'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.r_topic);
      ws.getRange('N'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa);
      ws.getRange('O'+(saveStatusBack.rowIndex+1).toString()).setValue(date);
      ws.getRange('Q'+(saveStatusBack.rowIndex+1).toString()).setValue(today);
   

    }

    return saveStatusBack

}


function saveToSheet_edit(saveData){

  // var saveData = {};
  // saveData.radio_pi_type = 'Yes'
  // saveData.slmc_pi = 'document.getElementById("slmc_pi").value;'
  // saveData.re_ID = 'ResearchID41'
  // saveData.slmc_pi = 'document.getElementById("slmc_pi").value;'
  // saveData.co_1 = 'document.getElementById("co_1").value;'
  // saveData.co_2 = 'document.getElementById("co_2").value;'
  // saveData.co_3 = 'document.getElementById("co_3").value;'
  // saveData.co_4 = 'document.getElementById("co_4").value;'
  // saveData.ra_type = 'document.getElementById("ra_type").value;'
  // saveData.ra_sub_type = 'document.getElementById("ra_sub_type").value;'
  // saveData.r_topic = 'document.getElementById("r_topic").value;'
  // saveData.file_raa = 'document.getElementById("file_raa").value;'


  var ss = SpreadsheetApp.openByUrl(url_research_table);
  var ws = ss.getSheetByName('Research Table');
  var wsIndex = ss.getSheetByName('Index');
  var dataIndex = wsIndex.getRange("A:A").getValues();

    var saveStatusBack = {};
    saveStatusBack.status = 'Not found';

    var i = 0;
    
    for (var i = 0; i < dataIndex.length; i++) {
      if(dataIndex[i][0]==saveData.re_ID){
        if(ws.getRange('D'+(i+1).toString()).getValue()==saveData.re_ID){
            saveStatusBack.status = 'Found';
            saveStatusBack.rowIndex = i;
        }
      }
    }

    Logger.log('saveStatusBack : %s', saveStatusBack)
    Logger.log('saveStatusBack.status : %s', saveStatusBack.status)
   
  //  if(saveStatusBack.status == 'Not found'){

  //   var date = Utilities.formatDate(new Date(), 'GMT+5:30', 'MMMM dd, yyyy HH:mm:ss Z')
  //   var today = new Date().toISOString().slice(0, 10)

  //     var lastRowIndex = ws.getLastRow();

  //     var date = Utilities.formatDate(new Date(), 'GMT+5:30', 'MMMM dd, yyyy HH:mm:ss Z')
  //     var today = new Date().toISOString().slice(0, 10)
  //     ws.getRange('B'+(lastRowIndex+1).toString()).setValue('1');
      
  //     ws.getRange('D'+(lastRowIndex+1).toString()).setValue(saveData.re_ID);
  //     ws.getRange('E'+(lastRowIndex+1).toString()).setValue(saveData.radio_pi_type);
  //     ws.getRange('F'+(lastRowIndex+1).toString()).setValue(saveData.slmc_pi);
  //     ws.getRange('G'+(lastRowIndex+1).toString()).setValue(saveData.co_1);
  //     ws.getRange('H'+(lastRowIndex+1).toString()).setValue(saveData.file_raa_ci_1);
  //     ws.getRange('I'+(lastRowIndex+1).toString()).setValue(saveData.co_2);
  //     ws.getRange('J'+(lastRowIndex+1).toString()).setValue(saveData.file_raa_ci_2);
  //     ws.getRange('K'+(lastRowIndex+1).toString()).setValue(saveData.co_3);
  //     ws.getRange('L'+(lastRowIndex+1).toString()).setValue(saveData.file_raa_ci_3);
  //     ws.getRange('M'+(lastRowIndex+1).toString()).setValue(saveData.co_4);
  //     ws.getRange('N'+(lastRowIndex+1).toString()).setValue(saveData.file_raa_ci_4);
  //     ws.getRange('O'+(lastRowIndex+1).toString()).setValue(saveData.ra_type);
  //     ws.getRange('P'+(lastRowIndex+1).toString()).setValue(saveData.ra_sub_type);
  //     ws.getRange('Q'+(lastRowIndex+1).toString()).setValue(saveData.r_topic);
  //     ws.getRange('R'+(lastRowIndex+1).toString()).setValue(saveData.file_raa1);
  //     ws.getRange('S'+(lastRowIndex+1).toString()).setValue(saveData.file_raa2);
  //     ws.getRange('T'+(lastRowIndex+1).toString()).setValue(saveData.file_raa3);
  //     ws.getRange('U'+(lastRowIndex+1).toString()).setValue(saveData.file_raa4);
  //     ws.getRange('V'+(lastRowIndex+1).toString()).setValue(saveData.file_raa5);
  //     ws.getRange('Z'+(lastRowIndex+1).toString()).setValue(date);
  //     ws.getRange('AA'+(lastRowIndex+1).toString()).setValue(today);
   

  //   }

    if(saveStatusBack.status == 'Found'){

    var date = Utilities.formatDate(new Date(), 'GMT+5:30', 'MMMM dd, yyyy HH:mm:ss Z')
    var today = new Date().toISOString().slice(0, 10)


      ws.getRange('B'+(saveStatusBack.rowIndex+1).toString()).setValue('5');
      
      // ws.getRange('D'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.re_ID);
      // ws.getRange('E'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.radio_pi_type);
      // ws.getRange('F'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.slmc_pi);
      // ws.getRange('G'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_1);
      // ws.getRange('H'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_2);
      // ws.getRange('I'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_3);
      // ws.getRange('J'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_4);
      // ws.getRange('K'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.ra_type);
      // ws.getRange('L'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.ra_sub_type);
      // ws.getRange('M'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.r_topic);
      // ws.getRange('N'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa);
      // ws.getRange('O'+(saveStatusBack.rowIndex+1).toString()).setValue(date);
      // ws.getRange('Q'+(saveStatusBack.rowIndex+1).toString()).setValue(today);

      if(saveData.re_ID!=""){ws.getRange('D'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.re_ID);}
      if(saveData.radio_pi_type!=""){ ws.getRange('E'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.radio_pi_type);}
      if(saveData.slmc_pi!=""){ws.getRange('F'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.slmc_pi);}
      if(saveData.co_1!=""){ws.getRange('G'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_1);}
      if(saveData.file_raa_ci_1!=""){ws.getRange('H'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa_ci_1);}
      if(saveData.co_2!=""){ws.getRange('I'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_2);}
      if(saveData.file_raa_ci_2!=""){ws.getRange('J'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa_ci_2);}
      if(saveData.co_3!=""){ws.getRange('K'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_3);}
      if(saveData.file_raa_ci_3!=""){ws.getRange('L'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa_ci_3);}
      if(saveData.co_4!=""){ws.getRange('M'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.co_4);}
      if(saveData.file_raa_ci_4!=""){ws.getRange('N'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa_ci_4);}
      if(saveData.ra_type!=""){ws.getRange('O'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.ra_type);}
      if(saveData.ra_sub_type!=""){ws.getRange('P'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.ra_sub_type);}
      if(saveData.r_topic!=""){ ws.getRange('Q'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.r_topic);}
      if(saveData.file_raa1!=""){ws.getRange('R'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa1);}
      if(saveData.file_raa2!=""){ws.getRange('S'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa2);}
      if(saveData.file_raa3!=""){ws.getRange('T'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa3);}
      if(saveData.file_raa4!=""){ws.getRange('U'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa4);}
      if(saveData.file_raa5!=""){ws.getRange('V'+(saveStatusBack.rowIndex+1).toString()).setValue(saveData.file_raa5);}
      ws.getRange('Z'+(saveStatusBack.rowIndex+1).toString()).setValue(date);
      ws.getRange('AA'+(saveStatusBack.rowIndex+1).toString()).setValue(today);
   

    }

    return saveStatusBack

}



function upload_and_save_to_folder(obj) {
    var blob = Utilities.newBlob(Utilities.base64Decode(obj.data), obj.mimeType, obj.fileName);
    var url = '';
    var dest = folder_ID;
    var destination = DriveApp.getFolderById(dest);
    var getId = destination.createFile(blob).getId();
    var files = DriveApp.getFileById(getId)
    var url = files.getUrl();
    return url
}























