var UserProperties = PropertiesService.getUserProperties();

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Insert Selection Field", functionName: "menu_insertSelection"},
                     {name: "Settings", functionName: "menu_settings"} ];
  ss.addMenu("Odoo", menuEntries);
  var a1 = ss.getRange("O60");
  if ((a1.getFormula().indexOf('oe_settings') > -1 && UserProperties.getProperty('url')) || a1.getFormula().indexOf('oe_call_signin') > -1){
    a1.setFormula("");
  }
  if (UserProperties.getProperty('oe_call_settings')){
    menu_settings();
    UserProperties.deleteProperty('oe_call_settings');
  }
  if (UserProperties.getProperty('oe_call_signin')){
    menu_settings([["username", "Username"], ["password", "Password"]]);
    UserProperties.deleteProperty('oe_call_signin');
  }
}

function menu_settings(params) {
  if (!params){
    params = [["url", "URL (with http:// or https://)"], ["dbname", "Database Name"], ["username", "Username"], ["password", "Password"]];
  }
  for (var i = 0; i < params.length; i++){
    var input = Browser.inputBox("Server Settings", params[i][1], Browser.Buttons.OK_CANCEL);
    if (input === "cancel"){
      break;
    }
    else{
      UserProperties.setProperty(params[i][0], input);
    }
  }
}
 
function menu_insertSelection() {
  var input = Browser.inputBox('Insert Selection', 'Format: model, field, domain', Browser.Buttons.OK_CANCEL);
  if (input !== "cancel"){
    input = input.split(",");
    var model = input[0].replace(/\s+/g, "");
    var field = input[1].replace(/\s+/g, "");
    var domain = input.slice(2,input.length).join(",");
    var range = SpreadsheetApp.getActiveRange();
    oe_select(range, model, field, domain);
  }
}

function oe_settings(url, dbname, username, password){
  if (url)UserProperties.setProperty('url', url);
  if (dbname)UserProperties.setProperty('dbname', dbname);
  if (username)UserProperties.setProperty('username', username);
  if (password)UserProperties.setProperty('password', password);
  else UserProperties.setProperty('oe_call_signin', true);
}

function oe_browse(model, fields, domain, sort, limit){
  if(typeof model !== "string"){
    throw "model arg expecting string";
  }
  if(typeof fields !== "string"){
    throw "fields arg expecting comma separated field names";
  }
  if (!domain) domain = "[]";
  if(typeof domain !== "string"){
    throw "domain arg expecting string";
  }
  if(sort && typeof sort !== "string"){
    throw "sort arg expecting string";
  }
  if(limit && typeof limit !== "number"){
    throw "limit arg expecting number";
  }
  
  fields = fields.replace(/\s+/g, ",").split(",");
  if(domain) {
    domain = domain.replace(/\'/g, '"');
  }
  domain = JSON.parse(domain);
  
  var records = seach_read(model, fields, domain, sort, limit);
  Logger.log(records);
  return parse_records_for_ss(records, fields);
//  return records;
}

function oe_read_group(model, fields, groupby, domain, orderby, limit){
  if (typeof model !== "string"){
    throw "model arg expecting string";
  }
  if (fields && typeof fields !== "string"){
    throw "fields arg expecting comma separated field names";
  }
  if (!groupby || typeof groupby !== "string"){
    throw "groupby arg required, expecting comma separated field names (at least one)";
  }
  if (!domain) domain = "[]";
  if(typeof domain !== "string"){
    throw "domain arg expecting string";
  }
  if(orderby && typeof orderby !== "string"){
    throw "orderby arg expecting string";
  }
  if(limit && typeof limit !== "number"){
    throw "limit arg expecting number";
  }
  
  fields = fields ? fields.replace(/\s+/g, ",").split(",") : [];
  var fields_tosend = fields.slice();
  var count_index = fields_tosend.indexOf("_count");
  if (count_index !== -1){
    fields_tosend.splice(count_index, 1);
  }
  groupby = groupby ? groupby.replace(/\s+/g, ",").split(",") : "";
  for (var i = 0; i < groupby.length; i++){
    if (fields_tosend.indexOf(groupby[i]) === -1){
      fields_tosend.splice(i, 0, groupby[i]);
    }
  }
  if(domain) {
    domain = domain.replace(/\'/g, '"');
  }
  domain = domain ? (JSON.parse(domain)) : [];

  var kwargs = {
    "context" : {"group_by":groupby},
    "domain" : domain,
    "fields" : fields_tosend,
    "groupby": groupby,
    "limit": limit ? limit : 10,
    "offset": 0,
    "orderby": orderby ? orderby : false,
  }
  var records = call_kw(model, "read_group", [], {}, 0, kwargs);
  if (groupby.length > 0){
    for (var i = 0; i < records.length; i++){
      if (records[i]["__context"] && records[i]["__context"]["group_by"].length > 0){
        kwargs["domain"] = records[i]["__domain"]
        kwargs["context"] = records[i]["__context"]
        kwargs["groupby"] = records[i]["__context"]["group_by"]
        var sub_records = call_kw(model, "read_group", [], {}, 0, kwargs);
        sub_records.forEach(function(item){ 
          for(var j = 0; j < this.groupby_fields.length;j++){
            item[this.groupby_fields[j]] = records[i][this.groupby_fields[j]]
          }
        },{
          "groupby_fields" : groupby.slice(0,groupby.indexOf(kwargs["groupby"][0]))
        });
        records.splice.apply(records, [i,1].concat(sub_records));
        i--;
      }
    }
  }
  var count_index = fields.indexOf('_count');
  if (count_index !== -1){
    fields[count_index] = groupby instanceof Array && groupby.length > 0 ? groupby[groupby.length-1]+"_count" : groupby+"_count";
  }
  return parse_records_for_ss(records, fields, "number");
}

function oe_select(range, model, field, domain){
  if(typeof model !== "string"){
    throw "model arg expecting string";
  }
  if(typeof field !== "string"){
    throw "field arg expecting field name";
  }
  if (!domain) domain = "[]";
  if(typeof domain !== "string"){
    throw "domain arg expecting String";
  }
  var records = oe_read_group(model, field, field, domain);
  var result = [];
  for (var i = 0; i < records.length; i++){
    var value = records[i][0];
    if (value)result.push(value.replace(",", ""));
  }
  result = result.slice(0,10);
  dv = SpreadsheetApp.newDataValidation().requireValueInList(result, true).build();
  range.setDataValidation(dv);
}

function parse_records_for_ss(records, fields, force_type){
  var result = [];
  var types = []; 
  if (fields.length === 0 && records.length > 0){
    fields = Object.keys(records[0]);
    result.push(fields);
  }
  for (var i = 0; i < records.length; i++){
    recordArr = [];
    for (var j = 0; j < fields.length; j++){
      var value = records[i][fields[j]];
      if (typeof value === "number") types[fields[j]] = "number"; 
      if (value instanceof Array && value.length === 2 && typeof value[1] === "string")value = value[1];
      else if(value instanceof Array) value = value.join(','); //TODO: name_get on ids
      else if(!value) {
        value = (force_type === "number") ? 0 : 'Undefined';
      }
      recordArr.push(value);
    }
    result.push(recordArr);
  }
  return result.length > 0 ? result : 'No Result';
}
function seach_read(model, fields, domain, sort, limit){
  if(!(fields instanceof Array)){
    throw "fields arg expecting an Array, not "+typeof fields;
  }
  if (!domain)domain = [];
  if(!(domain instanceof Array)){
    throw "domain arg expecting an Array, not "+typeof domain;
  }
  var session_id = getUserProperty("session_id");
  var context = {};
  var params = {
    "model" : model,
    "fields" : fields,
    "limit": limit ? limit : 80,
    "domain" : domain,
    "sort": sort,
    "context": context,
  }
  var options =
      {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : {
          "id": 1,
          "jsonrpc": "2.0",
          "method": "googlescript",
          "params" : params,
        }
      };
  var json_result = JSON.parse(oe_fetch(getUserProperty('url')+'/web/dataset/search_read', options));
  if (!!json_result.error){
    throw format_openerp_error(json_result.error);
  }
  return json_result.result.records;
}

function call_kw(model, method, args, context, debug, kwargs){
  if (typeof model !== "string"){
    throw "model arg expecting a String, not "+typeof model;
  }
  if (typeof method !== "string"){
    throw "method arg expecting a String, not "+typeof model;
  }
  if(!(args instanceof Array)){
    throw "args arg expecting an Array, not "+typeof args;
  }
  if(!(context instanceof Object)){
    throw "context arg expecting an Object, not "+typeof context;
  }
  if(typeof debug !== "number"){
    throw "debug arg expecting a boolean Number, not "+typeof debug;
  }
  if(!(kwargs instanceof Object)){
    throw "kwargs arg expecting an Object, not "+typeof kwargs;
  }
  var session_id = getUserProperty('session_id');
  var params = {
    "args": args,
    "context": context,
    "kwargs": kwargs,
    "method": method,
    "model": model,
  }
  var options =
      {
        "method" : "post",
        "contentType" : "application/json",
        "payload" : {
          "id": 1,
          "jsonrpc": "2.0",
          "method": "googlescript",
          "params" : params,
        }
      };
  var json_result = JSON.parse(oe_fetch(getUserProperty('url')+'/web/dataset/call_kw', options));
  if (!!json_result.error){
    throw format_openerp_error(json_result.error);
  }
  return json_result.result;
}

function authenticate(){
  Logger.log('Authentication requested!');
  var url = getUserProperty("url");
  var dbname = getUserProperty("dbname");
  var username = getUserProperty("username");
  var password = getUserProperty("password"); 
  if (!url || !dbname || !username || !password){
    throw "At least one connection detail is not set. You can set them Odoo > Settings in the menu bar";
  }
  var params = {
    "db": dbname,
    "login": username,
    "password": password,
  }
  var options ={
    "method" : "post",
    "contentType" : "application/json",
    "payload" : JSON.stringify({
      "id": 1,
      "jsonrpc": "2.0",
      "method": "googlescript",
      "params" : params,
    })
  };
  var response = UrlFetchApp.fetch(url+'/web/session/authenticate', options);
  var json_response = JSON.parse(response);
  if (json_response.result.uid){
    var sid = response.getHeaders()["Set-Cookie"].split(" ")[0];
    var session_id = json_response.result.session_id;
    UserProperties.setProperty("sid", sid);
    UserProperties.setProperty("session_id", session_id)
    return {"sid": sid, "session_id": session_id};
  }
  throw "Authentication Error";
}

function oe_fetch(url, options){
  var sid = getUserProperty("sid");
  var session_id = getUserProperty("session_id");
  if (!sid || !session_id){
    var authentication = authenticate();
    sid = authentication.sid;
    session_id = authentication.session_id;
  }
  if (typeof options.headers === 'undefined')options['headers'] = {'cookie': sid};
  else options.headers['cookie'] = sid;
  options['payload'] = JSON.stringify(options.payload);
  for (var i = 0; i < 2; i++){
    var result = UrlFetchApp.fetch(url, options);
    var json_result = JSON.parse(result);
    if (json_result.error){
      authentication = authenticate();
      options['payload'] = JSON.parse(options.payload);
      options.headers['cookie'] = authentication.sid;
      options['payload'] = JSON.stringify(options.payload);
    }
    else if(json_result.error){
      throw format_openerp_error(json_result.error);
    }
    else{
      return result;
    }
  }
  throw "Unable to fetch data due to session expired exception";
}

function getUserProperty(key) {
  var FailLimit = 100;
  var RetryInterval = 50;
  var UserPropertyValue = "";
  var Retries=0;
  var randomnumber = 0;
  var TryAgain=true;
  while (TryAgain)
  {
    Retries++;
    randomnumber=Math.floor(Math.random()*59);
    Utilities.sleep(randomnumber*RetryInterval);
    Logger.log(randomnumber*RetryInterval);
    try
    {
      TryAgain=false;
      UserPropertyValue = UserProperties.getProperty(key);
    }
    catch(err)
    {
      TryAgain = (Retries<FailLimit);
      if (!TryAgain){
        throw 'Too many attempts to acces script property';
      }
      continue;
  }
  return UserPropertyValue;
  }
}

function format_openerp_error(error){
  var error_type = error.data.type;
  var trace = "";
  if (error_type === "client_exception")trace = error.data.debug;
  else if (error_type === "server_exception")trace= error.data.fault_code;
  else trace = JSON.stringify(error.data);
  return error.message + ": "+error_type+", "+ trace;
}
//***********************************BIZ4AFRICA*******************************************************


function getPoByPurchaseId(tabPurchaseId,tabPo){
  var returnPo = [];
  for(var indexPurchaseId = 0; indexPurchaseId < tabPurchaseId.length; indexPurchaseId++){
      for(var indexPo = 0; indexPo < tabPo.length; indexPo++){
          if(tabPo[indexPo]["id"]===tabPurchaseId[indexPurchaseId]){
              returnPo.push(tabPo[indexPo]);
          }
      }
  }
  return returnPo;
}
function getSpByStockPinckingId(tabStockPickingId,tabSp){
  var returnSp = [];
  for(var indexSpId = 0; indexSpId < tabStockPickingId.length; indexSpId++){
    for(var indexSp = 0; indexSp <tabSp.length; indexSp++){
        if(tabSp[indexSp]["id"]===tabStockPickingId[indexSpId]){
          returnSp.push(tabSp[indexSp]);
        }
    }
  }
  return returnSp;
}
function joinPr_Po_Sp(tabPr,tabPo,tabSp){
  for(var indexPo = 0; indexPo < tabPo.length; indexPo++){
     tabPo[indexPo].Sp = getSpByStockPinckingId(tabPo[indexPo]["picking_ids"],tabSp);
   }
   for(var indexPr = 0; indexPr < tabPr.length; indexPr++){
     tabPr[indexPr].Po = getPoByPurchaseId(tabPr[indexPr]["purchase_ids"],tabPo);
   }
   return tabPr;
}
function renderPo_Sp(line){
    var Po_Sp = [];
    var Pofield = [];
    var Spfield = [];
    if(line["Po"].length===0){
        return Po_Sp;
    }else{
        for(var indexPo = 0; indexPo < line["Po"].length;indexPo++){
            for(var key in line["Po"][indexPo]){
                if(key==="Sp" || key==="picking_ids" || key==="requisition_id" || key==="id")
                  continue;
                var value=line["Po"][indexPo][key];
                if (value instanceof Array && value.length === 2 && typeof value[1] === "string")value = value[1];
                if(value===false)value=" ";
                Pofield.push(value);
                Logger.log(key+"Po:"+value);
            }
            if(line["Po"][indexPo]["Sp"].length===0){
                Po_Sp.push(Pofield);
                Pofield = [];
            }else{
                for(var indexSp = 0; indexSp < line["Po"][indexPo]["Sp"].length; indexSp++){
                    for(var key in line["Po"][indexPo]["Sp"][indexSp]){
                        if(key==="purchase_id" || key==="id" || key==="picking_type_id" || key==="picking_ids")
                            continue;
                        var value = line["Po"][indexPo]["Sp"][indexSp][key];
                        if (value instanceof Array && value.length === 2 && typeof value[1] === "string")value = value[1];
                        if(value===false)value=" ";
                        Spfield.push(value);
                        Logger.log(key+"Sp:"+value);
                    }
                    Po_Sp.push(Array.concat(Pofield,Spfield));
                    Spfield = [];
                }
                Pofield = [];
            }
            
        }
    }
    
    return Po_Sp;
}
function renderPr_Po_Sp(JSONPr,Po_Sp){
    var tabRender = [];
    var Prfield = [];
    if(Po_Sp.length === 0){
      for(var key in JSONPr){
        if(key==="Po" || key==="id" || key==="purchase_ids")
          continue;
        var value = JSONPr[key];
        if (value instanceof Array && value.length === 2 && typeof value[1] === "string")value = value[1];
        if(value===false)value=" ";
        Prfield.push(value);
      }
      tabRender.push(Prfield);
      return tabRender;
    }else{
      for(var indexPo_Sp = 0; indexPo_Sp < Po_Sp.length; indexPo_Sp++){
        if(Prfield.length === 0){
            for(var key in JSONPr){
              if(key==="Po" || key==="id" || key==="purchase_ids")
                  continue;
              var value = JSONPr[key];
              if (value instanceof Array && value.length === 2 && typeof value[1] === "string")value = value[1];
              if(value===false)value=" ";
              Prfield.push(value);
            }
        }
        tabRender.push(Array.concat(Prfield,Po_Sp[indexPo_Sp]));
      }
    }
    return tabRender;
}
function plotData(JsonData){
  var data = [];
  var lines = [];
  for(var indexData = 0; indexData < JsonData.length; indexData++){
      lines=renderPr_Po_Sp(JsonData[indexData],renderPo_Sp(JsonData[indexData]));
      for(var key in lines){
          data.push(lines[key]);
      }
  }
  return data;
}

function getData(){
tabPr = custom_oe_browse("purchase.requisition","company_id name x_studio_field_jOoPe  user_id ordering_date date_end schedule_date origin state x_studio_description purchase_ids","[]");
//    tabPr = custom_oe_browse("purchase.requisition","company_id name x_studio_field_jOoPe  user_id ordering_date date_end schedule_date origin state x_studio_description purchase_ids","[['state','not in',['draft','cancel']],"+filterDate()+"]");
    tabPo = custom_oe_browse("purchase.order","requisition_id name partner_id partner_ref x_studio_description date_approve date_planned user_id origin amount_untaxed amount_total currency_id state id picking_ids","[['requisition_id', '!=', 0]]");
    tabSp = custom_oe_browse("stock.picking","purchase_id name partner_id write_uid scheduled_date date_done backorder_id state priority picking_type_id","[]")
    JsonData = joinPr_Po_Sp(tabPr,tabPo,tabSp);
//    Logger.log(tabPr);
//    Logger.log(tabPo);
//    Logger.log(tabSp);  
    return plotData(JsonData);
}

function cleanData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var FUP_Sheet = ss.getActiveSheet();
  var cell = FUP_Sheet.getRange("A10");
  cell.clearContent();
  ss.toast("Mise à jours des données prete cliquer sur le bouton actualiser ...", "READY TO UPDATE DATA");
}
function updateData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var FUP_Sheet = ss.getActiveSheet();
  var cell = FUP_Sheet.getRange("A10");
  ss.toast("Mise à jours des données ...", "UPDATE DATA");
  cell.setFormula("=getData()");
  var date = Utilities.formatDate(new Date(), "GMT+2", "dd/MM/yyyy hh:mm,ss");
  var Last_update = FUP_Sheet.getRange("E1");
  Last_update.setValue(date);
}
function filterDate(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var FUP_Sheet = ss.getActiveSheet();
  var cellFrom = FUP_Sheet.getRange("B4");
  var cellTo = FUP_Sheet.getRange("E4");
  var domain = "['write_date','>=','"+Utilities.formatDate(cellFrom.getValue(),"GMT+2", "yyyy-MM-dd")+"']";
  Logger.log(domain);
  return domain;
}
function custom_oe_browse(model, fields, domain, sort, limit){
  if(typeof model !== "string"){
    throw "model arg expecting string";
  }
  if(typeof fields !== "string"){
    throw "fields arg expecting comma separated field names";
  }
  if (!domain) domain = "[]";
  if(typeof domain !== "string"){
    throw "domain arg expecting string";
  }
  if(sort && typeof sort !== "string"){
    throw "sort arg expecting string";
  }
  if(limit && typeof limit !== "number"){
    throw "limit arg expecting number";
  }
  
  fields = fields.replace(/\s+/g, ",").split(",");
  if(domain) {
    domain = domain.replace(/\'/g, '"');
  }
  domain = JSON.parse(domain);
  
  var records = seach_read(model, fields, domain, sort, limit);
//  return parse_records_for_ss(records, fields);
  Logger.log(records);
  Logger.log("éééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééééé");
  return records;
}
