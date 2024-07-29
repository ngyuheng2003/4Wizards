/* Constant and variable */
const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const monthsNum = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"];

/* Cache */
function setSupplierID(e) {
  PropertiesService.getScriptProperties().setProperty('supplier_id', e);
}

function getSupplierID() {
  return PropertiesService.getScriptProperties().getProperty('supplier_id');
}

/* General */
function getScriptUrl(){
  return ScriptApp.getService().getUrl();
}

function doGet(request) {
  var action = request.parameter.action;
  var employeeName = request.parameter.employeeName;

  if (action === "update" && employeeName) {
    var template = HtmlService
    .createTemplateFromFile("payrollOnhold");
    template.employeeName = employeeName;
    return template
    .evaluate()
    .setTitle("Finance Wizards | Payroll");
  }else if (action === "seePDetails") {
    var template = HtmlService.createTemplateFromFile("payrollPDetails");
    template.employeeName = employeeName;
    return template.evaluate()
    .setTitle("Finance Wizards | Payroll");
    
  }else if (action === "seePTDetails") {
    var template = HtmlService.createTemplateFromFile("payrollPTDetails");
    template.employeeName = employeeName;
    return template.evaluate()
    .setTitle("Finance Wizards | Payroll");
    
  }else if(action=="pay"){
      var template = HtmlService.createTemplateFromFile("payrollPay");
    template.employeeName = employeeName;
    return template.evaluate()
    .setTitle("Finance Wizards | Payroll");
  }

  // Navigation
  if(request.parameters.v == "payroll"){
    return HtmlService
            .createTemplateFromFile("payroll")
            .evaluate()
            .setTitle("Finance Wizards | Payroll")
  }
  else if(request.parameters.v == "inventory"){
    return HtmlService
            .createTemplateFromFile("inventory")
            .evaluate()
            .setTitle("Finance Wizards | Inventory")
  }
  else if(request.parameters.v == "supplier"){
    return HtmlService
            .createTemplateFromFile("supplier")
            .evaluate()
            .setTitle("Finance Wizards | Supplier")
  }
  else if(request.parameters.v == "profile"){
    return HtmlService
            .createTemplateFromFile("profile")
            .evaluate()
            .setTitle("Finance Wizards | Profile")
  }
  else if(request.parameters.v == "invoices"){
    return HtmlService
            .createTemplateFromFile("invoices")
            .evaluate()
            .setTitle("Finance Wizards | Invoices")
  }
  else if(request.parameters.v == "invoices-payment"){
    return HtmlService
            .createTemplateFromFile("invoices-payment")
            .evaluate()
            .setTitle("Finance Wizards | Payment")
  }
  else if(request.parameters.v == "receipt"){
    return HtmlService
            .createTemplateFromFile("invoices-receipt")
            .evaluate()
            .setTitle("Finance Wizards | Receipt")
  }
  else if(request.parameters.v == "receiptList"){
    return HtmlService
            .createTemplateFromFile("invoices-receiptList")
            .evaluate()
            .setTitle("Finance Wizards | Receipt")
  }
  else if(request.parameters.v == "new-invoice"){
    return HtmlService
            .createTemplateFromFile("invoices2")
            .evaluate()
            .setTitle("Finance Wizards | New Invoice")
  }
  else if(request.parameters.v == "new-invoice-2"){
    return HtmlService
            .createTemplateFromFile("invoices3")
            .evaluate()
            .setTitle("Finance Wizards | New Invoice")
  }
  else if(request.parameters.v == "new-employee"){
    return HtmlService
            .createTemplateFromFile("payroll2")
            .evaluate()
            .setTitle("Finance Wizards | New Employee")
  }
  else if(request.parameters.v == "pay"){
    return HtmlService
            .createTemplateFromFile("payrollPay")
            .evaluate()
            .setTitle("Finance Wizards | Payroll ")
  }
  else if(request.parameters.v == "details"){
    return HtmlService
            .createTemplateFromFile("payrollDetails")
            .evaluate()
            .setTitle("Finance Wizards | Payroll ")
  }
  else if(request.parameters.v == "update"){
    return HtmlService
            .createTemplateFromFile("payrollOnhold")
            .evaluate()
            .setTitle("Finance Wizards | Payroll ")
  }
  else if(request.parameters.v == "pp"){
    return HtmlService
            .createTemplateFromFile("pp")
            .evaluate()
            .setTitle("Finance Wizards | PP ")
  }
  else if(request.parameters.v == "new-pp"){
    return HtmlService
            .createTemplateFromFile("requestform")
            .evaluate()
            .setTitle("Finance Wizards | New Procurement Plan")
  }
  else if (request.parameter.v === 'procurementdetails') {
    var requestId = request.parameter.id;
    var htmlOutput = HtmlService.createHtmlOutput();
    htmlOutput.setContent(generateProcurementDetailsTable(requestId));
    return htmlOutput.setTitle("Finance Wizards | Procurement Details");
  }
  else if(request.parameters.v == "approved"){
    return HtmlService
            .createTemplateFromFile("approved")
            .evaluate()
            .setTitle("Finance Wizards | Approved Table")
  }
  else if(request.parameters.v == "rejected"){
    return HtmlService
            .createTemplateFromFile("rejected")
            .evaluate()
            .setTitle("Finance Wizards | Rejected Table")
  }
  else if(request.parameters.v == "po2"){
    return HtmlService
            .createTemplateFromFile("purchaseOrder2")
            .evaluate()
            .setTitle("Finance Wizards | Purchase Order")
  }
  else if(request.parameters.v == "po"){
    return HtmlService
            .createTemplateFromFile("purchaseOrder")
            .evaluate()
            .setTitle("Finance Wizards | Details")
  }
  else if(request.parameters.v == "pot"){
    return HtmlService
            .createTemplateFromFile("poTable")
            .evaluate()
            .setTitle("Finance Wizards | Purchase Order")
  }
  else{
    return HtmlService
            .createTemplateFromFile("home")
            .evaluate()
            .setTitle("Finance Wizards | Home")
  }
}

// get content to html
function include(filename) {
  return HtmlService
          .createHtmlOutputFromFile(filename)
          .getContent();
}

// get sheet
function getSheet(id, name){
  let sheet = SpreadsheetApp.openById(id);
  let info = sheet.getSheetByName(name);

  return info;
}

// get raw data
function getRawData(id, name){
  return getSheet(id, name).getDataRange().getValues();
}

// get data (json) from sheet
function getSheetData2(id, name, type, index, e){
  let data = getRawData(id, name);
  let headers = data[0];
  let jsonData = [];

  for(let i = 1; i < data.length; i++){
    let row = data[i];
    let rowObject = {};

    if(type == 0){
      for(let j = 0; j < headers.length; j++) {
        rowObject[headers[j]] = row[j];
      }
      jsonData.push(rowObject);
    }
    else if(type == 1 && data[i][index] == e){
      for(let j = 0; j < headers.length; j++) {
        rowObject[headers[j]] = row[j];
      }
      jsonData.push(rowObject);
    }
  }

  return jsonData;
}

// get option for input
function getOption(id, name, feature1, feature2){
  let jsonData = getSheetData2(id, name, 0, 0, 0);
  let tags = "";
  for(let d of jsonData){
    if(d[feature1] != ""){
      if(arguments.length == 3){
        tags += `
        <option>${d[feature1]}</option>
        `
      }
      else if(arguments.length == 4){
        tags += `
        <option>${d[feature1]} ${d[feature2]}</option>
        `
      }
    }
  }
  return tags;
}

// generate id (invoices, receipt, payslip)
function generateID(rawData, index, prefix){
  let todayDate = new Date().getFullYear() + monthsNum[new Date().getMonth()] + new Date().getDate();

  for(let i = rawData.length-1; i >=0; i--){
    if(rawData[i][index].includes(todayDate)){
      let id = parseInt(rawData[i][index].substring(3)) + 1;
      return prefix + id;
    }
  }
  
  return prefix + todayDate + "0001";
}


/* Table Generation */

// get html for table (No data to show)
function getNoDataHtml(col){
  return  ` 
          <tr>
            <td colspan=`+ col +` style="text-align: center;">No data to show</td>
          </tr>
          `
}

// Get date (dd MMMM yyyy)
function getFormattedDate(data){
  return data.getDate() + " " + months[data.getMonth()] + " " + data.getFullYear();
}

// Get currency (0.00)
function getFormattedCurrency(data){
  return data.toFixed(2);
}

// Get address
function getAddress(data){
  return data['Address'] + ", " + data['Postcode'] + " " + data['City'] + ", " + data['Province'] + ", " + data['Country']
}

// Generate table
function generateTable(id, name, type, index, e, col, features){
  let jsonData = getSheetData2(id, name, type, index, e);

  let tags = "";
  if(jsonData.length == 0){
    return getNoDataHtml(col);
  }
  else{
    for(let d of jsonData){
      tags += "<tr>"
      for(let feature of features){
        if(feature == 'Address'){
          tags += `
                    <td>`+ getAddress(d) +`</td>
                  `
        }
        else if(feature == 'Amount'){
          tags += `
                    <td>RM `+ getFormattedCurrency(d[feature]) +`</td>
                  `
        }
        else if(feature.includes('date')){
          tags += `
                    <td>`+ getFormattedDate(d[feature]) +`</td>
                  `
        }
        else if(feature == 'Detail'){
          tags += `
                    <td>
                      <button onclick="nextDetailPage(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer"><img src="https://drive.google.com/thumbnail?id=15oz_9RrbWCvE_gjlS_K_ohyFlVAMhMXY" title="See details" style="height:25px;"></Button>
                    </td>
                  `
        }
        else{
          tags += `
                    <td>${d[feature]}</td>
                  `
        }
      }
      tags += "</tr>"
    }
  }
  return tags;
}

/* Home */
function getTotalProfitMonth() {
  let jsonData = getSheetData2("1Mf-ZHyfGki5pZc0A1MPNByO4POTmuQeMCO1mYIE_3xw", "Profit", 1, 0, months[new Date().getMonth()]);
  return jsonData[0]['Profit / Loss'].toFixed(2);
}

function getDailyProfit() {
  return getSheetData2("1Mf-ZHyfGki5pZc0A1MPNByO4POTmuQeMCO1mYIE_3xw", "Profit", 1, 0, months[new Date().getMonth()]);
}

/* Inventory */

function getItemOption() {
  return getOption("10JUBz_zmt1aDG_Sl7KfCQ7K08yqAfTjzYyRAQDLtuWc", "Product", "Product ID", "Product Name");
}

// Generate Supplier Table
function generateSupplierTable() {
  let features = ['Supplier ID', 'Supplier name', 'Address', 'Rating', 'Detail']

  return generateTable("1ns-x6_wP_Ibxr1Aab7O9r3K-CvKUcL2Jk1pX8DCDufM", "Supplier", 0, 0, 0, 5, features);
}

// Get supplier data
function getSupplierInfo(e) {
  let jsonData = getSheetData2("1ns-x6_wP_Ibxr1Aab7O9r3K-CvKUcL2Jk1pX8DCDufM", "Supplier", 1, 0, getSupplierID());

  if (jsonData[0][e] == "") {
    return "No data to show";
  }
  else if (e == "Address") {
    return getAddress(jsonData[0]);
  }
  else {
    return jsonData[0][e];
  }
  
}

function getChartInfo(e){
  const array = ["chart_response_rate", "chart_delivery_rate", "chart_quality_rate"];
  const array2 = ["Response rate", "Delivery rate", "Product quality"];
  let jsonData = getSheetData2("1ns-x6_wP_Ibxr1Aab7O9r3K-CvKUcL2Jk1pX8DCDufM", "Supplier", 1, 0, getSupplierID());

  return [array[e], [jsonData[0][array2[e]], 100 - jsonData[0][array2[e]]], e];
}

// get stock info
function getStockInfo(form){
  const str = form.stock_id.split(" ");
  let jsonData = getSheetData2("1ns-x6_wP_Ibxr1Aab7O9r3K-CvKUcL2Jk1pX8DCDufM", "Product", 1, 0, str[0]);
  return jsonData;
}

/* Invoices */
// generate inoivce table
function generateInvoicesTable(){
  // get data from google sheets
  let jsonData = getSheetData2("1Wq0pDQquiUEWVmwjJ_vatJRewPSxsrAcTxNiMKKSw0I", new Date().getFullYear(), 0, 0, 0);

  // create table
  let tags = "";
  if(jsonData.length == 0){
    return getNoDataHtml(7);
  }
  else{
    for(let d of jsonData){
    tags += `<tr>
      <td>${d['Invoice ID']}</td>
      <td>${d.Company_Buyer_Name}</td>
      <td>${getFormattedDate(d.IssueDate)}</td>
      <td>${getFormattedDate(d.DueDate)}</td>
      <td>RM ${getFormattedCurrency(d['Amount'])}</td>
      `
    if(d["Balance"] > 0 && d.DueDate - new Date() < 0){
      tags += `
      <td style="padding: 0 30px;"><p id= "unpaid" title="This payment is unpaid before the due date">Late<p></td>`
      if(d["Balance"] == d["Amount"])
      tags += `
      <td>
      <button onclick="makePayment(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer"><img src="https://drive.google.com/thumbnail?id=1XJR-A_HECzXF0WcCL7Nu7qJWZ2UjRgYA" title="Recieve money" style="height:25px;"></Button>
      </tr>`
      else{
        tags += `
      <td>
      <button onclick="makePayment(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer"><img src="https://drive.google.com/thumbnail?id=1XJR-A_HECzXF0WcCL7Nu7qJWZ2UjRgYA" title="Recieve money" style="height:25px;"></Button>
      <button onclick="nextInvoicesList(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer"><img src="https://drive.google.com/thumbnail?id=15oz_9RrbWCvE_gjlS_K_ohyFlVAMhMXY" title="See details" style="height:25px;"></Button></td>
      </tr>`
      }
    }
    else if(d["Balance"] == 0){
      tags += `
      <td style="padding: 0 30px;"><p id= "paid" title="Salary has been paid to this employee">Recieved<p></td>`
      tags += `
      <td>
      <button onclick="nextInvoicesList(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer"><img src="https://drive.google.com/thumbnail?id=15oz_9RrbWCvE_gjlS_K_ohyFlVAMhMXY" title="See details" style="height:25px;"></Button></td>
      </tr>`
    }
    else if(d["Balance"] > 0 ){
      tags += `
      <td style="padding: 0 30px;"><p id= "onHold" title="Information required before performing a payout">Pending<p></td>`
      if(d["Balance"] == d["Amount"])
      tags += `
      <td>
      <button onclick="makePayment(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer"><img src="https://drive.google.com/thumbnail?id=1XJR-A_HECzXF0WcCL7Nu7qJWZ2UjRgYA" title="Recieve money" style="height:25px;"></Button>
      </tr>`
      else{
        tags += `
      <td>
      <button onclick="makePayment(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer"><img src="https://drive.google.com/thumbnail?id=1XJR-A_HECzXF0WcCL7Nu7qJWZ2UjRgYA" title="Recieve money" style="height:25px;"></Button>
      <button onclick="nextInvoicesList(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer"><img src="https://drive.google.com/thumbnail?id=15oz_9RrbWCvE_gjlS_K_ohyFlVAMhMXY" title="See details" style="height:25px;"></Button></td>
      </tr>`
      }
    }
     
    }
  }
  return tags;
}

// Generate receipt table
function generateReceiptTable(e){
  let features = ['Receipt ID', 'Company / Buyer name', 'Payment date', 'Amount', 'Detail']

  return generateTable("1sgMmwmn6wmAB61cqQFUBnXImiO_eqsF9moyT3OjIz50", new Date().getFullYear(), 1, 1, e, 5, features);
}


// print invoice details
function getInvoicesInfo(e){
  // get data from sheet
  let jsonData = getSheetData2("1Wq0pDQquiUEWVmwjJ_vatJRewPSxsrAcTxNiMKKSw0I", new Date().getFullYear(), 1, 0, e);

  var html = `
      <div class="company-info">
            <pre>Invoice ID&#09&#09&#09&#09:&#09</pre>
            <pre style="font-weight: bold; text-transform: uppercase;">`+ jsonData[0]["Invoice ID"] +`</pre>
          </div>
          <br>
          <div class="company-info">
            <pre>Company / Buyer name&#09&#09:&#09</pre>
            <pre style="font-weight: bold; text-transform: uppercase;">`+ jsonData[0]["Company_Buyer_Name"] +`</pre>
          </div>
          <br>
          <div class="company-info">
            <pre>Due date&#09&#09&#09&#09&#09:&#09</pre>
            <pre style="font-weight: bold; text-transform: uppercase;line-height: 25px;">`+ jsonData[0]["DueDate"].getDate() + ` ` + months[jsonData[0]["DueDate"].getMonth()] + ` ` + jsonData[0]["DueDate"].getFullYear() +`</pre>
          </div>
          <br>
          <div class="company-info">
            <pre>Amount&#09&#09&#09&#09&#09:&#09</pre>
            <pre style="font-weight: bold;">RM `+ jsonData[0]["Amount"].toFixed(2) +`</pre>
          </div>
          <br>
          <div class="company-info">
            <pre>Balance&#09&#09&#09&#09&#09:&#09</pre>
            <pre style="font-weight: bold;">RM `+ jsonData[0]["Balance"].toFixed(2) +`</pre>
          </div>
      `

      return html;
}

function getPaymentInfo(e){
  // get data from sheet
  let jsonData = getSheetData2("1sgMmwmn6wmAB61cqQFUBnXImiO_eqsF9moyT3OjIz50", new Date().getFullYear(), 1, 0, e);

  let html = `
      <div class="company-info">
            <pre>Receipt ID&#09&#09&#09&#09:&#09</pre>
            <pre style="font-weight: bold; text-transform: uppercase;">`+ jsonData[0]["Receipt ID"] +`</pre>
          </div>
          <br>
          <div class="company-info">
            <pre>Company / Buyer name&#09&#09:&#09</pre>
            <pre style="font-weight: bold; text-transform: uppercase;">`+ jsonData[0]["Company / Buyer name"] +`</pre>
          </div>
          <br>
          <div class="company-info">
            <pre>Payment date&#09&#09&#09&#09:&#09</pre>
            <pre style="font-weight: bold; text-transform: uppercase;line-height: 25px;">`+ jsonData[0]["Payment date"].getDate() + ` ` + months[jsonData[0]["Payment date"].getMonth()] + ` ` + jsonData[0]["Payment date"].getFullYear() +`</pre>
          </div>
          <br>
          <div class="company-info">
            <pre>Payment method&#09&#09&#09:&#09</pre>
            <pre style="font-weight: bold;">`+ jsonData[0]["Payment method"] +`</pre>
          </div>
          <br>
          <div class="company-info">
            <pre>Type of payment&#09&#09&#09:&#09</pre>
            <pre style="font-weight: bold;">`+ jsonData[0]["Type of payment"] +`</pre>
          </div>
          <br>
          <div class="company-info">
            <pre>Amount&#09&#09&#09&#09&#09:&#09</pre>
            <pre style="font-weight: bold;">RM `+ jsonData[0]["Amount"].toFixed(2) +`</pre>
          </div>
      `

      return html;
}

function getReceiptPreview(e){
  getSheet("1sgMmwmn6wmAB61cqQFUBnXImiO_eqsF9moyT3OjIz50", "Receipt Preview").getRange("H11").setValue(e);

  let html = `
      <iframe width=100% height=730px style='border:none' src="https://docs.google.com/spreadsheets/d/e/2PACX-1vTTYbVDvzIVmrYzN3ytXDArPKplAsiZhU1ZbqcWjsH-d4fHl0SQVpw3b98nNOFnEbiPXLRv37B2Etox/pubhtml?gid=1926056078&amp;single=true&amp;widget=false&amp;headers=false"></iframe>
      `;

  return html;
}

function getProvinceOption(){
  return getOption("10JUBz_zmt1aDG_Sl7KfCQ7K08yqAfTjzYyRAQDLtuWc", "Location", "Province");
}

function getPaymentMethodOption(){
  return getOption("10JUBz_zmt1aDG_Sl7KfCQ7K08yqAfTjzYyRAQDLtuWc", "Payment Method", "Payment Method");
}

function getPaymentTypeOption(){
  return getOption("10JUBz_zmt1aDG_Sl7KfCQ7K08yqAfTjzYyRAQDLtuWc", "Payment Method", "Type of Payment");
}

// append payment detail into sheets
function updatePaymentInfo(form){
  let receiptInfoSheets = SpreadsheetApp.openById("1sgMmwmn6wmAB61cqQFUBnXImiO_eqsF9moyT3OjIz50");
  let receiptInfo = receiptInfoSheets.getSheetByName(new Date().getFullYear());

  let data = receiptInfo.getDataRange().getValues();

  let invoicesListSheets = SpreadsheetApp.openById("1Wq0pDQquiUEWVmwjJ_vatJRewPSxsrAcTxNiMKKSw0I");
  let invoicesListInfo = invoicesListSheets.getSheetByName(new Date().getFullYear());

  let data2 = invoicesListInfo.getDataRange().getValues();

  // generate receipt id
  let receipt_id = generateID(data, 0, "FWR");

  // add info to sheets
  let data_array = [receipt_id, form.invoice_id, form.company_name, form.payment_date, form.payment_method, form.type_of_payment, form.reference_id, form.amount, form.remarks];

  getSheet().appendRow(data_array);
  
  for(let i = 1; i < data2.length; i++){
    if(data2[i][0] == form.invoice_id){
      let index = 'F' + (i + 1);
      let balance = data2[i][5] - form.amount;
      invoicesListInfo.getRange(index).setValue(balance);
      break;
    }
  }

  let url = getScriptUrl() + "?v=receipt";
  let array = [url, receipt_id]
  
  return array
}

/* Payroll */

function getEPFContribution(dob, monthlySalary, status) {
    // Calculate current age
    const currentYear = new Date().getFullYear();
    const birthYear = new Date(dob).getFullYear();
    const currentAge = currentYear - birthYear;


    let employeeContribution = 0;
    let employerContribution = 0;

    // Determine EPF contribution rates based on age and conditions
    if (currentAge >= 60) {
        // Stage 2 (Age 60 and above)
        if (status === "PERMANENT RESIDENT" || status === "NON-MALAYSIAN") {
            employeeContribution = 5.5;
            employerContribution = 6.5;
        } else if (status === "MALAYSIAN") {
            employeeContribution = 0;
            employerContribution = 4;
        } else {
            employeeContribution = 5.5;
            employerContribution = 5;
        }
    } else {
        // Stage 1 (Below 60 years old)
        if (monthlySalary <= 5000) {
            employeeContribution = 11;
            employerContribution = 13;
        } else if (monthlySalary > 5000) {
            employeeContribution = 11;
            employerContribution = 12;
        }
    }

    return {
        employeeContribution: employeeContribution,
        employerContribution: employerContribution
    };
}

const socsoRatesCat1 = [
  { min: 0, max: 30, employee: 0.10, employer: 0.40 },
  { min: 30, max: 50, employee: 0.20, employer: 0.70 },
  { min: 50, max: 70, employee: 0.30, employer: 1.10 },
  { min: 70, max: 100, employee: 0.40, employer: 1.50 },
  { min: 100, max: 140, employee: 0.60, employer: 2.10 },
  { min: 140, max: 200, employee: 0.85, employer: 2.95 },
  { min: 200, max: 300, employee: 1.25, employer: 4.35 },
  { min: 300, max: 400, employee: 1.75, employer: 6.15 },
  { min: 400, max: 500, employee: 2.25, employer: 7.85 },
  { min: 500, max: 600, employee: 2.75, employer: 9.65 },
  { min: 600, max: 700, employee: 3.25, employer: 11.35 },
  { min: 700, max: 800, employee: 3.75, employer: 13.15 },
  { min: 800, max: 900, employee: 4.25, employer: 14.85 },
  { min: 900, max: 1000, employee: 4.75, employer: 16.65},
  { min: 1000, max: 1100, employee: 5.25, employer: 18.35 },
  { min: 1100, max: 1200, employee: 5.75, employer: 20.15 },
  { min: 1200, max: 1300, employee: 6.25, employer: 21.85 },
  { min: 1300, max: 1400, employee: 6.75, employer: 23.65 },
  { min: 1400, max: 1500, employee: 7.25, employer: 25.35 },
  { min: 1500, max: 1600, employee: 7.75, employer: 27.15 },
  { min: 1600, max: 1700, employee: 8.25, employer: 28.85 },
  { min: 1700, max: 1800, employee: 8.75, employer: 30.65 },
  { min: 1800, max: 1900, employee: 9.25, employer: 32.35 },
  { min: 1900, max: 2000, employee: 9.75, employer: 34.15 },
  { min: 2000, max: 2100, employee: 10.25, employer: 35.85 },
  { min: 2100, max: 2200, employee: 10.75, employer: 37.65 },
  { min: 2200, max: 2300, employee: 11.25, employer: 39.35 },
  { min: 2300, max: 2400, employee: 11.75, employer: 41.15 },
  { min: 2400, max: 2500, employee: 12.25, employer: 42.85 },
  { min: 2500, max: 2600, employee: 12.75, employer: 44.65 },
  { min: 2600, max: 2700, employee: 13.25, employer: 46.35 },
  { min: 2700, max: 2800, employee: 13.75, employer: 48.15 },
  { min: 2800, max: 2900, employee: 14.25, employer: 49.85 },
  { min: 2900, max: Number.MAX_SAFE_INTEGER, employee: 14.75, employer: 51.65 },
  
];

const socsoRatesCat2 = [
  { min: 0, max: 30, employee: 0, employer: 0.30 },
  { min: 30, max: 50, employee: 0, employer: 0.50 },
  { min: 50, max: 70, employee: 0, employer: 0.80 },
  { min: 70, max: 100, employee: 0, employer: 1.10 },
  { min: 100, max: 140, employee: 0, employer: 1.50 },
  { min: 140, max: 200, employee: 0, employer: 2.10 },
  { min: 200, max: 300, employee: 0, employer: 3.10 },
  { min: 300, max: 400, employee: 0, employer: 4.10 },
  { min: 400, max: 500, employee: 0, employer: 5.60 },
  { min: 500, max: 600, employee: 0, employer: 6.90 },
  { min: 600, max: 700, employee: 0, employer: 8.10 },
  { min: 700, max: 800, employee: 0, employer: 9.40 },
  { min: 800, max: 900, employee: 0, employer: 10.60 },
  { min: 900, max: 1000, employee: 0, employer: 11.90},
  { min: 1000, max: 1100, employee: 0, employer: 13.10 },
  { min: 1000, max: 1100, employee: 0, employer: 14.40 },
  { min: 1200, max: 1300, employee: 0, employer: 15.60 },
  { min: 1300, max: 1400, employee: 0, employer: 16.90 },
  { min: 1400, max: 1500, employee: 0, employer: 18.10 },
  { min: 1500, max: 1600, employee: 0, employer: 19.40 },
  { min: 1600, max: 1700, employee: 0, employer: 20.60 },
  { min: 1700, max: 1800, employee: 0, employer: 31.90 },
  { min: 1800, max: 1900, employee: 0, employer: 23.10 },
  { min: 1900, max: 2000, employee: 0, employer: 24.40 },
  { min: 2000, max: 2100, employee: 0, employer: 25.60 },
  { min: 2100, max: 2200, employee: 0, employer: 26.90 },
  { min: 2200, max: 2300, employee: 0, employer: 28.10 },
  { min: 2300, max: 2400, employee: 0, employer: 29.40},
  { min: 2400, max: 2500, employee: 0, employer: 30.60 },
  { min: 2500, max: 2600, employee: 0, employer: 31.90 },
  { min: 2600, max: 2700, employee: 0, employer: 33.10 },
  { min: 2700, max: 2800, employee: 0, employer: 34.40 },
  { min: 2800, max: 2900, employee: 0, employer: 35.60 },
  { min: 2900, max: Number.MAX_SAFE_INTEGER, employee: 0, employer: 36.90 },
  
];

function getSocsoContribution(dob, yearJoined, monthlySalary) {
    // Calculate current age
    const currentYear = new Date().getFullYear();
    const birthYear = new Date(dob).getFullYear();
    const currentAge = currentYear - birthYear;

    // Calculate age when joined
    const ageWhenJoined = yearJoined - birthYear;

    // Determine SOCSO category based on age and status
    let selectedRates = socsoRatesCat1;

    // Check if employee is in category 2 (age 60 or above or non-Malaysian)
    if (currentAge >= 60 || ageWhenJoined>=55) {
        selectedRates = socsoRatesCat2;
    }

    // Find the applicable SOCSO rate based on the salary range
    for (let rate of selectedRates) {
        if (monthlySalary >= rate.min && monthlySalary < rate.max) {
            return {
                employeeContribution: rate.employee,
                employerContribution: rate.employer,
            };
        }
    }

    // If salary exceeds the range, use the highest range's rate
    const highestRate = selectedRates[selectedRates.length - 1];
    return {
        employeeContribution: highestRate.employee,
        employerContribution: highestRate.employer,
    };
}


function checkPay(name, type){
  var payrollHistory = SpreadsheetApp.openById("1RplPsRWxIPWdKPJ-pyzkNUPYfzqkWJu00Fla1Q1XY_E");
  const currentM=months[new Date().getMonth()];
  var monthSheet = getOrCreateSheet(payrollHistory.getId(), currentM);
  var data = monthSheet.getDataRange().getValues();


  for (var i = 1; i < data.length; i++) {
    console.log(data[i][2]);
    if (data[i][2] === name) {
      return 1
    }
  }
  if(type=='Permanent'){
        return 0; 
  }else{
        return 2;
  }
}

function generateEmployeeInfoTable(){
  var employeeListSheets = SpreadsheetApp.openById("1FDJBnqvEEz1rJQ3h7wOKNZea9cxCJXl9IOCpB6_4HQY");
  var employeePListInfo = employeeListSheets.getSheets()[0];
  var employeePTListInfo = employeeListSheets.getSheets()[1];

  var data = employeePListInfo.getDataRange().getValues();
  var headers = data[0];
  var jsonData = [];

  for(var i = 1; i < data.length; i++){
    var row = data[i];
    var rowObject = { Type: 'Permanent' }; // Add EmploymentStatus field
    rowObject.Status = checkPay(row[1], 'Permanent'); // Assuming name is in the first column

    for (var j = 0; j < headers.length; j++) {
      rowObject[headers[j]] = row[j];
    }
    jsonData.push(rowObject);
  }

  var dataPT = employeePTListInfo.getDataRange().getValues();
  var headersPT = dataPT[0];
  for(var i = 1; i < dataPT.length; i++){
    var row = dataPT[i];
    var rowObject = { Type: 'Part Time' }; // Add EmploymentStatus field
    rowObject.Status = checkPay(row[0], 'Part Time'); // Assuming name is in the first column

    for (var j = 0; j < headersPT.length; j++) {
      rowObject[headersPT[j]] = row[j];
    }
    jsonData.push(rowObject);
  }

  var tags = "";
  if(jsonData.length == 0){
    tags += `<tr>
                <td colspan="5" style="text-align: center;">No data to show</td>
            </tr>`;
  } else {
    for(let d of jsonData){
      console.log(d);
      tags += `<tr>
        <td>${d.Name}</td>
        <td>${d.Position}</td>
        <td>${d.Type}</td>`;
      
      if(d.Status == 0){
        
        tags += `
        <td><p id="unpaid" title="Salary has not been paid to this employee">Unpaid<p></td>`;
        
        tags += `
        <td>
        <a href="` +getScriptUrl() + `?action=pay&employeeName=${encodeURIComponent(d.Name)}" style="margin: 0 8px"><img src="https://drive.google.com/thumbnail?id=1OOMYawyvpwQidsaWZr5-JTghicopv05m" title="Pay" style="height:20px;"></a>
        </td></tr>`;
      } else if(d.Status == 1 && d.Type=='Permanent'){
        tags += `
        <td><p id="paid" title="Salary has been paid to this employee">Paid<p></td>`;
        tags += `
        <td><a href="` + getScriptUrl() + `?action=seePDetails&employeeName=${encodeURIComponent(d.Name)}"><img src="https://drive.google.com/thumbnail?id=15oz_9RrbWCvE_gjlS_K_ohyFlVAMhMXY" title="See details" style="height:25px;"></a></td>
        </tr>`;
      }else if(d.Status == 1 && d.Type=='Part Time'){
        tags += `
        <td><p id="paid" title="Salary has been paid to this employee">Paid<p></td>`;
        tags += `
        <td><a href="` + getScriptUrl() + `?action=seePTDetails&employeeName=${encodeURIComponent(d.Name)}"><img src="https://drive.google.com/thumbnail?id=15oz_9RrbWCvE_gjlS_K_ohyFlVAMhMXY" title="See details" style="height:25px;"></a></td>
        </tr>`;
      } else if(d.Status == 2){
        tags += `
        <td><p id="onHold" title="Information required before performing a payout">On Hold<p></td>`;
        tags += `
        <td>
         <a href="`+ getScriptUrl() + `?action=update&employeeName=${encodeURIComponent(d.Name)}" style="margin: 0 8px"><img src="https://drive.google.com/thumbnail?id=1januWH_E_wxhhvykwk8WhfJy5hEN3wF-" title="Update required information" style="height:20px;"></a>
        </td></tr>`;
      }
    }
  }
  return tags;
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd MMMM yyyy");
}

function getStaffInfo(index, employeeName) {
  var sheet = SpreadsheetApp.openById("1FDJBnqvEEz1rJQ3h7wOKNZea9cxCJXl9IOCpB6_4HQY");
  var employeePListInfo = sheet.getSheets()[1];
  var data = employeePListInfo.getDataRange().getValues();

  // Assuming the first row is the header and data starts from row 1
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === employeeName) { // Assuming employee name is in the second column (adjust as needed)
      return data[i][index];
    }
  }
  return 'N/A'; // Default value if no data found
}

function getEmployeeData(index,employeeName){
  var sheet = SpreadsheetApp.openById("1FDJBnqvEEz1rJQ3h7wOKNZea9cxCJXl9IOCpB6_4HQY");
  var employeePListInfo = sheet.getSheets()[0];
  var data = employeePListInfo.getDataRange().getValues();

  // Assuming the first row is the header and data starts from row 1
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === employeeName) { 
      return data[i][index];
    }
  }
  return 'N/A'; // Default value if no data found
}

function getPayData(index,employeeName){
   var payrollHistory = SpreadsheetApp.openById("1RplPsRWxIPWdKPJ-pyzkNUPYfzqkWJu00Fla1Q1XY_E");
  var monthSheet = payrollHistory.getSheetByName(months[new Date().getMonth()]);
  var data = monthSheet.getDataRange().getValues();


  // Assuming the first row is the header and data starts from row 1
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === employeeName) { 
      var value = data[i][index];
      // Check if the value is a Date object or a formatted date string
      if (value instanceof Date) {
        // Format Date object as '25 July 2024'
        return Utilities.formatDate(value, Session.getScriptTimeZone(), "d MMMM yyyy");
      } else {
        // Return value as is if it's not a Date object
        return value;
      }
    }
  }
  return 'N/A'; // Default value if no data found
}

// Function to get or create a sheet in the spreadsheet
function getOrCreateSheet(spreadsheetId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    let sheetExists = false;
    
    // Check if the sheet already exists
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === sheetName) {
        sheetExists = true;
        break;
      }
    }
    
    // Create the sheet if it doesn't exist
    if (!sheetExists) {
       sheet = ss.insertSheet(sheetName);
      Logger.log('Sheet created: ' + sheetName);
      // Append the header row
      const headers = [
        "Payroll ID", "EmployeeID", "Name", "NRIC", "Position", "Type", 
        "Basic Salary", "EPF No", "SOCSO No", "Bank", "Acc No", "Date", 
        "DayWorks/ Hour Work", "NormalOT(hour)", "RestOT", "PublicHolidayOT", 
        "Overtime pay", "Medical Leave", "Annual Leave", "EPFemployee", 
        "EPFemployer", "Socsoemployee", "Socsoemployer", "Claims", 
        "Allowances", "OtherDeduction", "Remark", "Gross Salary", 
        "Total Deduction", "Net Salary"
      ];
      sheet.appendRow(headers);
      Logger.log('Header row added.');


    } else {
      Logger.log('Sheet already exists: ' + sheetName);
    }
    
    // Return the sheet
    return ss.getSheetByName(sheetName);
  } catch (error) {
    Logger.log('Error: ' + error.message);
    throw new Error('Failed to get or create sheet: ' + error.message);
  }
}

function calcOT(normalOT,restOT,pubHolidayOT,employeeName){
  const basicSalary = parseFloat(getEmployeeData(7, employeeName)) || 0;
   var currentY=new Date().getFullYear();
   var hourlyRate=calculateHourlyRate(basicSalary,currentY,monthsNum[new Date().getMonth()]);

    const dayWorkingHour=8;
    var dayRate=parseFloat(hourlyRate)*parseFloat(dayWorkingHour);

    var pay= normalOT * 1.5 * parseFloat(hourlyRate) + 
            restOT  * 2 * parseFloat(dayRate) + 
            pubHolidayOT  * 3 * parseFloat(dayRate);

     return pay;

}

function calculateWorkingDays(year, month) {
  // Get the first and last days of the month
  var firstDay = new Date(year, month - 1, 1); // month is 0-indexed
  var lastDay = new Date(year, month, 0); // day 0 is the last day of the previous month

  var workingDays = 0;

  for (var day = firstDay.getDate(); day <= lastDay.getDate(); day++) {
    var currentDay = new Date(year, month - 1, day);
    var dayOfWeek = currentDay.getDay();
    if (dayOfWeek >= 1 && dayOfWeek <= 5) { // 1 is Monday, 5 is Friday
      workingDays++;
    }
  }

  return workingDays;
}


function calculateHourlyRate(basicMonthlySalary, year, month) {
  var workingDays = calculateWorkingDays(year, month);
  console.log(workingDays);
  var workingHoursPerDay = 8;
  var totalWorkingHours = workingDays * workingHoursPerDay;

  var hourlyRate = basicMonthlySalary / totalWorkingHours;
  return hourlyRate;
}

/* Add Employee */
function doPost(e) {
  try {
    // Log the received parameters
    Logger.log(JSON.stringify(e.parameter));

    // Get the form data
    const data = e.parameter;

    // Open the spreadsheet by ID and get the first sheet
    const ss = SpreadsheetApp.openById("1FDJBnqvEEz1rJQ3h7wOKNZea9cxCJXl9IOCpB6_4HQY");
    const staffSheet = ss.getSheetByName(data.employmentType);

    // Log spreadsheet information
    Logger.log('Spreadsheet and sheet accessed successfully.');

    if(data.employmentType=="Permanent"){
      // Check if the employee ID already exists
      const employeeID = data.employeeID;
      const existingEmployeeIDs = staffSheet.getRange('A2:A').getValues().flat(); // Assuming employee IDs are in column A
      if (existingEmployeeIDs.includes(employeeID)) {
        Logger.log('Duplicate employee ID detected: ' + employeeID);
        return ContentService.createTextOutput('Error: Employee ID already exists.');
      }
    }
    
    if(data.employmentType=='Part Time'){
      var newRow = [
        data.pt_employeeName,
        data.pt_ic,
        data.pt_position,
        data.pt_email,
        data.pt_bank,
        data.pt_acc
      ];
    }else{
      var newRow = [
        data.employeeID,
        data.employeeName,
        data.statuss||'',
        data.ic,
        data.dob,
        data.yearJoined,
        data.position,
        parseFloat(data.basicSalary),
        parseFloat(data.allowances),
        data.epfNo,
        data.socsoNo,
        data.email,
        data.bank,
        data.acc,
        formatDate(new Date())
      ];
    }

    // Log the new row data
    Logger.log('New row data: ' + JSON.stringify(newRow));

    // Append the new row to the sheet
    staffSheet.appendRow(newRow);

    // Log the success message
    Logger.log('Data appended successfully.');

    // Return a response to the client
    return ContentService.createTextOutput('Data successfully added!');
  } catch (error) {
    // Log any errors that occur
    Logger.log('Error: ' + error.message);
    return ContentService.createTextOutput('An error occurred: ' + error.message);
  }
}

function processFormData(data) {
  // Called by the client-side script
  const response = doPost({parameter: data});
  return response.getContent();
}


/* Onhold Page */
function processWorkingHours(data) {
  // Called by the client-side script
  const response = doPostWH({ parameter: data });
  return response.getContent();
}

function doPostWH(e) {
  try {
    // Log the received parameters
    Logger.log(JSON.stringify(e.parameter));

    // Get the form data
    const data = e.parameter;
    const employeeName=data.employeeName;

    // Open the spreadsheet by ID
    const ss = SpreadsheetApp.openById("1FDJBnqvEEz1rJQ3h7wOKNZea9cxCJXl9IOCpB6_4HQY");
    const staffSheet = ss.getSheets()[1];
    const ssd = staffSheet.getDataRange().getValues();

    // Log the sheet information
    Logger.log('Spreadsheet and sheet accessed successfully.');

    for (let i = 1; i < ssd.length; i++) {
      if (ssd[i][0] === employeeName) { // Assuming employeeName is in the first column
        staffSheet.getRange(i + 1, 7).setValue(data.ptH); 
        Logger.log('Data updated successfully in main spreadsheet.');
        break;
      }
    }

    // Update another spreadsheet
    var payrollHistory = SpreadsheetApp.openById("1RplPsRWxIPWdKPJ-pyzkNUPYfzqkWJu00Fla1Q1XY_E");
    const currentMonth = months[new Date().getMonth()];
    
    const anotherSheet = getOrCreateSheet(payrollHistory.getId(), currentMonth);

    const ptPayRate = 10;
    var netPay = ptPayRate * parseFloat(data.ptH);

    var newRow = [
      generateID(getRawData("1RplPsRWxIPWdKPJ-pyzkNUPYfzqkWJu00Fla1Q1XY_E", months[new Date().getMonth()]),0,"FWP"),
      '-',
      getStaffInfo(0,employeeName),
      getStaffInfo(2,employeeName),
      getStaffInfo(1,employeeName),
      "Part Time",
      '-',
      '-',
      '-',
      getStaffInfo(4,employeeName),
      getStaffInfo(5,employeeName),
      formatDate(new Date()),
      data.ptH,
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      '-',
      netPay
    ];

    // Append the same row to the other spreadsheet
    anotherSheet.appendRow(newRow);

    // Log the success message
    Logger.log('Data appended successfully.');

    // Return a response to the client
    return ContentService.createTextOutput('Data successfully added!');
  } catch (error) {
    // Log any errors that occur
    Logger.log('Error: ' + error.message);
    return ContentService.createTextOutput('An error occurred: ' + error.message);
  }
}



/* Unpaid Page */
function processPayment(data){
    // Called by the client-side script
  const response = doPostP({ parameter: data });
  return response.getContent();
}

function myFunc(){
  var basicSalary=getEmployeeData(7,"Chua Hui Min");
  var currentY=new Date().getFullYear();
   var hourlyRate=calculateHourlyRate(basicSalary,currentY,monthsNum[new Date().getMonth()]);
   console.log(currentY);
   console.log(monthsNum[new Date().getMonth()]);
   console.log(hourlyRate);
   

    const dayWorkingHour=8;
    var dayRate=parseFloat(hourlyRate)*parseFloat(dayWorkingHour);
    console.log(dayRate);
   /* var otPay = parseFloat(data.ot) * 1.5 * parseFloat(hourlyRate) + 
            parseFloat(data.restOT)  * 2 * parseFloat(dayRate) + 
            parseFloat(data.pubHoliday)  * 3 * parseFloat(dayRate);*/


}

function doPostP(e) {
  try {
    // Log the received parameters
    Logger.log(JSON.stringify(e.parameter));

    // Get the form data
    const data = e.parameter;
    const employeeName=data.employeeName;


    // Update another spreadsheet
    var payrollHistory = SpreadsheetApp.openById("1RplPsRWxIPWdKPJ-pyzkNUPYfzqkWJu00Fla1Q1XY_E");
    const currentMonth = months[new Date().getMonth()];
    
    const anotherSheet = getOrCreateSheet(payrollHistory.getId(), currentMonth);

    const basicSalary = parseFloat(getEmployeeData(7, employeeName)) || 0;
    const socsoContribution = getSocsoContribution(getEmployeeData(4, employeeName), getEmployeeData(5, employeeName), basicSalary);
    const socsoEmployee = socsoContribution.employeeContribution;
    const socsoEmployer = socsoContribution.employerContribution;

    const epfContribution = getEPFContribution(getEmployeeData(4, employeeName), basicSalary, getEmployeeData(2, employeeName));
    const epfEmployee = basicSalary * (epfContribution.employeeContribution || 0) / 100;
    const epfEmployer = basicSalary * (epfContribution.employerContribution || 0) / 100;

 
    var otPay = calcOT(parseFloat(data.ot),parseFloat(data.restOT),parseFloat(data.pubHoliday),employeeName);
   
    var totalDeduction=epfEmployee+socsoEmployee+parseFloat(data.deduction);
    var grossPay=basicSalary+parseFloat(otPay)+parseFloat(data.claim)+parseFloat(getEmployeeData(8,employeeName));
    var netPay=parseFloat(grossPay)-parseFloat(totalDeduction);



    var newRow = [
      generateID(getRawData("1RplPsRWxIPWdKPJ-pyzkNUPYfzqkWJu00Fla1Q1XY_E", months[new Date().getMonth()]),0,"FWP"),
      getEmployeeData(0,employeeName),
      getEmployeeData(1,employeeName),
      getEmployeeData(3,employeeName), 
      getEmployeeData(6,employeeName),
      "Permanent",
      getEmployeeData(7,employeeName),
      getEmployeeData(9,employeeName),
      getEmployeeData(10,employeeName),
      getEmployeeData(12,employeeName),
      getEmployeeData(13,employeeName),
      formatDate(new Date()),
      data.dayWork,
      data.ot,
      data.restOT,
      data.pubHoliday,
      parseFloat(otPay).toFixed(2),/*overtime pay*/ 
      data.medicalLeave,
      data.annualLeave,
      epfEmployee,
      epfEmployer,
      socsoEmployee,
      socsoEmployer,
      data.claim,
      getEmployeeData(8,employeeName),
      data.deduction,
      data.remarks,
      parseFloat(grossPay).toFixed(2),
      parseFloat(totalDeduction).toFixed(2),
      parseFloat(netPay).toFixed(2)
    ];

    // Append the same row to the other spreadsheet
    anotherSheet.appendRow(newRow);

    const employee = {
      ic: getEmployeeData(3, employeeName),
      bank: getEmployeeData(12, employeeName),
      acc: getEmployeeData(13, employeeName),
      position: getEmployeeData(6, employeeName),
      employeeID: getEmployeeData(0, employeeName),
      employeeName: employeeName,
      basicSalary: basicSalary,
      overtimePay: parseFloat(otPay),
      allowances: parseFloat(getEmployeeData(8, employeeName)),
      epfContributionEmployee: epfEmployee,
      socsoContributionEmployee: socsoEmployee,
      otherDeduction:parseFloat(data.deduction),
      grossPay: parseFloat(grossPay),
      deductions: parseFloat(totalDeduction),
      netPay: parseFloat(netPay),
      issuedDate: formatDate(new Date()), // Add issued date to template context
      socsoNo: getEmployeeData(10, employeeName),
      epfNo: getEmployeeData(9, employeeName),
      email: getEmployeeData(11, employeeName),
      payslipMonth: currentMonth
    };

    // Send payslip PDF
    sendPayslipPdf(employee);


    // Log the success message
    Logger.log('Data appended successfully.');

    // Return a response to the client
    return ContentService.createTextOutput('Data successfully added!');
  } catch (error) {
    Logger.log('Error occurred: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
    return ContentService.createTextOutput('An error occurred: ' + error.message);
  }
}

function sendPayslipPdf(employee) {
  const htmlTemplate = HtmlService.createTemplateFromFile('payslipTemplate');
  htmlTemplate.ic = employee.ic;
  htmlTemplate.bank = employee.bank;
  htmlTemplate.acc = employee.acc;
  htmlTemplate.position = employee.position;
  htmlTemplate.employeeID = employee.employeeID;
  htmlTemplate.employeeName = employee.employeeName;
  htmlTemplate.basicSalary = employee.basicSalary.toFixed(2);
  htmlTemplate.overtimePay = employee.overtimePay.toFixed(2);
  htmlTemplate.allowances = employee.allowances.toFixed(2);
  htmlTemplate.epfContributionEmployee = employee.epfContributionEmployee.toFixed(2);
  htmlTemplate.socsoContributionEmployee = employee.socsoContributionEmployee.toFixed(2);
  htmlTemplate.otherDeduction=employee.otherDeduction.toFixed(2);
  htmlTemplate.grossPay = employee.grossPay.toFixed(2);
  htmlTemplate.deductions = employee.deductions.toFixed(2);
  htmlTemplate.netPay = employee.netPay.toFixed(2);
  htmlTemplate.issuedDate = employee.issuedDate; 
  htmlTemplate.socsoNo=employee.socsoNo;
  htmlTemplate.epfNo=employee.epfNo;
  htmlTemplate.authorizedName="name";
  htmlTemplate.payslipMonth=employee.payslipMonth;
  

  const htmlContent = htmlTemplate.evaluate().getContent();
  const blob = Utilities.newBlob(htmlContent, 'text/html').getAs('application/pdf');

  // Create a file name based on the employee's name and the date
  const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const fileName = `${employee.employeeName}_${formattedDate}.pdf`;
  blob.setName(fileName);

  // Upload the PDF to Google Drive
  const folderId = '15sLjY9kOM-MWa66OSmfvwrnvB9g-_tOU'; 
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);
  const fileUrl = file.getUrl();

  // Copy the file URL into the spreadsheet
  const spreadsheetId = '1RplPsRWxIPWdKPJ-pyzkNUPYfzqkWJu00Fla1Q1XY_E'; 
  const sheetName = months[new Date().getMonth()];; 
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

  // Find the row with the employee name
  const data = sheet.getDataRange().getValues();
  let rowToUpdate = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][2] === employee.employeeName) { // Assuming employee names are in the first column
      rowToUpdate = i + 1; // Row numbers are 1-based
      break;
    }
  }

  if (rowToUpdate > 0) {
    // Find the last column with data and set the file URL in the next column
    const lastColumn = sheet.getLastColumn();
    sheet.getRange(rowToUpdate, lastColumn ).setValue(fileUrl);
  } else {
    Logger.log('Employee not found in the spreadsheet.');
  }

  const emailOptions = {
    attachments: [blob],
    name: 'Finance Wizards'
  };

  MailApp.sendEmail(employee.email, `Payslip for ${employee.payslipMonth}`, 'Please find your payslip attached.', emailOptions);
}





/* Profile */

function getCompanyInfo(e){
  var employeeListSheets = SpreadsheetApp.openById("1dibgXEjQK9J4Bi1WGXnbh_c1SjXLJw6-OPRdPupRguo");
  var employeeListInfo = employeeListSheets.getSheetByName("Company");

  if(e == 0)
    if(employeeListInfo.getRange("B1").getValue() == ""){
      return "No data to show";
    }
    else{
      return employeeListInfo.getRange("B1").getValue();
    }
  else if(e == 1){
    if(employeeListInfo.getRange("B2").getValue() == ""){
      return "No data to show";
    }
    else{
      return employeeListInfo.getRange("B2").getValue();
    }
  }
  else if(e == 2){
    if(employeeListInfo.getRange("B3").getValue() == ""){
      return "No data to show";
    }
    else{
      return employeeListInfo.getRange("B3").getValue();
    }
  }
  else if(e == 3){
    var arr = [];
    for(var i = 4; i < 9 ; i++){
      var a = "B" + i;
      arr.push(employeeListInfo.getRange(a).getValue());
    }
    return arr;
  }
  else if(e == 5){
    var arr = [];
    for(var i = 1; i < 9 ; i++){
      var a = "B" + i;
      arr.push(employeeListInfo.getRange(a).getValue());
    }
    return arr;
  }
  else{
    return "No data to show";
  }
}

function updateCompanyInfo(form){
  var companyInfoSheets = SpreadsheetApp.openById("1dibgXEjQK9J4Bi1WGXnbh_c1SjXLJw6-OPRdPupRguo");
  var companyInfo = companyInfoSheets.getSheetByName("Company");

  companyInfo.getRange("B1").setValue(form.company_name);
  companyInfo.getRange("B2").setValue(form.registration_no);
  companyInfo.getRange("B3").setValue(form.contact);
  companyInfo.getRange("B4").setValue(form.address);
  companyInfo.getRange("B5").setValue(form.postcode);
  companyInfo.getRange("B6").setValue(form.city);
  companyInfo.getRange("B7").setValue(form.province);
  companyInfo.getRange("B8").setValue(form.country);

  return getScriptUrl() + "?v=profile";
}

/* Procurement Plan */

const SPREADSHEET_ID = '1hHZKf1TWIeG8iVHawni7_3B64otV7ifJguuzG2VSLFY';

// get requests data
function getRequests() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProcurementRequests');
  const data = sheet.getDataRange().getValues();
  return data.slice(1); // Exclude header row
}


// submit procurement request
function submitProcurementRequest(formData) {
  var sheet = getSheet(SPREADSHEET_ID, 'ProcurementRequests');
  sheet.appendRow([
    generateUniqueId(),
    formData.name,
    formData.department,
    formData.plan,
    0, // Initial total cost
    'Pending'
  ]);
}

// add procurement item and update total budget cost
function addProcurementItem(itemData) {
  var sheet = getSheet(SPREADSHEET_ID, 'ProcurementItems');
  var totalCost = itemData.quantity * itemData.price;
  sheet.appendRow([
    itemData.id,
    itemData.item,
    itemData.quantity,
    itemData.price,
    totalCost
  ]);

  updateTotalBudgetCost(itemData.id);
}

// update total budget cost in the ProcurementRequests sheet
function updateTotalBudgetCost(id) {
  var itemsData = getSheetData(SPREADSHEET_ID, 'ProcurementItems', 1, 0, id);
  var totalBudgetCost = itemsData.reduce(function(sum, row) {
    return sum + row.Amount; // Total cost column
  }, 0);

  var requestsSheet = getSheet(SPREADSHEET_ID, 'ProcurementRequests');
  var requestsData = getRawData(SPREADSHEET_ID, 'ProcurementRequests');
  for (var i = 1; i < requestsData.length; i++) {
    if (requestsData[i][0] == id) {
      requestsSheet.getRange(i + 1, 5).setValue(totalBudgetCost); // Total Budget Cost column
      break;
    }
  }
}

// generate unique id
function generateUniqueId() {
  return Math.floor(Date.now() / 1000);
}

// generate procurement details table html
function generateProcurementDetailsTable(id) {
  var itemsData = getSheetData(SPREADSHEET_ID, 'ProcurementItems', 1, 0, id);
  var html = '';
  for (var i = 0; i < itemsData.length; i++) {
    html += '<tr>';
    for (var j = 1; j < itemsData[i].length; j++) {
      html += '<td>' + itemsData[i][j] + '</td>';
    }
    html += '</tr>';
  }
  return html;
}

function getInventoryItems() {
  const sheet = SpreadsheetApp.openById('1ns-x6_wP_Ibxr1Aab7O9r3K-CvKUcL2Jk1pX8DCDufM').getSheetByName('Product');
  const data = sheet.getDataRange().getValues();
  const items = data.slice(1).map(row => ({
    name: row[1],
    productID: row[0]
  }));
  return items;
}

function storeProcurementPlan(procurementPlan) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ProcurementRequests');
  const itemsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ProcurementItems');
  const requestId = generateUniqueId();

  sheet.appendRow([
    requestId,
    procurementPlan.name,
    procurementPlan.department,
    procurementPlan.plan,
    procurementPlan.urgency,
    procurementPlan.items.reduce((sum, item) => sum + parseFloat(item.totalCost), 0),
    'Pending' // Set status as 'Pending'
  ]);

  procurementPlan.items.forEach(item => {
    itemsSheet.appendRow([
      requestId,
      item.item,
      item.quantity,
      item.unitPrice,
      item.totalCost
    ]);
  });
}

function generateProcurementPlanTable() {
  const sheet = SpreadsheetApp.openById('1hHZKf1TWIeG8iVHawni7_3B64otV7ifJguuzG2VSLFY').getSheetByName('ProcurementRequests');
  const data = sheet.getDataRange().getValues();
  
  // Start constructing the HTML table
  let html = '';
  
  // Iterate over rows of the sheet (skip header row)
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'Pending') { // Only fetch pending requests
      html += '<tr>';
      for (let j = 0; j < data[i].length; j++) {
        if (j === 0) { // Procurement Description column
          html += `<td><a href="${ScriptApp.getService().getUrl()}?v=procurementdetails&id=${data[i][0]}">${data[i][j]}</a></td>`;
        } else {
          html += `<td>${data[i][j]}</td>`;
        }
      }
      html += '</tr>';
    }
  }
  
  return html;
}

function generateProcurementDetailsTable(requestId) {
  const sheet = SpreadsheetApp.openById('1hHZKf1TWIeG8iVHawni7_3B64otV7ifJguuzG2VSLFY').getSheetByName('ProcurementRequests');
  const itemsSheet = SpreadsheetApp.openById('1hHZKf1TWIeG8iVHawni7_3B64otV7ifJguuzG2VSLFY').getSheetByName('ProcurementItems');

  // Fetch the request details
  const requestsData = sheet.getDataRange().getValues();
  let requestDetails = {};
  
  for (let i = 1; i < requestsData.length; i++) {
    if (requestsData[i][0] == requestId) {
      requestDetails = {
        requestId: requestsData[i][0],
        name: requestsData[i][1],
        department: requestsData[i][2],
        plan: requestsData[i][3],
        urgency: requestsData[i][4],
        status: requestsData[i][6], // Status column
        items: []
      };
      break;
    }
  }
  
  // Fetch the items for this request
  const itemsData = itemsSheet.getDataRange().getValues();
  
  for (let i = 1; i < itemsData.length; i++) {
    if (itemsData[i][0] == requestId) {
      requestDetails.items.push({
        productName: itemsData[i][1],
        unitPrice: itemsData[i][2],
        quantity: itemsData[i][3],
        totalCost: itemsData[i][4]
      });
    }
  }

  // Construct the HTML for the procurement details page
  let html = `
    <!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
    }

    body {
      display: flex;
      flex-direction: column;
    }

    .logo {
      font-size: 30px;
      font-weight: bold;
      padding: 0 20px;
    }

    nav {
      height: 80px;
      background: white;
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 0 20px;
    }

    nav a {
      text-decoration: none;
      color: black;
      padding: 0 20px;
      font-size: 20px;
    }

    .home-info {
      padding: 20px 40px;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    table {
      margin: 40px auto 20px auto;
      border-collapse: collapse;
      width: 100%;
      max-width: 1000px;
    }

    td, th {
      border: 1px solid #dddddd;
      padding: 10px 0;
    }

    td {
      text-align: center;
    }

    label{
      min-width: 80px;
      margin-top: 10px;
      margin-bottom: 10px;
      display: inline-block;
    }

    .field input{
      font-size: 16px;
      line-height: 28px;
      padding: 8px 16px;
      width: 100%;
      min-height: 44px;
      border: unset;
      border-radius: 4px;
      outline-color: rgb(254 208 63 / 0.5);
      background-color: rgb(255, 255, 255);
      box-shadow: rgba(0, 0, 0, 0) 0px 0px 0px 0px, 
        rgba(0, 0, 0, 0) 0px 0px 0px 0px, 
        rgba(0, 0, 0, 0) 0px 0px 0px 0px, 
        rgba(60, 66, 87, 0.16) 0px 0px 0px 1px, 
        rgba(0, 0, 0, 0) 0px 0px 0px 0px, 
        rgba(0, 0, 0, 0) 0px 0px 0px 0px, 
        rgba(0, 0, 0, 0) 0px 0px 0px 0px;
    }

    .footer {
      margin-top: auto;
      height: 100px;
      justify-content: center;
      text-align: center;
      width: 100%;
      background-color: #fed03f;
      bottom: 0;
      padding: 30px;
    }

    .button {
      padding: 8px 20px;
      border-color: transparent;
      border-radius: 5px;
      background: #fed03f;
      font-size: 15px;
      cursor: pointer;
      text-decoration: none;
      color: black;
    }

    .button:hover {
      background: #fed03f60;
    }

    .button:active {
      background: #fed03f90;
    }

    .modal {
      display: none;
      position: fixed;
      z-index: 1;
      left: 0;
      top: 0;
      width: 100%;
      height: 100%;
      overflow: auto;
      background-color: rgb(0, 0, 0);
      background-color: rgba(0, 0, 0, 0.4);
      padding-top: 60px;
    }

    .modal-content {
      background-color: #fefefe;
      margin: 5% auto;
      padding: 20px;
      border: 1px solid #888;
      width: 80%;
      max-width: 600px;
    }

    .close {
      color: #aaa;
      float: right;
      font-size: 28px;
      font-weight: bold;
    }

    .close:hover,
    .close:focus {
      color: black;
      text-decoration: none;
      cursor: pointer;
    }
  </style>
</head>
<body style="justify-content: center; padding: 0; min-height: 100vh;">
  <nav>
    <div class="logo">Finance Wizards</div>
    <div class="menu">
        <a href="${getScriptUrl()}?v=index">Home</a>
        <a href="${getScriptUrl()}?v=pp">Procurement</a>
        <a href="${getScriptUrl()}?v=inventory">Inventory</a>
        <a href="${getScriptUrl()}?v=invoices">Invoices</a>
        <a href="${getScriptUrl()}?v=payroll">Payroll</a>
        <a href="${getScriptUrl()}?v=profile">Profile</a>
    </div>
  </nav>
  <div class="home-info" style="background-color: antiquewhite;">
    <div>
        <h2 id="datetime">Today is N/A</h2>
        <p id="dayleft">N/A more days till Payday</p>
    </div>
  </div>
  <h2 style="margin: 50px 50px 0 50px;">Procurement Details</h2>
  <div id="details" style="margin: 50px;">
    <p><strong>Requestor Name:</strong> ${requestDetails.name}</p>
    <br>
    <p><strong>Department:</strong> ${requestDetails.department}</p>
    <br>
    <p><strong>Procurement Plan:</strong> ${requestDetails.plan}</p>
    <br>
    <p><strong>Urgency:</strong> ${requestDetails.urgency}</p>
    <br>
    <p><strong>Status:</strong> ${requestDetails.status}</p>
  </div>
  <table border="1" style="margin: 50px;">
    <thead>
      <tr>
        <th>Product Name</th>
        <th>Unit Price</th>
        <th>Quantity</th>
        <th>Total Cost</th>
      </tr>
    </thead>
    <tbody>
`;

requestDetails.items.forEach(function(item) {
  html += `
    <tr>
      <td>${item.productName}</td>
      <td>${item.unitPrice}</td>
      <td>${item.quantity}</td>
      <td>${item.totalCost}</td>
    </tr>
  `;
});

let totalAmount = requestDetails.items.reduce((sum, item) => sum + parseFloat(item.totalCost), 0).toFixed(2);

html += `
    </tbody>
    <tfoot>
      <tr>
        <td colspan="3" style="text-align: center;">Total amount (RM)</td>
        <td>${totalAmount}</td>
      </tr>
    </tfoot>
  </table>

  <div style="text-align: center; margin: 20px;">
    <button class="button" id="approveButton">Approve</button>
    <button class="button" id="rejectButton">Reject</button>
  </div>

  <div id="approvalModal" class="modal">
    <div class="modal-content">
      <span class="close">&times;</span>

      <h2 style="margin: 50px 0px 0px 0px;">Approval Details</h2>
        <div class="field" style="margin: 20px 0px 0 0px;padding-bottom: 24px; display:inline-block; width: 100%;">
          <div style="width:100%;">
            <label for="approverName">Approver Name:</label>
            <input type="text" id="approverName" name="approverName">
          </div>
          <div style="width:100px"></div>
          <div style="width:100%;">
            <label for="comments">Comments:</label>
            <input type="comments" id="comments" name="comments">
          </div>
        </div>
      <br><br>
      <button class="button" id="confirmButton">Confirm</button>
    </div>
  </div>

  <script>
  const modal = document.getElementById("approvalModal");
  const span = document.getElementsByClassName("close")[0];

  span.onclick = function() {
    modal.style.display = "none";
  }

  window.onclick = function(event) {
    if (event.target == modal) {
      modal.style.display = "none";
    }
  }

  document.getElementById("approveButton").onclick = function() {
    showModal('Approved');
  }

  document.getElementById("rejectButton").onclick = function() {
    showModal('Rejected');
  }

  function showModal(status) {
    modal.style.display = "block";
    document.getElementById("confirmButton").onclick = function() {
      const approverName = document.getElementById("approverName").value;
      const comments = document.getElementById("comments").value;
      const approvalDetails = {
        requestId: '${requestId}',
        approverName: approverName,
        approvalDate: new Date().toLocaleDateString('en-GB'),
        status: status,
        comments: comments
      };
      google.script.run.withSuccessHandler(function() {
        window.location.href = '${getScriptUrl()}?v=pp';
      }).saveApprovalDetails(approvalDetails);
      modal.style.display = "none";
    };
  }
</script>
</body>
</html>
`;
  return html;
}

function saveApprovalDetails(approvalDetails) {
  const sheet = SpreadsheetApp.openById('1hHZKf1TWIeG8iVHawni7_3B64otV7ifJguuzG2VSLFY').getSheetByName('Approvals');
  
  // Append the new approval details
  sheet.appendRow([
    approvalDetails.requestId,
    approvalDetails.approverName,
    approvalDetails.approvalDate,
    approvalDetails.status,
    approvalDetails.comments,
    approvalDetails.status === 'Approved' ? 'Pending' : '',
    '', // Skip one column
    approvalDetails.status === 'Approved' ? '0' : ''
  ]);

  // Update the status in 'ProcurementRequests'
  updateProcurementRequestStatus(approvalDetails.requestId, approvalDetails.status);
  
  return true;
}

function updateProcurementRequestStatus(requestId, status) {
  const sheet = SpreadsheetApp.openById('1hHZKf1TWIeG8iVHawni7_3B64otV7ifJguuzG2VSLFY').getSheetByName('ProcurementRequests');
  const requestsData = sheet.getDataRange().getValues();

  for (let i = 1; i < requestsData.length; i++) {
    if (requestsData[i][0] == requestId) {
      sheet.getRange(i + 1, 7).setValue(status); // Update the status column
      break;
    }
  }
}

function generateApprovedTable() {
  try {
    // get data from google sheets
    let jsonData = getSheetData("1hHZKf1TWIeG8iVHawni7_3B64otV7ifJguuzG2VSLFY", "Approvals");

    // create table
    let tags = "";
    if (jsonData.length == 0) {
      return getNoDataHtml(7);
    } else {
      for (let d of jsonData) {
        if (d['Status'] === 'Approved') {
          tags += `<tr>
            <td>${d['Approval ID']}</td>
            <td>${d['Approver Name']}</td>
            <td>${getFormattedDate(d['Approval Date'])}</td>
            <td>${d['Comments']}</td>
            <td>${d['Purchase Status']}</td>
            <td>${d['Purchase Date'] ? getFormattedDate(d['Purchase Date']) : '-'}</td>`;
          
          if (d['Action'] === 1) {
            tags += `
            <td>
              <button onclick="nextPurchaseOrder(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer">
                <img src="https://drive.google.com/thumbnail?id=15oz_9RrbWCvE_gjlS_K_ohyFlVAMhMXY" title="See Details" style="height:25px;">
              </button>
            </td>`;
          } else { 
            tags += `
            <td>
              <button onclick="nextPurchaseOrder(this)" style="margin: 0 8px; background-color: transparent; border:none; cursor: pointer">
                <img src="https://drive.google.com/thumbnail?id=1XJR-A_HECzXF0WcCL7Nu7qJWZ2UjRgYA" title="Purchase" style="height:25px;">
              </button>
            </td>`;
          }
          tags += `</tr>`;
        }
      }
    }
    return tags;
  } catch (error) {
    console.error('Error generating approved table:', error);
    return getNoDataHtml(7);
  }
}

function getSheetData(spreadsheetId, sheetName) {
  try {
    // Mock implementation to get data from Google Sheets, replace with actual implementation
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    
    // Convert data to JSON format
    const headers = data.shift();
    return data.map(row => {
      let record = {};
      headers.forEach((header, index) => {
        record[header] = row[index];
      });
      return record;
    });
  } catch (error) {
    console.error('Error fetching sheet data:', error);
    return [];
  }
}

function getFormattedDate(dateString) {
  try {
    const date = new Date(dateString);
    const options = { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' };
    return date.toLocaleDateString('en-GB', options);
  } catch (error) {
    console.error('Error formatting date:', error);
    return dateString;
  }
}
function generateRejectedTable() {
  try {
    // get data from google sheets
    let jsonData = getSheetData("1hHZKf1TWIeG8iVHawni7_3B64otV7ifJguuzG2VSLFY", "Approvals");

    // create table
    let tags = "";
    if (jsonData.length == 0) {
      return getNoDataHtml(4);
    } else {
      for (let d of jsonData) {
        if (d['Status'] === 'Rejected') {
          tags += `<tr>
            <td>${d['Approval ID']}</td>
            <td>${d['Approver Name']}</td>
            <td>${getFormattedDate(d['Approval Date'])}</td>
            <td>${d['Comments']}</td>
          </tr>`;
        }
      }
    }
    return tags;
  } catch (error) {
    console.error('Error generating rejected table:', error);
    return getNoDataHtml(4);
  }
}

function getSupplierOptions() {
  var sheet = SpreadsheetApp.openById('1ns-x6_wP_Ibxr1Aab7O9r3K-CvKUcL2Jk1pX8DCDufM').getSheetByName('Supplier');
  var data = sheet.getRange('A2:B').getValues(); // Adjust range according to your data
  return data.map(function(row) {
    return '<option value="' + row[0] + '">' + row[1] + '</option>'; // Adjust according to your columns
  }).join('');
}

// Function to get Delivery Type Options
function getDTypeOptions() {
  var sheet = SpreadsheetApp.openById('10JUBz_zmt1aDG_Sl7KfCQ7K08yqAfTjzYyRAQDLtuWc').getSheetByName('Delivery Method');
  var data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
  return data.map(function(row) {
    var type = row[0];
    return type ? '<option value="' + type + '">' + type + '</option>' : ''; // Handle empty cells
  }).join('');
}

function getDMethodOptions() {
  var sheet = SpreadsheetApp.openById('10JUBz_zmt1aDG_Sl7KfCQ7K08yqAfTjzYyRAQDLtuWc').getSheetByName('Delivery Method');
  var data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
  return data.map(function(row) {
    var method = row[0];
    return '<option value="' + method + '">' + method + '</option>';
  }).join('');
}

function getPTypeOptions() {
  var sheet = SpreadsheetApp.openById('10JUBz_zmt1aDG_Sl7KfCQ7K08yqAfTjzYyRAQDLtuWc').getSheetByName('Payment Method');
  var data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
  return data.map(function(row) {
    var type = row[0];
    return type ? '<option value="' + type + '">' + type + '</option>' : ''; // Handle empty cells
  }).join('');
}

function getPMethodOptions() {
  var sheet = SpreadsheetApp.openById('10JUBz_zmt1aDG_Sl7KfCQ7K08yqAfTjzYyRAQDLtuWc').getSheetByName('Payment Method');
  var data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
  return data.map(function(row) {
    var method = row[0];
    return '<option value="' + method + '">' + method + '</option>';
  }).join('');
}

function getSuppliersInfo(supplierName) {
  Logger.log('Supplier name: ' + supplierName);
  var sheet = SpreadsheetApp.openById('1ns-x6_wP_Ibxr1Aab7O9r3K-CvKUcL2Jk1pX8DCDufM').getSheetByName('Supplier');
  var data = sheet.getRange('A2:M' + sheet.getLastRow()).getValues(); // Adjust range to include all columns and rows with data
  Logger.log('Data: ' + JSON.stringify(data));

  var info = data.find(function(row) { return row[1] === supplierName; });
  Logger.log('Found supplier info: ' + JSON.stringify(info));
  if (info) {
    return { email: info[9], contact: info[8] }; 
  } else {
    return { email: '', contact: '' };
  }
}

function getCompanyAddress() {
  var sheet = SpreadsheetApp.openById('1dibgXEjQK9J4Bi1WGXnbh_c1SjXLJw6-OPRdPupRguo').getSheetByName('Company');
  var data = sheet.getRange('A2:B8').getValues(); // Adjust the range to cover all relevant rows

  var addressData = {};
  
  // Loop through each row and assign the value to the corresponding key in the addressData object
  data.forEach(function(row) {
    switch(row[0]) {
      case 'Company Address 1':
        addressData.address = row[1];
        break;
      case 'City':
        addressData.city = row[1];
        break;
      case 'Postcode':
        addressData.postcode = row[1];
        break;
      case 'Province':
        addressData.province = row[1];
        break;
      case 'Country':
        addressData.country = row[1];
        break;
    }
  });

  return addressData;
}

function getApprovedDetails(approvedId) {
  const itemsSheet = SpreadsheetApp.openById('1hHZKf1TWIeG8iVHawni7_3B64otV7ifJguuzG2VSLFY').getSheetByName('ProcurementItems');
  const itemsData = itemsSheet.getDataRange().getValues();
  let approvedDetails = { items: [] };

  for (let i = 1; i < itemsData.length; i++) {
    if (itemsData[i][0] == approvedId) {
      approvedDetails.items.push({
        productName: itemsData[i][1],
        unitPrice: itemsData[i][2],
        quantity: itemsData[i][3],
        totalCost: itemsData[i][4]
      });
    }
  }
  return approvedDetails;
}

function generatePurchaseOrderTable() {
  try {
    const spreadsheetId = "1ns-x6_wP_Ibxr1Aab7O9r3K-CvKUcL2Jk1pX8DCDufM";
    const sheetName = "purchaseOrder";
    let jsonData = getSheetData(spreadsheetId, sheetName);

    if (jsonData.length === 0) {
      return getNoDataHtml(7); // Assuming 7 columns for the no data message
    } else {
      let tags = "";
      jsonData.forEach(row => {
        tags += `<tr>
                  <td>${row["PO ID"]}</td>
                  <td>${getFormattedDate(row["Purchase Date"])}</td>
                  <td>${row["Supplier ID"]}</td>
                  <td>${row["Total Payment"]}</td>
                  <td>${row["Payment Status"]}</td>
                  <td>${row["Delivery Status"]}</td>
                  <td>${getFormattedDate(row["Expected Date Received"])}</td>
                 </tr>`;
      });
      return tags;
    }
  } catch (error) {
    console.error("Error generating purchase order table: ", error);
    return "<tr><td colspan='7'>Error generating table.</td></tr>";
  }
}