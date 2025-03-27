// Function to submit the data to Database sheet

function submitData() {

  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 

  var shUserForm= myGooglSheet.getSheetByName("فاتورة اوردر"); //delcare a variable and set with the User Form worksheet
  var datasheet1 = myGooglSheet.getSheetByName("مبيعات"); ////delcare a variable and set with the Database worksheet
  var datasheet2 = myGooglSheet.getSheetByName("تفاصيل الاوردر");
  var datasheet3 = myGooglSheet.getSheetByName("بوليصة");
  var repo = myGooglSheet.getSheetByName("المخزن");


  

  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  var response = ui.alert("Submit", 'Do you want to submit the data?',ui.ButtonSet.YES_NO);
  // Checking the user response and proceed with clearing the form if user selects Yes
  if (response == ui.Button.NO) {
    return;//exit from this function
  } 

  //Validating the entry. If validation is true then proceed with transferring the data to Database sheet

 if (validateEntry(shUserForm)==true && validateRepo(shUserForm)==true) {

//      For مبيعات

    var lastRow1=datasheet1.getLastRow()+1; //identify the next blank row
    var newValues1 = 
    [
      [lastRow1, new Date().toJSON().slice(0, 10), shUserForm.getRange("E6").getValue(), 
      shUserForm.getRange("E9").getValue(), shUserForm.getRange("E12").getValue(), 
      shUserForm.getRange("H12").getValue(), shUserForm.getRange("E14").getValue(), 
      shUserForm.getRange("H14").getValue(), shUserForm.getRange("I26").getValue(), 
      shUserForm.getRange("I27").getValue(), shUserForm.getRange("E16").getValue(), 
      shUserForm.getRange("E18").getValue(), shUserForm.getRange("H16").getValue(), 
      'لم يتم التحضير']
    ];
    datasheet1.getRange(lastRow1, 1, 1, newValues1[0].length).setValues(newValues1);
    datasheet1.getRange(lastRow1, 2).setValue(new Date().toJSON().slice(0, 10)).setNumberFormat('yyyy-mm-dd'); //Date



//      For تفاصيل الاوردر


    // Get the range from D21 to H25 in the source sheet
  var dataRange = shUserForm.getRange('D21:I25');
  var originalValues = dataRange.getValues();
  // Get the target range in the target sheet
  var lastRow = datasheet2.getLastRow();
  Logger.log("lastRow: " + lastRow);
  var targetRange = datasheet2.getRange(lastRow + 1,4 , 5, 6);
  // Set values in the target range
  targetRange.setValues(originalValues);
  var shippingInTarget = datasheet2.getRange(lastRow + 1,12);
  shippingInTarget.setValue(newValues1[0][9]);
  var seller = datasheet2.getRange(lastRow + 1,14,5,1);
  seller.setValue(newValues1[0][7]);

  // Get the orderNumber, Date and pageName and paste it in all rows
  //var orderNumber = shUserForm.getRange('H16').getValue();
  //var pageName = shUserForm.getRange('E16').getValue();
  targetRange.offset(0,-1,5,1).setValue(newValues1[0][12]);
  targetRange.offset(0,-2,5,1).setValue(new Date().toJSON().slice(0, 10)).setNumberFormat('yyyy-mm-dd');
  targetRange.offset(0,-3,5,1).setValue(newValues1[0][10]);


//      For المخزن

  var keyColumn = repo.getRange("A:A").getValues();

  for (var i = 0; i < originalValues.length; i++) {
    if(originalValues[i][0] !==""){
      var product = originalValues[i][0];
      var rowIndex = keyColumn.findIndex(row => row[0] === product);
      var value = repo.getRange(rowIndex+1, 2).getValue();
      repo.getRange(rowIndex+1, 2).setValue(value - originalValues[i][4]);
    }

  }


//      For بوليصة
/*

  // Get the range to be copied (A2:J12)
  var rangeToCopy = datasheet3.getRange('A2:K13');

  // Find the last row in the target sheet
  var lastRow2 = datasheet3.getLastRow();
  Logger.log("lastRow2: " + lastRow2);


  // Increment the value in the top A cell of the pasted data
  var topAValue = Number(datasheet3.getRange(lastRow2 , 8).getValue()) + 1;

  // Get the range to paste data (below existing data)
  var rangeToPaste = datasheet3.getRange(lastRow2 + 1, 1);

  // Copy and paste the data with the same formulas and format
  rangeToCopy.copyTo(rangeToPaste, { contentsOnly: false, formatOnly: false });

  // Set the value in the top A cell of the pasted data
  datasheet3.getRange(lastRow2 + 1, 1, 12, 1).setValue(topAValue);

*/


  // Show message after complete
  ui.alert("New Data Saved");
  

  //Clearnign the data from the Data Entry Form

    shUserForm.getRange("E6:E18").clearContent();
    shUserForm.getRange("H12:H16").clearContent();
    shUserForm.getRange("I27").clearContent();
    shUserForm.getRange("D21:H25").clearContent();

 }

  


}

function submitData1() {
  var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shUserForm = myGoogleSheet.getSheetByName("فاتورة اوردر");
  var datasheet1 = myGoogleSheet.getSheetByName("مبيعات");
  var datasheet2 = myGoogleSheet.getSheetByName("تفاصيل الاوردر");
  var datasheet3 = myGoogleSheet.getSheetByName("بوليصة");
  var repo = myGoogleSheet.getSheetByName("المخزن");

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Submit", 'Do you want to submit the data?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) {
    return;
  }

  if (validateEntry(shUserForm) && validateRepo(shUserForm)) {
    var lastRow1 = datasheet1.getLastRow() + 1;
    var newValues1 = [
      [lastRow1, new Date().toJSON().slice(0, 10), shUserForm.getRange("E6").getValue(), 
      shUserForm.getRange("E9").getValue(), shUserForm.getRange("E12").getValue(), 
      shUserForm.getRange("H12").getValue(), shUserForm.getRange("E14").getValue(), 
      shUserForm.getRange("H14").getValue(), shUserForm.getRange("I26").getValue(), 
      shUserForm.getRange("I27").getValue(), shUserForm.getRange("E16").getValue(), 
      shUserForm.getRange("E18").getValue(), shUserForm.getRange("H16").getValue(), 
      'لم يتم التحضير', shUserForm.getRange("E18").getValue()]
    ];
    datasheet1.getRange(lastRow1, 1, 1, newValues1[0].length).setValues(newValues1);

    var dataRange = shUserForm.getRange('D21:I25');
    var originalValues = dataRange.getValues();
    var lastRow2 = datasheet2.getLastRow() + 1;
    var targetRange2 = datasheet2.getRange(lastRow2, 4, originalValues.length, originalValues[0].length);
    targetRange2.setValues(originalValues);
    targetRange2.offset(0, 8).setValue(shUserForm.getRange("I27").getValue());
    targetRange2.offset(0, 10).setValue(shUserForm.getRange("H14").getValue());
    targetRange2.offset(0, -3).setValue(shUserForm.getRange('H16').getValue());
    targetRange2.offset(0, -2).setValue(new Date().toJSON().slice(0, 10));
    targetRange2.offset(0, -1).setValue(shUserForm.getRange('E16').getValue());

    var keyColumn = repo.getRange("A:A").getValues();
    for (var i = 0; i < originalValues.length; i++) {
      if (originalValues[i][0] !== "") {
        var product = originalValues[i][0];
        var rowIndex = keyColumn.findIndex(row => row[0] === product);
        var value = repo.getRange(rowIndex + 1, 2).getValue();
        repo.getRange(rowIndex + 1, 2).setValue(value - originalValues[i][4]);
      }
    }

    var rangeToCopy = datasheet3.getRange('A2:K13');
    var lastRow3 = datasheet3.getLastRow();
    var topAValue = Number(datasheet3.getRange(lastRow3, 8).getValue()) + 1;
    var rangeToPaste = datasheet3.getRange(lastRow3 + 1, 1, rangeToCopy.getNumRows(), rangeToCopy.getNumColumns());
    rangeToCopy.copyTo(rangeToPaste, { contentsOnly: false, formatOnly: false });
    datasheet3.getRange(lastRow3 + 1, 1, rangeToCopy.getNumRows(), 1).setValue(topAValue);

    ui.alert("New Data Saved");

    shUserForm.getRange("E6:E18").clearContent();
    shUserForm.getRange("H12:H16").clearContent();
    shUserForm.getRange("I27").clearContent();
    shUserForm.getRange("D21:H25").clearContent();
  }
}


function fetchDataBasedOnOrderNum() {
  var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shShippingStatus = myGoogleSheet.getSheetByName("حالة الشحن2");
  var datasheet1 = myGoogleSheet.getSheetByName("مبيعات");
  var detailsSheet = myGoogleSheet.getSheetByName("تفاصيل الاوردر");
  var ui = SpreadsheetApp.getUi();


  // Get the key from cell H2 where the order number
  var keyToSearch = shShippingStatus.getRange("H1").getValue().trim();
  if(shShippingStatus.getRange("H1").isBlank()==true){
    ui.alert("من فضلك ادخل رقم الاوردر");
  }
  else{

    // Search for the key in column H (assuming column H contains the keys in مبيعات sheet)
    var keyColumn = datasheet1.getRange("M:M").getValues();
    var rowIndex = keyColumn.findIndex(row => row[0] === keyToSearch);
    // Search for the key in column C of تفاصيل الاوردر
    var keyColumn2 = detailsSheet.getRange("C:C").getValues();
    var rowIndex2 = keyColumn2.findIndex(row => row[0] === keyToSearch);

    Logger.log("rowIndex: " + rowIndex);
    Logger.log("rowIndex2: " + rowIndex2);

    if (rowIndex !== -1 && rowIndex2 !== -1) {
      // Key found, retrieve data and populate the shipping status form
      var dataRow = datasheet1.getRange(rowIndex + 1, 1, 1, datasheet1.getLastColumn()).getValues()[0];

      // Assuming the shipping status form has the same layout as the مبيعات sheet
      shShippingStatus.getRange("E6").setValue(dataRow[2]);  // Update the cell references based on your actual layout
      //shShippingStatus.getRange("E9").setValue(dataRow[3]);
      //shShippingStatus.getRange("E12").setValue(dataRow[4]);
      //shShippingStatus.getRange("H12").setValue(dataRow[5]);
      //shShippingStatus.getRange("E14").setValue(dataRow[6]);
      //shShippingStatus.getRange("E16").setValue(dataRow[10]);
      //shShippingStatus.getRange("H26").setValue(dataRow[8]);
      shShippingStatus.getRange("H16").setValue(dataRow[12]);
      shShippingStatus.getRange("I27").setValue(dataRow[9]);
      //shShippingStatus.getRange("H14").setValue(dataRow[7]);
      //shShippingStatus.getRange("E16").setValue(dataRow[10]);
      //shShippingStatus.getRange("E18").setValue(dataRow[13]);

      var dataRange = detailsSheet.getRange(rowIndex2 + 1, 4, 5, 8); // Assuming columns D to I
      var dataValues = dataRange.getValues();

      // Set the values in حالة الشحن2 starting from D21 to I25
      var targetRange = shShippingStatus.getRange("D21:K25");
      targetRange.setValues(dataValues);

      ui.alert("Data retrieved successfully!");

    } else {
      // Key not found, display a message or take appropriate action
      ui.alert("رقم الاوردر غير موجود");
    }

  }
}


function fetchDataBasedOnPhoneNum() {
  var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shShippingStatus = myGoogleSheet.getSheetByName("حالة الشحن2");
  var datasheet1 = myGoogleSheet.getSheetByName("مبيعات");
  var detailsSheet = myGoogleSheet.getSheetByName("تفاصيل الاوردر");
  var ui = SpreadsheetApp.getUi();


  // Get the key from cell H2 where the order number
  var phoneNum = shShippingStatus.getRange("E1").getValue().trim();
  if(shShippingStatus.getRange("E1").isBlank()==true){
    ui.alert("من فضلك ادخل رقم التليفون");
  }
  else{


    // Search for the key in column H (assuming column H contains the keys in مبيعات sheet)
    var keyColumn = datasheet1.getRange("E:F").getValues();
    var rowIndex = keyColumn.findIndex(row => row[0] === phoneNum);
    // Search for the key in column C of تفاصيل الاوردر
    

    if (rowIndex !== -1 ) {
      // Key found, retrieve data and populate the shipping status form
      var dataRow = datasheet1.getRange(rowIndex + 1, 1, 1, datasheet1.getLastColumn()).getValues()[0];

      // Assuming the shipping status form has the same layout as the مبيعات sheet
      shShippingStatus.getRange("E6").setValue(dataRow[2]);  // Update the cell references based on your actual layout
      //shShippingStatus.getRange("E9").setValue(dataRow[3]);
      //shShippingStatus.getRange("E12").setValue(dataRow[4]);
      //shShippingStatus.getRange("H12").setValue(dataRow[5]);
      //shShippingStatus.getRange("E14").setValue(dataRow[6]);
      //shShippingStatus.getRange("E16").setValue(dataRow[10]);
      //shShippingStatus.getRange("H26").setValue(dataRow[8]);
      shShippingStatus.getRange("H16").setValue(dataRow[12]);
      shShippingStatus.getRange("I27").setValue(dataRow[9]);
      //shShippingStatus.getRange("H14").setValue(dataRow[7]);
      //shShippingStatus.getRange("E16").setValue(dataRow[10]);
      //shShippingStatus.getRange("E18").setValue(dataRow[13]);

    var keyColumn2 = detailsSheet.getRange("C:C").getValues();;
    var rowIndex2 = keyColumn2.findIndex(row => row[0] === dataRow[12]);

    var dataRange = detailsSheet.getRange(rowIndex2 + 1, 4, 5, 8); // Assuming columns D to I
    var dataValues = dataRange.getValues();

    // Set the values in حالة الشحن2 starting from D21 to I25
    var targetRange = shShippingStatus.getRange("D21:K25");
    targetRange.setValues(dataValues);

    SpreadsheetApp.getUi().alert("Data retrieved successfully!");

    } else {
      // Key not found, display a message or take appropriate action
      SpreadsheetApp.getUi().alert("رقم التليفون غير موجود");
    }

  }
}



function saveShippingState() {
  var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shShippingStatus = myGoogleSheet.getSheetByName("حالة الشحن2");
  var detailsSheet = myGoogleSheet.getSheetByName("تفاصيل الاوردر");
  var datasheet1 = myGoogleSheet.getSheetByName("مبيعات");


  // Get the key from cell E3 in the shipping status form
  var keyToSearch = shShippingStatus.getRange("H16").getValue();

  // Search for the key in column H (assuming column H contains the keys in مبيعات sheet)
  var keyColumn = datasheet1.getRange("M:M").getValues();
  var rowIndex = keyColumn.findIndex(row => row[0] === keyToSearch);
  // Search for the key in column C of تفاصيل الاوردر
  var keyColumn2 = detailsSheet.getRange("C:C").getValues();
  var rowIndex2 = -1; // Initialize rowIndex2 to -1, indicating no match

  for (var i = 0; i < keyColumn2.length; i++) {
    if (keyColumn2[i][0] === keyToSearch) {
      rowIndex2 = i; // Set rowIndex2 to the index of the first match
      break; // Exit the loop once the first match is found
    }
  }

  var shippingState = shShippingStatus.getRange("J3").getValue();  //shipping state
  var shippingValue = shShippingStatus.getRange("K27").getValue();  //shipping state

  
  if(shShippingStatus.getRange("K27").isBlank()==true){
    SpreadsheetApp.getUi().alert("من فضلك ادخل قيمة الشحن المدفوع");
  }
  else{

    if(shippingState == "تم التسليم"){
      var dataRange = detailsSheet.getRange(rowIndex2 +1, 8, 5, 2); // داتا البيع
      var dataValues = dataRange.getValues();

      // المكان الجديد
      var targetRange = detailsSheet.getRange(rowIndex2 +1, 10, 5, 2);
      targetRange.setValues(dataValues);
      var paidShipping = detailsSheet.getRange(rowIndex2 +1, 13);
      paidShipping.setValue(shippingValue);

      datasheet1.getRange(rowIndex+1, 17).setValue(shShippingStatus.getRange("J3").getValue());

    } else if(shippingState == "تم التسليم جزئيا"){
      var dataRange = shShippingStatus.getRange("J21:K25"); // داتا البيع
      var dataValues = dataRange.getValues();

      // المكان الجديد
      var targetRange = detailsSheet.getRange(rowIndex2 +1, 10, 5, 2);
      targetRange.setValues(dataValues);
      var paidShipping = detailsSheet.getRange(rowIndex2 +1, 13);
      paidShipping.setValue(shippingValue);

      datasheet1.getRange(rowIndex+1, 17).setValue(shShippingStatus.getRange("J3").getValue());

    }else if(shippingState == "مرتجع"){
      // المكان الجديد
      var targetRange = detailsSheet.getRange(rowIndex2 +1, 10, 5, 2);
      targetRange.setValue(0);
      var paidShipping = detailsSheet.getRange(rowIndex2 +1, 13);
      paidShipping.setValue(shippingValue);

      datasheet1.getRange(rowIndex+1, 17).setValue(shShippingStatus.getRange("J3").getValue());

    }else{
      SpreadsheetApp.getUi().alert("ادخل حالة الشحن بشكل صحيح");
    }

    SpreadsheetApp.getUi().alert("تم حفظ حالة الشحن");


    //Clearnign the data from the Data Entry Form

      shShippingStatus.getRange("E6").clearContent();
      shShippingStatus.getRange("E9").clearContent();
      shShippingStatus.getRange("E12").clearContent();
      shShippingStatus.getRange("H12").clearContent();
      shShippingStatus.getRange("E14").clearContent();
      shShippingStatus.getRange("E16").clearContent();
      shShippingStatus.getRange("H16").clearContent();
      shShippingStatus.getRange("I27").clearContent();
      shShippingStatus.getRange("K27").clearContent();
      shShippingStatus.getRange("H14").clearContent();
      shShippingStatus.getRange("E16").clearContent();
      shShippingStatus.getRange("E18").clearContent();
      shShippingStatus.getRange("J3").clearContent();
      shShippingStatus.getRange("E1").clearContent();
      shShippingStatus.getRange("H1").clearContent();

      shShippingStatus.getRange("D21:K25").clearContent();
  }
}





function updateShippingStatus() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var salesSheet = spreadsheet.getSheetByName('مبيعات'); // Replace with your actual sheet name

  var dataRange = salesSheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var status = values[i][13]; // Assuming the status is in column N (index 13)
    var row = i + 1;

    if (status === 'تم الشحن' && values[i][14] != "ايجاد") {
      // Update O column with "ايجاد"
      salesSheet.getRange(row, 15).setValue('ايجاد'); // Assuming O column is column 15 (index 15)

      // Update P column with the current date
      var currentDate = new Date();
      salesSheet.getRange(row, 16).setValue(currentDate); // Assuming P column is column 16 (index 16)
    } else if(status === 'حالة الاوردر'){
      // Do Nothing
    } else {
      // Clear values in O and P columns
      salesSheet.getRange(row, 15, 1, 2).clearContent(); // Assuming O and P columns are columns 15 and 16 (indices 15 and 16)
    }
  }
}

var G_editOrder_rowIndex = 0;
var G_editOrder_rowIndex2 = 0;

function editOrder_fetchDataBasedOnOrderNum() {
  var myGoogleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shShippingStatus = myGoogleSheet.getSheetByName("تعديل اوردر");
  var datasheet1 = myGoogleSheet.getSheetByName("مبيعات");
  var detailsSheet = myGoogleSheet.getSheetByName("تفاصيل الاوردر");
  var ui = SpreadsheetApp.getUi();


  // Get the key from cell H2 where the order number
  var keyToSearch = shShippingStatus.getRange("E1").getValue().trim();
  if(shShippingStatus.getRange("E1").isBlank()==true){
    ui.alert("من فضلك ادخل رقم الاوردر");
  }
  else{

    // Search for the key in column H (assuming column H contains the keys in مبيعات sheet)
    var keyColumn = datasheet1.getRange("M:M").getValues();
    var rowIndex = keyColumn.findIndex(row => row[0] === keyToSearch);
    // Search for the key in column C of تفاصيل الاوردر
    var keyColumn2 = detailsSheet.getRange("C:C").getValues();
    var rowIndex2 = keyColumn2.findIndex(row => row[0] === keyToSearch);

    G_editOrder_rowIndex = rowIndex;
    G_editOrder_rowIndex2= rowIndex2;

    if (rowIndex !== -1 && rowIndex2 !== -1) {
      // Key found, retrieve data and populate the shipping status form
      var dataRow = datasheet1.getRange(rowIndex + 1, 1, 1, datasheet1.getLastColumn()).getValues()[0];

      // Assuming the shipping status form has the same layout as the مبيعات sheet
      shShippingStatus.getRange("E6").setValue(dataRow[2]);  // Update the cell references based on your actual layout
      shShippingStatus.getRange("E9").setValue(dataRow[3]);
      shShippingStatus.getRange("E12").setValue(dataRow[4]);
      shShippingStatus.getRange("H12").setValue(dataRow[5]);
      shShippingStatus.getRange("E14").setValue(dataRow[6]);
      shShippingStatus.getRange("E16").setValue(dataRow[10]);
      //shShippingStatus.getRange("H26").setValue(dataRow[8]);
      shShippingStatus.getRange("H16").setValue(dataRow[12]);
      shShippingStatus.getRange("I27").setValue(dataRow[9]);
      shShippingStatus.getRange("H14").setValue(dataRow[7]);
      shShippingStatus.getRange("E18").setValue(dataRow[11]);

      var dataRange = detailsSheet.getRange(rowIndex2 + 1, 4, 5, 6); // Assuming columns D to I
      var dataValues = dataRange.getValues();

      // Set the values in حالة الشحن2 starting from D21 to I25
      var targetRange = shShippingStatus.getRange("D21:I25");
      targetRange.setValues(dataValues);

      ui.alert("Data retrieved successfully!");

    } else {
      // Key not found, display a message or take appropriate action
      ui.alert("رقم الاوردر غير موجود");
    }

  }
}



function editOrder_save() {

  var googleSheet = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet1 = googleSheet.getSheetByName("مبيعات");
  var editOrder = googleSheet.getSheetByName("تعديل اوردر");
  var detailsSheet = googleSheet.getSheetByName("تفاصيل الاوردر");
  var repo = googleSheet.getSheetByName("المخزن");

  var ui = SpreadsheetApp.getUi();

  var keyToSearch = editOrder.getRange("E1").getValue().trim();
  // Search for the key in column H (assuming column H contains the keys in مبيعات sheet)
  var keyColumn = datasheet1.getRange("M:M").getValues();
  var rowIndex = keyColumn.findIndex(row => row[0] === keyToSearch);
  // Search for the key in column C of تفاصيل الاوردر
  var keyColumn2 = detailsSheet.getRange("C:C").getValues();
  var rowIndex2 = -1; // Initialize rowIndex2 to -1, indicating no match

  for (var i = 0; i < keyColumn2.length; i++) {
    if (keyColumn2[i][0] === keyToSearch) {
      rowIndex2 = i; // Set rowIndex2 to the index of the first match
      break; // Exit the loop once the first match is found
    }
  }

Logger.log("rowIndex: " + rowIndex);
Logger.log("rowIndex2: " + rowIndex2);

if (validateEntry(editOrder)==true && rowIndex !== -1 && rowIndex2 !== -1) 
{
  // Get the range from D21 to H25 in the source sheet
  var dataRange = editOrder.getRange('D21:I25');
  var originalValues = dataRange.getValues();  // داتا البيع الجديدة
  
  var targetRange = detailsSheet.getRange(rowIndex2 +1, 4, 5, 6); //   داتا البيع القديمة
  var oldValues = targetRange.getValues();


  if(datasheet1.getRange(rowIndex+1, 14).getValue()=="ملغى")           // already cancelled
  {
    ui.alert("الاوردر ملغى بالفعل");
  }
  else if(editOrder.getRange("J4").getValue()=="ملغى")                 // need to be cancelled
  {
    var response = ui.alert("الغاء", 'هل تريد الغاء هذا الاوردر',ui.ButtonSet.YES_NO);
    // Checking the user response and proceed with clearing the form if user selects Yes
    if (response == ui.Button.NO) 
    {
      return;//exit from this function
    } 
    //      For المخزن

    var repokeyColumn = repo.getRange("A:A").getValues();
    //    نرجع القديم للمخزن
    for (var i = 0; i < oldValues.length; i++) {
      if(oldValues[i][0] !==""){
        var product = oldValues[i][0];
        var repoRowIndex = repokeyColumn.findIndex(row => row[0] === product);
        Logger.log("repoIndex: " + repoRowIndex);
        var value = repo.getRange(repoRowIndex+1, 2).getValue();
        repo.getRange(repoRowIndex+1, 2).setValue(value + oldValues[i][4]);
      }
    }

    datasheet1.getRange(rowIndex+1, 14).setValue('ملغى'); // حالة الاوردر
    var targetRange2 = detailsSheet.getRange(rowIndex2 +1, 6, 5, 4);
    targetRange2.setValue(0);

    ui.alert("تم الغاء الاوردر");

  }
  else                                                                     // just edit
  {
    var repokeyColumn = repo.getRange("A:A").getValues();
    //    نرجع القديم للمخزن
    for (var i = 0; i < oldValues.length; i++) 
    {
      if(oldValues[i][0] !=="")
      {
        var product = oldValues[i][0];
        var repoRowIndex = repokeyColumn.findIndex(row => row[0] === product);
        Logger.log("repoIndex: " + repoRowIndex);
        var value = repo.getRange(repoRowIndex+1, 2).getValue();
        repo.getRange(repoRowIndex+1, 2).setValue(value + oldValues[i][4]);
      }
    }

    //datasheet1.getRange(rowIndex+1, 2).setValue(new Date()).setNumberFormat('yyyy-mm-dd'); //Date
    datasheet1.getRange(rowIndex+1, 3).setValue(editOrder.getRange("E6").getValue()); //الاسم
    datasheet1.getRange(rowIndex+1, 4).setValue(editOrder.getRange("E9").getValue()); //العنوان
    datasheet1.getRange(rowIndex+1, 5).setValue(editOrder.getRange("E12").getValue()); //ت1
    datasheet1.getRange(rowIndex+1, 6).setValue(editOrder.getRange("H12").getValue()); // ت2
    datasheet1.getRange(rowIndex+1, 7).setValue(editOrder.getRange("E14").getValue()); //المجافظة
    datasheet1.getRange(rowIndex+1, 8).setValue(editOrder.getRange("H14").getValue());// البائع
    datasheet1.getRange(rowIndex+1, 9).setValue(editOrder.getRange("I26").getValue());// سعر الاوردر
    datasheet1.getRange(rowIndex+1, 10).setValue(editOrder.getRange("I27").getValue());// الشحن
    datasheet1.getRange(rowIndex+1, 11).setValue(editOrder.getRange("E16").getValue());// البيدج
    datasheet1.getRange(rowIndex+1, 12).setValue(editOrder.getRange("E18").getValue());// ملاحظات
    datasheet1.getRange(rowIndex+1, 13).setValue(editOrder.getRange("H16").getValue());// رقم الاوردر
    datasheet1.getRange(rowIndex+1, 14).setValue('لم يتم التحضير'); // حالة الاوردر

    targetRange.setValues(originalValues);
    var shippingInTarget = detailsSheet.getRange(rowIndex2 + 1,12);
    shippingInTarget.setValue(editOrder.getRange("I27").getValue());

    detailsSheet.getRange(rowIndex2 +1, 3,5,1).setValue(editOrder.getRange("H16").getValue());
    detailsSheet.getRange(rowIndex2 +1, 1,5,1).setValue(editOrder.getRange("E16").getValue());


    validateRepo(editOrder);

    //    نخصم الجديد من المخزن
    for (var i = 0; i < originalValues.length; i++) 
    {
      if(originalValues[i][0] !=="")
      {
        var product = originalValues[i][0];
        var repoRowIndex = repokeyColumn.findIndex(row => row[0] === product);
        var value = repo.getRange(repoRowIndex+1, 2).getValue();
        repo.getRange(repoRowIndex+1, 2).setValue(value - originalValues[i][4]);
      }
    }

    ui.alert("تم تعديل الاوردر");
  }

  //Clearnign the data from the Data Entry Form

  editOrder.getRange("E6").clearContent();
  editOrder.getRange("E9").clearContent();
  editOrder.getRange("E12").clearContent();
  editOrder.getRange("H12").clearContent();
  editOrder.getRange("E14").clearContent();
  editOrder.getRange("E16").clearContent();
  editOrder.getRange("H16").clearContent();
  editOrder.getRange("I27").clearContent();
  editOrder.getRange("H14").clearContent();
  editOrder.getRange("E16").clearContent();
  editOrder.getRange("E18").clearContent();
  editOrder.getRange("J4").clearContent();
  editOrder.getRange("D21:I25").clearContent();
}
else
{
  ui.alert("رقم الاوردر غير موجود");

}

  
}






// Function to submit the data to Database sheet

function buyProducts() {

  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 

  var shUserForm= myGooglSheet.getSheetByName("فاتورة مشتريات"); //delcare a variable and set with the User Form worksheet
  var datasheet2 = myGooglSheet.getSheetByName("مشتريات");
  var repo = myGooglSheet.getSheetByName("المخزن");


  

  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  var response = ui.alert("Submit", 'Do you want to submit the data?',ui.ButtonSet.YES_NO);
  // Checking the user response and proceed with clearing the form if user selects Yes
  if (response == ui.Button.NO) {
    return;//exit from this function
  } 

  //Validating the entry. If validation is true then proceed with transferring the data to Database sheet
  if(!(shUserForm.getRange("E6").isBlank()==true) && 
     (!(shUserForm.getRange("D12").isBlank()==true &&
        shUserForm.getRange("D13").isBlank()==true &&
        shUserForm.getRange("D14").isBlank()==true &&
        shUserForm.getRange("D15").isBlank()==true &&
        shUserForm.getRange("D16").isBlank()==true))){
      // Filter non-empty rows
    var filteredValues = [];
      // Get the range from D21 to H25 in the source sheet
    var dataRange = shUserForm.getRange('D12:I16');
    var originalValues = dataRange.getValues();

    for (var i = 0; i < originalValues.length; i++) {
      var row = originalValues[i];

      if (row[0]!=="") {
        filteredValues.push(row);
      }
    }

    // Calculate the dimensions of the new range
    var numRows = filteredValues.length;
    var numCols = originalValues[0].length;


    // Get the target range in the target sheet
    var lastRow = datasheet2.getLastRow();
    var targetRange = datasheet2.getRange(lastRow + 1,3 , numRows, numCols);
    // Set values in the target range
    targetRange.setValues(filteredValues);


    targetRange.offset(0,-1,numRows,1).setValue(shUserForm.getRange("E6").getValue());
    targetRange.offset(0,-2,numRows,1).setValue(new Date().toJSON().slice(0, 10)).setNumberFormat('yyyy-mm-dd');

    var keyColumn = repo.getRange("A:A").getValues();

    for (var i = 0; i < filteredValues.length; i++) {
      var product = filteredValues[i][0];
      var rowIndex = keyColumn.findIndex(row => row[0] === product);
      var quantity = repo.getRange(rowIndex+1, 2).getValue();                 //last quantity
      repo.getRange(rowIndex+1, 2).setValue(quantity + filteredValues[i][3]); //update quantity
      var price = repo.getRange(rowIndex+1, 3).getValue();                    //last price
      if(quantity>0)
      {
        repo.getRange(rowIndex+1, 3).setValue(((quantity*price)+(filteredValues[i][3]*filteredValues[i][2]))/(quantity + filteredValues[i][3])); //update price
      }
      else
      {
        repo.getRange(rowIndex+1, 3).setValue(filteredValues[i][2]);
      }
      

    }

    // Show message after complete
    ui.alert("New Data Saved");
      

    //Clearnign the data from the Data Entry Form

      shUserForm.getRange("E6").clearContent();
      shUserForm.getRange("D12:I16").clearContent();

  }else{
    ui.alert("من فضلك اكمل بيانات الفاتورة");
  }



}

  




//Declare a function to validate the entry made by user in UserForm

function validateEntry(shUserForm){

  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  //var shUserForm= myGooglSheet.getSheetByName("فاتورة اوردر");//delcare a variable and set with the User Form worksheet
  var repo = myGooglSheet.getSheetByName("المخزن");
  var ui = SpreadsheetApp.getUi();

  

//Validating Employee ID

  if(shUserForm.getRange("E6").isBlank()==true ||
     shUserForm.getRange("E9").isBlank()==true ||
    shUserForm.getRange("E14").isBlank()==true ||
    shUserForm.getRange("E16").isBlank()==true ||
    shUserForm.getRange("H16").isBlank()==true ||
    shUserForm.getRange("H14").isBlank()==true ||
    shUserForm.getRange("E12").isBlank()==true ){

    ui.alert("من فضلك أكمل بيانات العميل");

    
    return false;

  }

 //Validating Employee Name

  else if((shUserForm.getRange("D21").isBlank()==true &&
    shUserForm.getRange("D22").isBlank()==true &&
    shUserForm.getRange("D23").isBlank()==true &&
    shUserForm.getRange("D24").isBlank()==true &&
    shUserForm.getRange("D25").isBlank()==true)){

    ui.alert("لا يمكن اضافة اوردر بدون اى منتجات");


    return false;

  }

  //Validating Gender

  else if(shUserForm.getRange("I27").isBlank()==true){

    ui.alert("من فضلك ادخل سعر الشحن او ادخل كلمة مجانى");

    return false;

  }

  else if(shUserForm.getRange("I27").isBlank()==true){

    ui.alert("من فضلك ادخل سعر الشحن او ادخل كلمة مجانى");

    return false;

  }


  
  return true ;

  


}



function validateRepo(shUserForm){

  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  //var shUserForm= myGooglSheet.getSheetByName("فاتورة اوردر");//delcare a variable and set with the User Form worksheet
  var repo = myGooglSheet.getSheetByName("المخزن");
  var ui = SpreadsheetApp.getUi();

  var dataRange = shUserForm.getRange('D21:I25');
  var originalValues = dataRange.getValues();
  var keyColumn = repo.getRange("A:A").getValues();

  for (var i = 0; i < originalValues.length; i++) {
    if(originalValues[i][0] !==""){
      var product = originalValues[i][0];
      var rowIndex = keyColumn.findIndex(row => row[0] === product);
      var value = repo.getRange(rowIndex+1, 2).getValue();
      if(value < originalValues[i][4])
      {
        ui.alert(" لا يوجد كمية متوفرة داخل المخزن من المنتج"  + "  " + originalValues[i][0] );
      }
      
    }

  }

  return true;


}











