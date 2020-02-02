/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var foodItems;
var selectBox;


Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
if (!Office.context.requirements.isSetSupported('ExcelApi', '1.9')) {
  document.write('Sorry. The add-in uses Excel.js APIs that are not available in your version of Office.');
}

// Assign event handlers and other initialization logic.
//
if (localStorage.getItem("foodItems") === null) {
  foodItems = [];
}
else
{

  foodItems = JSON.parse(localStorage.foodItems);
}

selectBox = document.getElementById('foodItems');

refreshList();


document.getElementById("insert-food").onclick = insertFood;
document.getElementById("add-food").onclick = addFood;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function insertFood() {
  if (typeof(Storage) !== "undefined") {
    // Code for localStorage/sessionStorage.
  Excel.run(function (ctx) { 
    var foodNameRange = ctx.workbook.getSelectedRange();
    var caloriesRange = foodNameRange.getColumnsAfter(1);
    var foodItemIndex = selectBox.selectedIndex;
    var foodName = "";
    var foodCalories = "";

    
    foodNameRange.load("values");
    caloriesRange.load("values");
    foodNameRange.load('address');
    caloriesRange.load('address');

   return ctx.sync().then(function() {
     // document.write(foodNameRange.values);
  

    if(foodNameRange.values=="") {
      foodName = foodItems[foodItemIndex].foodName;
      foodCalories = "= "+ foodItems[foodItemIndex].foodCalories;
    }
    else
    {
      foodName =  foodNameRange.values + " + "+foodItems[foodItemIndex].foodName;
      foodCalories = "= "+caloriesRange.values + " + "+foodItems[foodItemIndex].foodCalories;
    }


    
      foodNameRange.values = [[ foodName ]];
      caloriesRange.values = [[ foodCalories ]];
    });
}).catch(function(error) {
  document.write("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      document.write("Debug info: " + JSON.stringify(error.debugInfo));
    }
});


  // Excel.run(function (context) {

  //  // var range = context.workbook.getSelectedRange();
  //  // range.load(['address', 'values']);

  //  // var firstSelectedCellValue = range.values[0][0];
    
  // //  OfficeHelpers.UI.notify('Selected range is: ');

  //     // TODO1: Queue table creation logic here.
  //     var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
  //     //var expensesTable = currentWorksheet.tables.add("A1:B1", true /*hasHeaders*/);
  //     //expensesTable.name = "ExpensesTable";

  //    const selectedRange = context.workbook.getSelectedRange();
  //    // document.write(Object.getOwnPropertyNames(selectedRanges.address));
  //   //  selectedRanges.format.fill.color = "blue";
  
  //   //  var firstSelectedCellValue = range.values[0][0];


  //  //  expensesTable.rows.add(null /*add at the end*/, ["1/1/2017", range] ]);

  // //  currentWorksheet.rows.add(null /*add at the end*/, [["1/1/2017", "The Phone Company"]]);

  //     // TODO3: Queue commands to format the table.
  //   //  expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
  //   //  expensesTable.getRange().format.autofitColumns();
  //   //  expensesTable.getRange().format.autofitRows();
  // //  document.write(Date());

  //     return context.sync().then(function() {document.write(selectedRange.address);});
  // })
  // .catch(function (error) {
  //     //console.log("Error: " + error);
  //     document.write("Error: " + error);
  //     if (error instanceof OfficeExtension.Error) {
  //       //  console.log("Debug info: " + JSON.stringify(error.debugInfo));
  //       document.write("Debug info: " + JSON.stringify(error.debugInfo));
  //     }
  // });
} else {
  // Sorry! No Web Storage support..
  document.write("Looks like you need SQL. Fuck!");
}
}

function addFood() {
  if (typeof(Storage) !== "undefined") {
    // Code for localStorage/sessionStorage.
   // document.write(Object.keys(localStorage));

    var foodName = document.getElementById("foodname").value;
    var foodCalories = document.getElementById("caloriecount").value;
    foodItems.push({foodName:foodName, foodCalories:foodCalories});

             // Store
            localStorage.setItem("foodItems", JSON.stringify(foodItems));

            refreshList();

            // Retrieve
           // document.write(localStorage.foodItems);

  } else {
    // Sorry! No Web Storage support..
    document.write("Looks like you need SQL. Fuck!");
  }
}

function refreshList() {
  for(var i = 0, l = foodItems.length; i < l; i++){
    var option = foodItems[i];
    selectBox.options.add( new Option(option.foodName, option.foodCalories) );
    //foodName:foodName, foodCalories:foodCalories
  }
}


