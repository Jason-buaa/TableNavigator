/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;
    document.getElementById("setup").onclick = setup;
  }
});

async function setup() {
  await Excel.run(async (context) => {

      // Queue table creation logic here.

      context.workbook.worksheets.getItemOrNullObject("Sample").delete();
      const currentWorksheet = context.workbook.worksheets.add("Sample");

      let temperatureTable = currentWorksheet.tables.add("A1:M1", true);
      temperatureTable.name = "TemperatureTable";
      temperatureTable.getHeaderRowRange().values = [
        ["Category", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
      ];
      temperatureTable.rows.add(null, [
        ["Avg High", 40, 38, 44, 45, 51, 56, 67, 72, 79, 59, 45, 41],
        ["Avg Low", 34, 33, 38, 41, 45, 48, 51, 55, 54, 45, 41, 38],
        ["Record High", 61, 69, 79, 83, 95, 97, 100, 101, 94, 87, 72, 66],
        ["Record Low", 0, 2, 9, 24, 28, 32, 36, 39, 35, 21, 12, 4]
      ]);


      let salesTable = currentWorksheet.tables.add("A7:E7", true);
      salesTable.name = "SalesTable";
      salesTable.getHeaderRowRange().values = [["Sales Team", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];
      salesTable.rows.add(null, [
        ["Asian Team 1", 500, 700, 654, 234],
        ["Asian Team 2", 400, 323, 276, 345],
        ["Asian Team 3", 1200, 876, 845, 456],
        ["Euro Team 1", 600, 500, 854, 567],
        ["Euro Team 2", 5001, 2232, 4763, 678],
        ["Euro Team 3", 130, 776, 104, 789]
      ]);

      let projectTable = currentWorksheet.tables.add("A15:D15", true);
      projectTable.name = "ProjectTable";
      projectTable.getHeaderRowRange().values = [["Project", "Alpha", "Beta", "Ship"]];
      projectTable.rows.add(null, [
        ["Project 1", "Complete", "Ongoing", "On Schedule"],
        ["Project 2", "Complete", "Complete", "On Schedule"],
        ["Project 3", "Ongoing", "Not Started", "Delayed"]
      ]);

      let profitLossTable = currentWorksheet.tables.add("A20:E20", true);
      profitLossTable.name = "ProfitLossTable";
      profitLossTable.getHeaderRowRange().values = [["Company", "2013", "2014", "2015", "2016"]];
      profitLossTable.rows.add(null, [
        ["Contoso", 256.0, -55.31, 68.9, -82.13],
        ["Fabrikam", 454.0, 75.29, -88.88, 781.87],
        ["Northwind", -858.21, 35.33, 49.01, 112.68]
      ]);
      profitLossTable.getDataBodyRange().numberFormat = [["$#,##0.00"]];
      
      await context.sync();
  });
}



/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}