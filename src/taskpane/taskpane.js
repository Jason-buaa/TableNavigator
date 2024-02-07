/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("setup").onclick = setup;
    document.getElementById("list-tableNames").onclick = listTableNames;
    document.getElementById("save-conditionalFormat").onclick = saveConditionalFormats;
    document.getElementById("enable-CellHighlight").onclick = enableCellHighlight;
    document.getElementById("disable-CellHighlight").onclick = disableCellHighlight;
  }
});

async function enableCellHighlight(){
  await saveConditionalFormats();
  await Excel.run(async (context) => {
    let workbook = context.workbook;
    cellHightHandlerResult = workbook.onSelectionChanged.add(CellHighlightHandler);
    await context.sync();
  });
}

async function disableCellHighlight(){
  await clearMeekouFormat();
  await Excel.run(cellHightHandlerResult.context, async (context) => {
    cellHightHandlerResult.remove();
    await context.sync();
    cellHightHandlerResult = null;
  });
}
async function saveConditionalFormats() {
  // 用于存储条件格式的数据结构，可以根据需求选择数组、对象等
  let savedFormats = {};

  await Excel.run(async (context) => {
    let workbook = context.workbook;
    let worksheets = workbook.worksheets;
    worksheets.load("items/name");

    // 同步以获取工作表信息
    await context.sync();

    // 遍历每个工作表
    for (let i = 0; i < worksheets.items.length; i++) {
      let sheet = worksheets.items[i];

      // 获取工作表上范围的条件格式
      let conditionalFormats = sheet.getUsedRange().conditionalFormats;
      conditionalFormats.load("items");

      // 同步以获取条件格式信息
      await context.sync();

      // 如果有条件格式，则保存到数据结构中
      if (conditionalFormats.items.length > 0) {
        savedFormats[sheet.name] = conditionalFormats.items.map((format) => {
          let savedFormat = {
            type: format.type,
          };

          // 根据条件格式类型选择性地保存属性
          switch (format.type) {
            case Excel.ConditionalFormatType.custom:
              savedFormat.rule = format.custom.rule.formula;
              savedFormat.fill = format.custom.format.fill.color;
              savedFormat.fontColor = format.custom.format.font.color;
              savedFormat.borders = format.custom.format.borders;
              break;
            case Excel.ConditionalFormatType.dataBar:
              savedFormat.dataBar = format.dataBar;
              break;
            case Excel.ConditionalFormatType.colorScale:
              savedFormat.colorScale = format.colorScale;
              break;
            case Excel.ConditionalFormatType.iconSet:
              savedFormat.iconSet = format.iconSet;
              break;
            // 可根据需要添加其他条件格式类型的处理
          }

          return savedFormat;
        });
      }
    }
  });

  // 在控制台输出保存的条件格式信息（可根据需求将其发送到服务器等）
  console.log(savedFormats);
}

async function clearMeekouFormat() {
  await Excel.run(async (context) => {
    let workbook = context.workbook;
    let worksheets = workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();
    worksheets.items.forEach(async (s) => {
      let conditionalFormats = s.getRange().conditionalFormats;
      conditionalFormats.clearAll();
      await context.sync();
    });
  });
}
async function CellHighlightHandler(){
      //await clearMeekouFormat();
      await Excel.run(async (context) => {
        let workbook = context.workbook;
        let sheets = workbook.worksheets;
        let selection = workbook.getSelectedRange();
        selection.load("rowIndex,columnIndex");
        sheets.load("items");
        await context.sync();
        console.log(sheets.items);
        console.log(`=ROW()= + ${selection.rowIndex + 1})`);
        // add new conditional format
        await context.sync();
        let rowConditionalFormat = selection.getEntireRow().conditionalFormats.add(Excel.ConditionalFormatType.custom);
        rowConditionalFormat.custom.format.fill.color = "green";
        rowConditionalFormat.custom.rule.formula = `=ROW()=  ${selection.rowIndex + 1}`;
        let columnConditionalFormat = selection
          .getEntireColumn()
          .conditionalFormats.add(Excel.ConditionalFormatType.custom);
        columnConditionalFormat.custom.format.fill.color = "green";
        columnConditionalFormat.custom.rule.formula = `=Column()=  ${selection.columnIndex + 1}`;
        await context.sync();
      });
    }
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



async function listTableNames() {
  await Excel.run(async (context) => {

    let tables = context.workbook.tables;

    // 加载表格的名称
    tables.load("name");
  
    // 执行同步操作以获取加载的数据
    return context.sync()
      .then(function () {
        // 输出表格的名称到控制台
        for (var i = 0; i < tables.items.length; i++) {
          console.log('Table Name '+ i+ ' ' + tables.items[i].name);
        }
      });
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