/* eslint-disable no-undef */
import { Component, Inject } from "@angular/core";
import { HttpClient } from '@angular/common/http';

// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
import { map } from "rxjs-compat/operator/map";

/* global console, Excel, require */

@Component({
  selector: "app-home",
  template: require('./app.component.html')
})
export default class AppComponent {
  dataString: string;
  simpleSheetName: string = "SimpleSheet";
  dataTestSheetName: string = "DataTestSheet";
  welcomeMessage = "Welcome";
  dataStringPost: string;
  //Injection not easy
  //https://github.com/OfficeDev/Office-Addin-TaskPane-Angular/issues/45
  constructor(@Inject(HttpClient) private http: HttpClient) { }

  async run() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }

  async addTable() {
    await Excel.run(async context => {
      /* add sheet*/

      const sheets = context.workbook.worksheets;
      const sheet = sheets.add(this.simpleSheetName);
      sheet.load();
      await context.sync();

      console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);

      /* add table */
      var simpleSheet = context.workbook.worksheets.getItem(this.simpleSheetName);
      var expensesTable = simpleSheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

      expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
      ]);

      // eslint-disable-next-line no-undef
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      simpleSheet.activate();

      return context.sync();
    }).catch(Error);
    console.error(Error);
  }

  async deleteWorksheet() {
    await Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");

      await context.sync();

      if (sheets.items.length > 1) {
        const lastSheet = sheets.items[sheets.items.length - 1];

        console.log(`Deleting worksheet named "${lastSheet.name}"`);
        lastSheet.delete();

        await context.sync();
      } else {
        console.log("Unable to delete the last worksheet in the workbook");
      }
    });
  }

  async filter() {
    Excel.run(async context => {
      var expensesTable = this.getExppensesTable(context);
      var sortFields = [
        {
          key: 1, //la deuxième column : Merchant
          ascending: true
        }
      ];

      expensesTable.sort.apply(sortFields);

      return context.sync();
    }).catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  async createChart() {
    Excel.run(async context => {
      // TODO1: Queue commands to get the range of data to be charted.
      var currentWorksheet = context.workbook.worksheets.getItem(this.simpleSheetName);
      var expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
      var dataRange = expensesTable.getDataBodyRange();

      // TODO2: Queue command to create the chart and define its type.
      var chart = currentWorksheet.charts.add("ColumnClustered", dataRange, "Auto");
      // TODO3: Queue commands to position and format the chart.
      chart.setPosition("A15", "F30");
      chart.title.text = "Expenses";
      chart.legend.position = "Right";
      chart.legend.format.fill.setSolidColor("white");
      chart.dataLabels.format.font.size = 15;
      chart.dataLabels.format.font.color = "black";
      chart.series.getItemAt(0).name = "Value in &euro;";

      return context.sync();
    }).catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  async insertCell() {
    await Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      const sheet = sheets.add(this.dataTestSheetName);
      sheet.load();
      await context.sync();

      var rangeAddress = "A1:B3";
      //docs.microsoft.com/en-us/javascript/api/excel/excel.numberformatcategory?view=excel-js-preview
      var numberFormat = [
        [null, "d-mmm"],
        [null, "d-mmm"],
        [null, null]
      ];
      var values = [
        ["Today", 42147],
        ["Tomorrow", "5/24"],
        ["Difference in days", null]
      ];
      var formulas = [
        [null, null],
        [null, null],
        [null, "=B2-B1"]
      ];

      var sheetData = context.workbook.worksheets.getItem(this.dataTestSheetName);
      var range = sheetData.getRange(rangeAddress);
      range.numberFormat = numberFormat;
      range.values = values;
      range.formulas = formulas;
      range.calculate();
      range.format.autofitColumns();

      const header = range.getRow(1);
      header.format.fill.color = "#4472C4";
      header.format.font.color = "pink";

      range.load("text");

      //merge cells
      var mergeAddress = "C1:D2";
      var mergeRange = context.workbook.worksheets.getItem(this.dataTestSheetName).getRange(mergeAddress);

      const mergeCellvalues = [
        ["Merge cell", null],
        [null, null]
      ];
      mergeRange.values = mergeCellvalues;
      //cell merge has two type true and false, when true it will merge line per line
      //when false it will merge all to a cell
      //https://stackoverflow.com/questions/60588147/is-it-possible-to-merge-cells-of-different-rows-e-g-a1c2-using-excel-js-api
      mergeRange.merge(false);
      mergeRange.format.fill.color = "pink";
      mergeRange.format.font.color = "white";
      mergeRange.format.font.bold = true;
      mergeRange.format.autofitColumns();
      mergeRange.format.horizontalAlignment = "Center";
      mergeRange.format.verticalAlignment = "Center";
      mergeRange.load("text");

      //dropdowon cell
      //https://docs.microsoft.com/en-us/javascript/api/excel/excel.datavalidation?view=excel-js-preview#rule
      var dropDownRange = sheetData.getRange("E1:F3");

      dropDownRange.dataValidation.clear();
      dropDownRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "Drop Down Cell, GRECO, REFOG, LOL, DOTA2"
        }
      };
      dropDownRange.dataValidation.errorAlert = {
        message: "Sorry, your selection not in the list",
        showAlert: true,
        style: "Stop",
        title: "Data not valide"
      };
      //Insert default value in this part
      var defaultValue = "Drop Down Cell";
      dropDownRange.values = [
        [defaultValue, defaultValue],
        [defaultValue, defaultValue],
        [defaultValue, defaultValue]
      ];

      sheetData.activate();
      return context.sync().then(function() {
        console.log(range.text);
      });
    });
  }

  async callHttpGetApiBack() {
    //let url = "https://jsonplaceholder.typicode.com/todos/1";

    let url = "https://localhost:443/api/hello";
    
    console.log("before http")
    this.http.get(url).subscribe(
      data => {
        this.dataString = JSON.stringify(data);
        console.log(this.dataString);
      }
    )
    console.log("after http")
    //let fetchResult = await fetch(url);
    //let json = await fetchResult.json();
    //let result = JSON.stringify(json);
    console.log("result" + this.dataString)
    await Excel.run(async context => {
      var apiRangeAddress = "I1:J1";
      var sheetData = context.workbook.worksheets.getItem(this.dataTestSheetName);
      var apiRange = sheetData.getRange(apiRangeAddress);
      apiRange.values = [[this.dataString, null]];
      apiRange.merge(false);
      apiRange.format.fill.color = "pink";
      apiRange.format.font.color = "white";
      apiRange.format.font.bold = true;
      apiRange.format.autofitColumns();
      apiRange.format.horizontalAlignment = "Center";
      apiRange.format.verticalAlignment = "Center";
      apiRange.load("text");

      return context.sync().then(function() {
        console.log(apiRange.text);
      });
    }).catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }


  async callHttpPostApiBack(){

    console.log("before https post")
    const url = "https://localhost:443/api/addFive"

    this.http.post(url,2).subscribe(
      data => this.dataStringPost = JSON.stringify(data)
    );
    console.log("result" + this.dataStringPost)

    await Excel.run(async context => {
      var apiRangeAddress = "L1:M1";
      var sheetData = context.workbook.worksheets.getItem(this.dataTestSheetName);
      var apiRange = sheetData.getRange(apiRangeAddress);
      apiRange.values = [[this.dataStringPost, null]];
      apiRange.merge(false);
      apiRange.format.fill.color = "black";
      apiRange.format.font.color = "white";
      apiRange.format.font.bold = true;
      apiRange.format.autofitColumns();
      apiRange.format.horizontalAlignment = "Center";
      apiRange.format.verticalAlignment = "Center";
      apiRange.load("text");

      return context.sync().then(function() {
        console.log(apiRange.text);
      });
    }).catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  private getExppensesTable(context: Excel.RequestContext) {
    var currentWorksheet = context.workbook.worksheets.getItem(this.simpleSheetName);
    var expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    return expensesTable;
  }
}
