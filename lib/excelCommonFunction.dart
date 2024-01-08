
import 'package:flutter/material.dart';
import 'package:office_addin_helper/office_addin_helper.dart';
import 'package:officejs/officejs.dart';

const excelMaxRowsCellsCount=1048576;
const excelMaxColumnsCellsCount=16384;

extension SheetModelExtension on SheetModel {
  ExcelSheetModel<Worksheet> toExcelSheetNewCode() {
    if (this is! ExcelSheetModel) throw UnsupportedError('$this');
    return this as ExcelSheetModel<Worksheet>;
  }
}

// for formatting use below link
// https://learn.microsoft.com/en-us/javascript/api/excel/excel.rangeformat?view=excel-js-preview

// https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.clearapplyto?view=office-scripts
enum RangeClearProperties {
  All,
  Contents, //Clears the contents of the range.

  Formats, //Clears all formatting for the range.

  Hyperlinks, //Clears all hyperlinks, but leaves all content and formatting intact.

  RemoveHyperlinks, // Removes hyperlinks and formatting for the cell but leaves content, conditional formats, and data validation intact.
}
// https://learn.microsoft.com/en-us/javascript/api/excel/excel.changedirectionstate?view=excel-js-preview
enum RangeDeleteProperties {
  Up,
  Left,
  Down,
  Right,
}


//https://learn.microsoft.com/en-us/javascript/api/excel/excel.insertshiftdirection?view=excel-js-preview
//enum for insertion
enum RowColumnInsertion{
  Down,
  Right
}


class ExcelInteraction with ChangeNotifier{

  late ExcelApiI excelApi;

  void declareApi({required ExcelApiI api}){
    this.excelApi=api;
  }

  //activate sheet by name
  Future<void> activateSheetByName({required String sheetName})async{
    try{
      await excelApi.onLoad();
      //PROOF:1 Activate Sheet By Name
      //code to activate sheet by name
      var sheets=await excelApi.getSheets();
      var Sheet1Name=sheets.where((element) => element.name=='${sheetName}').first;
      excelApi.setActiveSheet(Sheet1Name);
    }catch(error){
      showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
    }
  }

  //read data from sheet
  Future<List<dynamic>> getRowData({
    required ExcelSheetModel<Worksheet> sheet,
    required int rowIndex,
    required int columnCount,
    required int startColumn,
    required int rowCount,
  }) async {
    try{
      final excelRange = sheet.worksheet.getRangeByIndexes(
        startRow: rowIndex,
        startColumn: startColumn,
        rowCount: rowCount,
        columnCount: columnCount,
      );

      // Load the values for the specified row
      // Load the values for the specified row
      excelRange.load(['values', 'rowCount', 'columnCount', 'format']);

      // Synchronize the context to execute the load operation
      await excelRange.context.sync();

      // Return the values as a List<dynamic>
      return excelRange.values;
    }catch(error){
      showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
    }
    return [];
  }

  //update row data
  Future<void> updateRowData({
    required ExcelSheetModel<Worksheet> sheet,
    required int rowIndex,
    required int columnCount,
    required int startColumn,
    required int rowCount,
    required List<List<dynamic>> newValues,
  }) async {
    try{
      // Get the range for the specified row
      final excelRange = sheet.worksheet.getRangeByIndexes(
        startRow: rowIndex,
        startColumn: startColumn,
        rowCount: rowCount,
        columnCount: columnCount,
      );

      // Add the range to tracked objects
      sheet.worksheet.context.trackedObjects.add(excelRange);

      try {
        // Load the necessary properties (e.g., 'values', 'rowCount', 'columnCount', 'format')
        excelRange.load(['values', 'rowCount', 'columnCount', 'format']);

        // Synchronize the context to execute the load operation
        await sheet.worksheet.context.sync();

        // Update the values of the range
        excelRange.values = newValues;

        // Synchronize the context to apply the changes
        await sheet.worksheet.context.sync();
      }catch(error){
        showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
      } finally {
        // Remove the range from tracked objects
        sheet.worksheet.context.trackedObjects.remove(excelRange);
      }
    }catch(error){
      showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
    }
  }


  //insert column/row range
  Future<void> insertColumnOrRowRangeCustom({
    required ExcelSheetModel<Worksheet> sheet,
    required int startRowIndex,
    required int startColumnIndex,
    required int rowCount,
    required int columnCount,
    bool insertColumn=true,
  }) async {
    try{
      var range = sheet.worksheet.getRangeByIndexes(
        startRow: startRowIndex,
        startColumn: startColumnIndex,
        rowCount: rowCount,
        columnCount:
        columnCount, // Assuming you want to insert rows in a single column.
      );

      //
      var excelRange = range;

      // Add the range to tracked objects
      sheet.worksheet.context.trackedObjects.add(excelRange);

      try {
        // Load the necessary properties (e.g., 'values', 'rowCount', 'columnCount', 'format')
        excelRange.load(['values', 'rowCount', 'columnCount', 'format']);

        // Synchronize the context to execute the load operation
        await sheet.worksheet.context.sync();

        //this is for moving single cell
        //we can use Right/Down
        await range.insert(insertColumn?'Right':'Down');
        //this is for moving row cell
        // range.insertRows();

        // Synchronize the context to apply the changes
        await sheet.worksheet.context.sync();
      }catch(error){
        showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
      } finally {
        // Remove the range from tracked objects
        sheet.worksheet.context.trackedObjects.remove(excelRange);
      }
    }catch(error){
      showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
    }
  }


  //clear custom
  Future<void> clearRangeCustom({
    required ExcelSheetModel<Worksheet> sheet,
    required int startRowIndex,
    required int startColumnIndex,
    required int rowCount,
    required int columnCount,
    RangeClearProperties t=RangeClearProperties.All,
  }) async {
    try{
      var range = sheet.worksheet.getRangeByIndexes(
        startRow: startRowIndex,
        startColumn: startColumnIndex,
        rowCount: rowCount,
        columnCount:
        columnCount, // Assuming you want to insert rows in a single column.
      );

      var excelRange = range;

      // Add the range to tracked objects
      sheet.worksheet.context.trackedObjects.add(excelRange);

      try {
        // Load the necessary properties (e.g., 'values', 'rowCount', 'columnCount', 'format')
        excelRange.load(['values', 'rowCount', 'columnCount', 'format']);

        // Synchronize the context to execute the load operation
        await sheet.worksheet.context.sync();

        await range.clear(t.name.toString());

        // Synchronize the context to apply the changes
        await sheet.worksheet.context.sync();
      }catch(error){
        showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
      } finally {
        // Remove the range from tracked objects
        sheet.worksheet.context.trackedObjects.remove(excelRange);
      }
    }catch(error){
      showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
    }
  }

  //delete common function
  Future<void> deleteRangeCustom({
    required ExcelSheetModel<Worksheet> sheet,
    required int startRowIndex,
    required int startColumnIndex,
    required int rowCount,
    required int columnCount,
    RangeDeleteProperties fromT=RangeDeleteProperties.Left
  }) async {

    try{
      var range = sheet.worksheet.getRangeByIndexes(
        startRow: startRowIndex,
        startColumn: startColumnIndex,
        rowCount: rowCount,
        columnCount:
        columnCount, // Assuming you want to insert rows in a single column.
      );

      var excelRange = range;
      // Add the range to tracked objects
      sheet.worksheet.context.trackedObjects.add(excelRange);

      try {
        // Load the necessary properties (e.g., 'values', 'rowCount', 'columnCount', 'format')
        excelRange.load(['values', 'rowCount', 'columnCount', 'format']);

        // Synchronize the context to execute the load operation
        await sheet.worksheet.context.sync();

        //   Up,
        // Left,
        // Down,
        // Right,

        await excelRange.delete(fromT.name);

        // Synchronize the context to apply the changes
        await sheet.worksheet.context.sync();
      }catch(error){
        showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
      } finally {
        // Remove the range from tracked objects
        sheet.worksheet.context.trackedObjects.remove(excelRange);
      }
    }catch(error){
      showCustomDialogCommon(message: 'Error ${error}',title: 'Error!');
    }
  }



}

showCustomDialogCommon({String message='',String title='Info!',}){

}