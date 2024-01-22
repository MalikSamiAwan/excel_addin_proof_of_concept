import 'package:flutter/material.dart';
import 'package:officejs/officejs.dart' as officeHelper;
import 'package:office_addin_helper/office_addin_helper.dart' as addinHelper;
import 'package:office_addin_helper/office_addin_helper.dart';
import 'package:officejs/officejs.dart';
import 'package:officejs/src/office/excel.dart' as  excelFile;
import 'package:officejs/src/office_interops/excel_js_impl.dart';
// import 'package:js/js.dart' as js;
import 'package:provider/provider.dart';

import 'package:http/http.dart' as http;
import 'dart:js' as js;


enum RangeBorderSideNames{
  EdgeTop,
  EdgeBottom,
  EdgeLeft,
  EdgeRight,
  InsideVertical,
  InsideHorizontal
}
enum RangeBorderSideStyles{
  None,
  Continuous,
  Dash,
  DashDot,
  DashDotDot,
  Dot,
  Double,
  SlantDashDot
}

const excelMaxRowsCellsCount=1048576;
const excelMaxColumnsCellsCount=16384;
const excelMaxColumnsCellsCountWithSubtract=16383;

extension SheetModelExtension on SheetModel {
  ExcelSheetModel<Worksheet> toExcelSheetNewCode() {
    if (this is! ExcelSheetModel) throw UnsupportedError('$this');
    return this as ExcelSheetModel<Worksheet>;
  }
}

extension RangeModelExtension on RangeModel {
  ExcelRangeModel<Range> toExcelRangeNewModal() {
    if (this is! ExcelRangeModel) throw UnsupportedError('$this');
    return this as ExcelRangeModel<Range>;
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
}


//https://learn.microsoft.com/en-us/javascript/api/excel/excel.insertshiftdirection?view=excel-js-preview
//enum for insertion
enum RowColumnInsertion{
  Down,
  Right
}

class NewHomeProvider extends StatefulWidget {
  const NewHomeProvider({Key? key}) : super(key: key);

  @override
  State<NewHomeProvider> createState() => _NewHomeProviderState();
}

class _NewHomeProviderState extends State<NewHomeProvider> {
  bool isLoading = false;
  bool isExcelAvailable = false;

  @override
  void didChangeDependencies() {
    refreshCode();
    super.didChangeDependencies();
  }

  Future<void> refreshCode() async {
    setState(() {
      isLoading = true;
    });
    try {
      Future<void> check() async {
        isExcelAvailable =
            await addinHelper.ExcelHelper.checkIsExcelAvailable();
      }

      for (var i = 0; i < 4; i++) {
        await Future.delayed(const Duration(milliseconds: 300));
        await check();
        if (isExcelAvailable) break;
      }
    } catch (e) {
      print(e);
    }

    setState(() {
      isLoading = false;
    });
  }

  @override
  Widget build(BuildContext context) {
    if (isLoading) {
      return const Material(
        child: Center(
          child: CircularProgressIndicator(),
        ),
      );
    }
    if (isExcelAvailable == false) {
      return const Material(
        child: Center(
          child: Text('Current Not Running on Excel'),
        ),
      );
    }
    return MultiProvider(
      providers: [
        Provider<ExcelApiI>(
          key: const ValueKey('ExcelApi'),
          create: (final context) => ExcelApiImpl(),
        ),

        // Provider<addinHelper.ExcelApiI>(
        //   key: const ValueKey('ExcelApi'),
        //   create: (final context) => addinHelper.ExcelApiImpl(),
        // ),

        // Provider<ExcelApiI>(
        // key: const ValueKey('ExcelApiMock'),
        // create: (final context) => ExcelApiMockImpl(),
        // )
      ],
      child: MyHomePage(title: 'Home Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  // This widget is the home page of your application. It is stateful, meaning
  // that it has a State object (defined below) that contains fields that affect
  // how it looks.

  // This class is the configuration for the state. It holds the values (in this
  // case the title) provided by the parent (in this case the App widget) and
  // used by the build method of the State. Fields in a Widget subclass are
  // always marked "final".

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  String _counter = '${0}';


  Future<void> applyCustomFormatting({
    required ExcelSheetModel<Worksheet> sheet,
    required int rowIndex,
    required int columnCount,
    required int startColumn,
    required int rowCount,
  }) async {
    try{

      //new code
      final excelRange = sheet.worksheet.getRangeByIndexes(
        startRow: rowIndex,
        startColumn: startColumn,
        rowCount: rowCount,
        columnCount: columnCount,
      );

// Load the values for the specified range
      excelRange.load(['values', 'rowCount', 'columnCount', 'format']);

// Add the range to the tracked objects
      sheet.worksheet.context.trackedObjects.add(excelRange);

// Synchronize the context to execute the load operation
      await sheet.worksheet.context.sync();

// Now you can access and modify the properties of excelRange
      excelRange.format.bold = true;

      // Change the background color (fill)
      excelRange.format.fontSize = 24; // Set to a specific color code

      excelRange.format.fontColor="#FF0000";
      excelRange.format.italic=true;
      excelRange.format.fontName="Arial";

      excelRange.format.fillBackgroundColor="pink";

      // var topBorder = excelRange.format.borders.getItem('EdgeTop');
      var topBorder = excelRange.format.borders.getItem('${RangeBorderSideNames.EdgeTop.name}');
      topBorder.style = '${RangeBorderSideStyles.None.name}';
      topBorder.color = '#000000';

      var btmBorder = excelRange.format.borders.getItem('${RangeBorderSideNames.EdgeBottom.name}');
      btmBorder.style = '${RangeBorderSideStyles.Continuous.name}';
      btmBorder.color = '#000000';


      var lftBorder = excelRange.format.borders.getItem('${RangeBorderSideNames.EdgeLeft.name}');
      lftBorder.style = '${RangeBorderSideStyles.Dash.name}';
      lftBorder.color = '#000000';


      var rightBorder = excelRange.format.borders.getItem('${RangeBorderSideNames.EdgeRight.name}');
      rightBorder.style = '${RangeBorderSideStyles.Dot.name}';
      rightBorder.color = '#000000';

      var insideHrzBorder = excelRange.format.borders.getItem('${RangeBorderSideNames.InsideHorizontal.name}');
      insideHrzBorder.style = '${RangeBorderSideStyles.DashDotDot.name}';
      insideHrzBorder.color = '#000000';

      var insideVrtBorder = excelRange.format.borders.getItem('${RangeBorderSideNames.InsideVertical.name}');
      insideVrtBorder.style = '${RangeBorderSideStyles.SlantDashDot.name}';
      insideVrtBorder.color = '#000000';


      await excelRange.format.context.sync();

      // excelRange.format.borderRight=true;
      // excelRange.format.borderBottom=true;
      // excelRange.format.borderLeft=true;
      // excelRange.format.borderTop=true;
      // excelRange.format.borderRight=true;






// Synchronize the context to apply the changes
      await sheet.worksheet.context.sync();

// Optionally, you can remove the object from trackedObjects if it's no longer needed
      sheet.worksheet.context.trackedObjects.remove(excelRange);


      // Step 2: Set the formatting options
      // excelRange.format.bold=true;
      // rangeFormat
      //   ..horizontalAlignment = previousFormatting.horizontalAlignment
      //   ..verticalAlignment = previousFormatting.verticalAlignment
      //   ..bold = previousFormatting.bold
      //   ..italic = previousFormatting.italic
      //   ..fontColor = previousFormatting.fontColor
      //   ..fontName = previousFormatting.fontName
      //   ..fontSize = previousFormatting.fontSize
      //   ..fillBackgroundColor = previousFormatting.fillBackgroundColor
      //   ..fillPattern = previousFormatting.fillPattern
      //   ..borderBottom = previousFormatting.borderBottom
      //   ..borderTop = previousFormatting.borderTop
      //   ..borderLeft = previousFormatting.borderLeft
      //   ..borderRight = previousFormatting.borderRight
      //   ..numberFormat = previousFormatting.numberFormat
      //   ..locked = previousFormatting.locked
      //   ..hidden = previousFormatting.hidden;

// Step 3: Apply the formatting to cell A1
//       await rangeFormat.context.sync().then((_) {
//         print("Formatting applied to cell A1");
//       });

    // await excelRange.format.context.sync();
    }catch(error){
      await showDialog(context: context,builder: (BuildContext){
        return SelectableText('Error ${error}');
      });
    }

  }

  //find method testing
  Future<void> applySearch({
    required ExcelSheetModel<Worksheet> sheet,
    required int rowIndex,
    required int columnCount,
    required int startColumn,
    required int rowCount,
  }) async {
    try {
      // ... other code ...

      // Create the range
      final range = sheet.worksheet.getRangeByIndexes(
        startRow: rowIndex,
        startColumn: startColumn,
        rowCount: rowCount,
        columnCount: columnCount,
      );

      // Add the range to the tracked objects collection
      range.context.trackedObjects.add(range);

      // Load the properties of the range
      range.load(['values', 'rowCount', 'columnCount', 'format']);

      // Synchronize the context to execute the load operation
      await range.context.sync();

      var searchText = "test"; // Define the value to search for

      // Call the find method on the range
      var foundRange = range.find(searchText, completeMatch: false, matchCase: false, searchDirection: 'Forward');


      // Load the address of the found range using the loadAddress method
      var address = await foundRange.loadAddress();

      print("Value found at: $address");
      _counter += "Value found at: $address";

      // await range.findAndActivate(searchText, completeMatch: false, matchCase: false,);

      // Remove the range from the tracked objects collection
      range.context.trackedObjects.remove(range);

      // Run the queued-up command to ensure the range is removed from tracked objects
      await range.context.sync();

      setState(() {
        // ... update state ...
      });
    } catch (error) {
      await showDialog(context: context, builder: (BuildContext) {
        return SelectableText('Error ${error}');
      });
    }

  }

  //find and replace method testing
  Future<void> applySearchAndReplace({
    required ExcelSheetModel<Worksheet> sheet,
    required int rowIndex,
    required int columnCount,
    required int startColumn,
    required int rowCount,
  }) async {
    try {
      // ... other code ...

      // Create the range
      final range = sheet.worksheet.getRangeByIndexes(
        startRow: rowIndex,
        startColumn: startColumn,
        rowCount: rowCount,
        columnCount: columnCount,
      );

      // Add the range to the tracked objects collection
      range.context.trackedObjects.add(range);

      // Load the properties of the range
      range.load(['values', 'rowCount', 'columnCount', 'format']);

      // Synchronize the context to execute the load operation
      await range.context.sync();

      var searchText = "2"; // Define the value to search for
      var replaceText = "test"; // Define the value to search for

      // Call the find method on the range
     await range.findAndReplace(replaceText: replaceText,searchText: searchText, completeMatch: false, matchCase: false);

      range.context.trackedObjects.remove(range);

      // Run the queued-up command to ensure the range is removed from tracked objects
      await range.context.sync();

    } catch (error) {
      await showDialog(context: context, builder: (BuildContext) {
        return SelectableText('Error ${error}');
      });
    }

  }

  //find and activate
  //find method testing
  Future<void> applyFindAndActivate({
    required ExcelSheetModel<Worksheet> sheet,
    required int rowIndex,
    required int columnCount,
    required int startColumn,
    required int rowCount,
  }) async {
    try {
      // ... other code ...

      // Create the range
      final range = sheet.worksheet.getRangeByIndexes(
        startRow: rowIndex,
        startColumn: startColumn,
        rowCount: rowCount,
        columnCount: columnCount,
      );

      // Add the range to the tracked objects collection
      range.context.trackedObjects.add(range);

      // Load the properties of the range
      range.load(['values', 'rowCount', 'columnCount', 'format']);

      // Synchronize the context to execute the load operation
      await range.context.sync();

      var searchText = "test"; // Define the value to search for

      // Call the find method on the range
      // var foundRange = range.find(searchText, completeMatch: false, matchCase: false, searchDirection: 'Forward');
      //
      //
      // // Load the address of the found range using the loadAddress method
      // var address = await foundRange.loadAddress();
      //
      // print("Value found at: $address");
      // _counter += "Value found at: $address";

      await range.findAndActivate(searchText, completeMatch: false, matchCase: false,);

      // Remove the range from the tracked objects collection
      range.context.trackedObjects.remove(range);

      // Run the queued-up command to ensure the range is removed from tracked objects
      await range.context.sync();

      setState(() {
        // ... update state ...
      });
    } catch (error) {
      await showDialog(context: context, builder: (BuildContext) {
        return SelectableText('Error ${error}');
      });
    }

  }

  Future<List<dynamic>> getRowData({
    required ExcelSheetModel<Worksheet> sheet,
    required int rowIndex,
    required int columnCount,
    required int startColumn,
    required int rowCount,
  }) async {
    // sheet.worksheet.load(['name', 'position', 'tabColor', 'showGridlines']);

    final excelRange = sheet.worksheet.getRangeByIndexes(
      startRow: rowIndex,
      startColumn: startColumn,
      rowCount: rowCount,
      // startColumn: 0,
      // rowCount: 1,
      columnCount: columnCount,
    );

    // Load the values for the specified row
    // Load the values for the specified row
    excelRange.load(['values', 'rowCount', 'columnCount', 'format']);

    // Synchronize the context to execute the load operation
    await excelRange.context.sync();

    // Return the values as a List<dynamic>
    return excelRange.values;
  }

  Future<void> updateRowData({
    required ExcelSheetModel<Worksheet> sheet,
    required int rowIndex,
    required int columnCount,
    required int startColumn,
    required int rowCount,
    required List<List<dynamic>> newValues,
  }) async {
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
    } finally {
      // Remove the range from tracked objects
      sheet.worksheet.context.trackedObjects.remove(excelRange);
    }
  }

  Future<void> hideRowsAndColumns({
    required ExcelSheetModel<Worksheet> sheet,
    required int startRowIndex,
    required int startColumnIndex,
    required int rowCount,
    required int columnCount,
  }) async {
    // Get the range to hide
    final excelRange = sheet.worksheet.getRangeByIndexes(
      startRow: startRowIndex,
      startColumn: startColumnIndex,
      rowCount: rowCount,
      columnCount: columnCount,
    );

    excelRange.format.wrapText =
        true; // Set to true or false based on your requirement

    // Perform sync to apply the changes
    await sheet.worksheet.context.sync();
  }

  // Future<void> hideRowsAndColumns2({
  //   required ExcelSheetModel<Worksheet> sheet,
  //   required int startRow,
  //   required int rowCount,
  //   required int startColumn,
  //   required int columnCount,
  // }) async {
  //   // Get the range for the specified rows and columns
  //   final excelRange = sheet.worksheet.getRangeByIndexes(
  //     startRow: startRow,
  //     startColumn: startColumn,
  //     rowCount: rowCount,
  //     columnCount: columnCount,
  //   );
  //
  //   // Hide rows by setting row height to zero
  //   for (int row = 0; row < rowCount; row++) {
  //     final currentRow = startRow + row;
  //     var range=excelRange.getRow(currentRow);
  //     await range.format.setColumnWidth(0.0);
  //   }
  //
  //   // Hide columns by setting column width to zero
  //   for (int col = 0; col < columnCount; col++) {
  //     final currentColumn = startColumn + col;
  //     var range=excelRange.getRow(currentColumn);
  //     await range.format.setColumnWidth(0.0);
  //   }
  //
  //   // Perform sync to apply the changes
  //   await sheet.worksheet.context.sync();
  // }

  Future<void> clearRangeCustom({
    required ExcelSheetModel<Worksheet> sheet,
    required int startRowIndex,
    required int startColumnIndex,
    required int rowCount,
    required int columnCount,
    RangeClearProperties t=RangeClearProperties.All,
  }) async {
    var range = sheet.worksheet.getRangeByIndexes(
      startRow: startRowIndex,
      startColumn: startColumnIndex,
      rowCount: rowCount,
      columnCount:
          columnCount, // Assuming you want to insert rows in a single column.
    );

    // await range.clear();
    //  await sheet.worksheet.context.sync();
    //
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
    } finally {
      // Remove the range from tracked objects
      sheet.worksheet.context.trackedObjects.remove(excelRange);
    }
  }

  Future<void> deleteRangeCustom({
    required ExcelSheetModel<Worksheet> sheet,
    required int startRowIndex,
    required int startColumnIndex,
    required int rowCount,
    required int columnCount,
    RangeDeleteProperties side=RangeDeleteProperties.Left
  }) async {
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
      //only left and up are working
      // Down,
      // Right,

      await excelRange.delete("${side.name}");

      // Synchronize the context to apply the changes
      await sheet.worksheet.context.sync();
    }catch(error){
      await showDialog(context: context,builder: (BuildContext){
        return Text('Side: ${side.name}\nError ${error}');
      });
    } finally {
      // Remove the range from tracked objects
      sheet.worksheet.context.trackedObjects.remove(excelRange);
    }
  }


  Future<void> styleRangeCustom({
    required ExcelSheetModel<Worksheet> sheet,
    required int startRowIndex,
    required int startColumnIndex,
    required int rowCount,
    required int columnCount,
  }) async {
    var range = sheet.worksheet.getRangeByIndexes(
      startRow: startRowIndex,
      startColumn: startColumnIndex,
      rowCount: rowCount,
      columnCount:
      columnCount, // Assuming you want to insert rows in a single column.
    );

    // await range.clear();
    //  await sheet.worksheet.context.sync();
    //
    var excelRange = range;

    // Add the range to tracked objects
    sheet.worksheet.context.trackedObjects.add(excelRange);

    try {
      // Load the necessary properties (e.g., 'values', 'rowCount', 'columnCount', 'format')
      excelRange.load(['values', 'rowCount', 'columnCount', 'format']);

      // Synchronize the context to execute the load operation
      await sheet.worksheet.context.sync();

      await range.format.wrapText;

      // Synchronize the context to apply the changes
      await sheet.worksheet.context.sync();
    } finally {
      // Remove the range from tracked objects
      sheet.worksheet.context.trackedObjects.remove(excelRange);
    }
  }
  
  
  //funtion1
  Future<void> insertColumnOrRowRangeCustom({
    required ExcelSheetModel<Worksheet> sheet,
    required int startRowIndex,
    required int startColumnIndex,
    required int rowCount,
    required int columnCount,
    bool insertColumn=true,
  }) async {
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
      _counter='${error}';
      setState(() {

      });
    } finally {
      // Remove the range from tracked objects
      sheet.worksheet.context.trackedObjects.remove(excelRange);
    }
  }

  Future<void> _incrementCounter({required BuildContext context}) async {
    try {
      //
      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
      // var sheet = await Provider.of<addinHelper.ExcelApiI>(context,listen: false).getActiveSheet();

      // ExcelApiI excelApi= ExcelApiImpl();
      await excelApi.onLoad();

      // addinHelper.ExcelApiI api=context.read();
      // var sheet = await api.getActiveSheet();

      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();


      //PROOF:6 Activate Sheet By Name
      //code to activate sheet by name
      // var sheets=await excelApi.getSheets();
      // var Sheet1Name=sheets.where((element) => element.name=='Sheet1').first;
      // excelApi.setActiveSheet(Sheet1Name);

      //PROOF:5 Read/Get Ranges both rows/columns
      // var res=await getRowData(rowIndex: 13,startColumn: 0,rowCount: 2,columnCount: 5,sheet: nSheet);
      // _counter="${res}";

      //PROOF:4 Update/Insert Ranges both rows/columns
      //this is used to insert/update values to specific range
      // List<List<dynamic>> values=[[10, "test", "12/12/12", "wednesday", 20],[10, "test", '${DateTime.now()}', "wednesday", 20]];
      // var res=await updateRowData(rowIndex: 13,startColumn: 1,rowCount: 2,columnCount: 5,sheet: nSheet,newValues: values);


      // var res=await hideRowsAndColumns(startRowIndex: 2,startColumnIndex: 0,rowCount: 2,columnCount: 5,sheet: nSheet,);
      // var res=await clearRangeCustom(startRowIndex: 13,startColumnIndex: 0,rowCount: 4,columnCount: 5,sheet: nSheet,);


      //PROOF:3 Delete Ranges in progress
      //delete/remove cell/rows/columns from excel
      // we have these options (  Up, Left, Down, Right, )
      // var res=await deleteRangeCustom(startRowIndex: 13,startColumnIndex: 2,rowCount: 3,columnCount: 3,sheet: nSheet,);


      // PROOF:2 Insert Ranges both column/rows With Right/Down option
      //this function is used to insert rows/column (ranges) proof done
      //we can use Right/Down options
      // var res=await insertColumnOrRowRangeCustom(startRowIndex: 13,startColumnIndex: 2,rowCount: 3,columnCount: 5,sheet: nSheet,);


      //PROOF:1 clear content or cells with formating
      //this function is used to clear values from rows/columns/cells ranges
      //and we can clear (content only or all) or others options also available
      // var res = await clearRangeCustom(
      //   startRowIndex: 13,
      //   startColumnIndex: 2,
      //   rowCount: 2,
      //   columnCount: 2,
      //   sheet: nSheet,
      //   t: RangeClearProperties.All
      // );

      // ExcelTableApi tableApi;
      // tableApi.excelApi=excelApi;

      // RangeModel range = RangeModel.excelRangeModel(rowsCount: 1, columnsCount: 5, topLeftCell: CellModel(columnIndex: 0,rowIndex: 1), relativeTopLeftCell: CellModel(columnIndex: 0,rowIndex: 1), range: "A2:E2");;
      // var list=await tableApi.loadRangeValues(range: range);

      // _counter = " ${sheet.name}\n${res}";
      _counter += " ${sheet.name}\n";
    } catch (e) {
      print(e);
      _counter = "${e}";
    }

    setState(() {});
  }

  @override
  Widget build(BuildContext context) {
    // This method is rerun every time setState is called, for instance as done
    // by the _incrementCounter method above.
    //
    // The Flutter framework has been optimized to make rerunning build methods
    // fast, so that you can just rebuild anything that needs updating rather
    // than having to individually change instances of widgets.
    return Scaffold(
      appBar: AppBar(
        // TRY THIS: Try changing the color here to a specific color (to
        // Colors.amber, perhaps?) and trigger a hot reload to see the AppBar
        // change color while the other colors stay the same.
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        // Here we take the value from the MyHomePage object that was created by
        // the App.build method, and use it to set our appbar title.
        title: Text(widget.title),
      ),
      body: Center(
        // Center is a layout widget. It takes a single child and positions it
        // in the middle of the parent.
        child: SingleChildScrollView(
          child: Column(
            // Column is also a layout widget. It takes a list of children and
            // arranges them vertically. By default, it sizes itself to fit its
            // children horizontally, and tries to be as tall as its parent.
            //
            // Column has various properties to control how it sizes itself and
            // how it positions its children. Here we use mainAxisAlignment to
            // center the children vertically; the main axis here is the vertical
            // axis because Columns are vertical (the cross axis would be
            // horizontal).
            //
            // TRY THIS: Invoke "debug painting" (choose the "Toggle Debug Paint"
            // action in the IDE, or press "p" in the console), to see the
            // wireframe for each widget.
            mainAxisAlignment: MainAxisAlignment.center,
            children: <Widget>[
              const Text(
                'You have pushed the button this many times:',
              ),
              SelectableText(
                '$_counter',
                style: Theme.of(context).textTheme.headlineMedium,
              ),


              SingleChildScrollView(
                child: Column(
                  children: [


                    //current problems
                    Text('Current Bugs',style: TextStyle(color: Colors.red),),
                    //apply download and open
                    TextButton(onPressed: ()async{

                      try{
                        var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                        await excelApi.onLoad();
                        await Excel.copySheetsFromAsset('assets/files/template.xlsx');
                      }catch(error){
                        await showDialog(context: context, builder: (BuildContext) {
                          return SelectableText('Error ${error}');
                        });
                      }



                    }, child: Text('Import From Dart')),
                    SizedBox(height: 10,),
                    TextButton(onPressed: ()async{
                      try{
                        js.context.callMethod('saveCreateAndImportWorkbookFromAssets');
                      }catch(e){
                        print('Error ${e}');
                        await showDialog(context: context, builder: (BuildContext) {
                          return SelectableText('Error ${e}');
                        });

                      }
                    }, child: Text('Calling Pure JS Function')),
                    SizedBox(height: 10,),


                    //current problems
                    Text('Current Functional Features',style: TextStyle(color: Colors.green),),

                    //activate sheet by name
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //PROOF:1 Activate Sheet By Name
                      //code to activate sheet by name
                      var sheets=await excelApi.getSheets();
                      var Sheet1Name=sheets.where((element) => element.name=='Sheet1').first;
                      excelApi.setActiveSheet(Sheet1Name);


                    }, child: Text('Activate Sheet By Name')),
                    SizedBox(height: 10,),

                    //read data from sheet
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //PROOF:2 Read/Get Ranges both rows/columns
                      var res=await getRowData(rowIndex: 13,startColumn: 0,rowCount: 2,columnCount: 5,sheet: nSheet);
                      _counter="${res}";
                      setState(() {

                      });


                    }, child: Text('Read Data From Sheet')),
                    SizedBox(height: 10,),


                    //update sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //PROOF:3 Update Ranges both rows/columns
                      //this is used to update values to specific range
                      List<List<dynamic>> values=[[10, "test", "dummy", "wednesday", 20],[10, "test", 'dummy', "wednesday", 20]];
                      var res=await updateRowData(rowIndex: 13,startColumn: 0,rowCount: 2,columnCount: 5,sheet: nSheet,newValues: values);


                    }, child: Text('Update Data From Sheet')),
                    SizedBox(height: 10,),


                    //insert column sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      // PROOF:4 Insert Ranges both column/rows With Right/Down option
                      //this function is used to insert rows/column (ranges) proof done
                      //we can use Right/Down options
                      var res=await insertColumnOrRowRangeCustom(startRowIndex: 13,startColumnIndex: 2,rowCount: 3,columnCount: 5,sheet: nSheet, insertColumn: true);


                    }, child: Text('Insert Range into Sheet')),
                    SizedBox(height: 10,),

                    //insert row sheet data
                    // TextButton(onPressed: ()async{
                    //
                    //   var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                    //   await excelApi.onLoad();
                    //
                    //   //get active sheet
                    //   SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                    //   ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();
                    //
                    //   // PROOF:4 Insert Ranges both column/rows With Right/Down option
                    //   //this function is used to insert rows/column (ranges) proof done
                    //   //we can use Right/Down options
                    //   var res=await insertColumnOrRowRangeCustom(startRowIndex: 13,startColumnIndex: 2,rowCount: 3,columnCount: 5,sheet: nSheet, insertColumn: false);
                    //
                    //
                    // }, child: Text('Insert Row Range into Sheet')),
                    // SizedBox(height: 10,),


                    //insert complete row sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      // PROOF:4 Insert Ranges both column/rows With Right/Down option
                      //this function is used to insert rows/column (ranges) proof done
                      //we can use Right/Down options
                      // var res=await insertColumnOrRowRangeCustom(startRowIndex: 13,startColumnIndex: 0,rowCount: 1,columnCount: excelMaxColumnsCellsCount-1,sheet: nSheet, insertColumn: true);
                      // var res=await insertColumnOrRowRangeCustom(startRowIndex: 3,startColumnIndex: 0,rowCount: 2,columnCount: excelMaxColumnsCellsCount-1,sheet: nSheet, insertColumn: false);
                      var res=await insertColumnOrRowRangeCustom(startRowIndex: 3,startColumnIndex: 0,rowCount: 2,columnCount: 1000,sheet: nSheet, insertColumn: false);


                    }, child: Text('Insert Full Row')),
                    SizedBox(height: 10,),

                    //insert complete column sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      // PROOF:4 Insert Ranges both column/rows With Right/Down option
                      //this function is used to insert rows/column (ranges) proof done
                      //we can use Right/Down options
                      var res=await insertColumnOrRowRangeCustom(startRowIndex: 0,startColumnIndex: 3,rowCount: excelMaxRowsCellsCount,columnCount: 3,sheet: nSheet, insertColumn: true);


                    }, child: Text('Insert Full Column')),
                    SizedBox(height: 10,),


                    //clear range sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      // PROOF: clear content or cells with formating
                      // this function is used to clear values from rows/columns/cells ranges
                      // and we can clear (content only or all) or others options also available
                      var res = await clearRangeCustom(
                        startRowIndex: 13,
                        startColumnIndex: 2,
                        rowCount: 2,
                        columnCount: 2,
                        sheet: nSheet,
                        t: RangeClearProperties.All
                      );

                    }, child: Text('Clear All')),
                    SizedBox(height: 10,),

                    //clear range sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      // PROOF: clear content or cells with formating
                      // this function is used to clear values from rows/columns/cells ranges
                      // and we can clear (content only or all) or others options also available
                      var res = await clearRangeCustom(
                          startRowIndex: 13,
                          startColumnIndex: 2,
                          rowCount: 2,
                          columnCount: 2,
                          sheet: nSheet,
                          t: RangeClearProperties.Contents
                      );

                    }, child: Text('Clear Content Only')),
                    SizedBox(height: 10,),


                    //delete range sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //PROOF:3 Delete Ranges in progress
                      //delete/remove cell/rows/columns from excel
                      // we have these options (  Up, Left, Down, Right, )
                      var res=await deleteRangeCustom(startRowIndex: 2,startColumnIndex: 2,rowCount: 3,columnCount: 3,sheet: nSheet,side: RangeDeleteProperties.Left);


                    }, child: Text('Delete Range Left ')),
                    SizedBox(height: 10,),

                    //delete range sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //PROOF:3 Delete Ranges in progress
                      //delete/remove cell/rows/columns from excel
                      // we have these options (  Up, Left, Down, Right, )
                      var res=await deleteRangeCustom(startRowIndex: 2,startColumnIndex: 2,rowCount: 3,columnCount: 3,sheet: nSheet,side: RangeDeleteProperties.Up);


                    }, child: Text('Delete Range Up ')),
                    SizedBox(height: 10,),

                    //delete range sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //PROOF:3 Delete Ranges in progress
                      //delete/remove cell/rows/columns from excel
                      // we have these options (  Up, Left, Down, Right, )

                      // startRowIndex: 13,startColumnIndex: 0,rowCount: 1,columnCount: excelMaxColumnsCellsCount-1,
                      //by changing row count i can delete as many rows as i want
                      var res=await deleteRangeCustom(startRowIndex: 2,startColumnIndex: 0,rowCount: 1,columnCount: excelMaxColumnsCellsCountWithSubtract,sheet: nSheet,side: RangeDeleteProperties.Up);


                    }, child: Text('Delete Full Row ')),
                    SizedBox(height: 10,),

                    //delete range sheet data
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //PROOF:3 Delete Ranges in progress
                      //delete/remove cell/rows/columns from excel
                      // we have these options (  Up, Left, Down, Right, )

                      //startRowIndex: 0,startColumnIndex: 3,rowCount: excelMaxRowsCellsCount,columnCount: 3,
                      //by changing row count i can delete as many rows as i want
                      var res=await deleteRangeCustom(startRowIndex: 0,startColumnIndex: 3,rowCount: excelMaxRowsCellsCount,columnCount: 1,sheet: nSheet,side: RangeDeleteProperties.Left);


                    }, child: Text('Delete Full Column')),
                    SizedBox(height: 10,),


                    //apply formatting
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //by changing row count i can delete as many rows as i want
                      var res=await applyCustomFormatting(rowIndex: 13,startColumn: 1,rowCount: 2,columnCount: 5,sheet: nSheet);


                    }, child: Text('Formate Text')),
                    SizedBox(height: 10,),

                    // apply search
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //by changing row count i can delete as many rows as i want
                      var res=await applySearch(rowIndex: 13,startColumn: 1,rowCount: 2,columnCount: 5,sheet: nSheet);


                    }, child: Text('Search Text')),
                    SizedBox(height: 10,),

                    //apply search and activate cell
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //by changing row count i can delete as many rows as i want
                      var res=await applyFindAndActivate(rowIndex: 13,startColumn: 1,rowCount: 2,columnCount: 5,sheet: nSheet);


                    }, child: Text('Search & Activate Text')),
                    SizedBox(height: 10,),


                    //apply search and replace
                    TextButton(onPressed: ()async{

                      var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                      await excelApi.onLoad();

                      //get active sheet
                      SheetModel<dynamic> sheet = await excelApi.getActiveSheet();
                      ExcelSheetModel<Worksheet> nSheet = sheet.toExcelSheetNewCode();

                      //by changing row count i can delete as many rows as i want
                      var res=await applySearchAndReplace(rowIndex: 0,rowCount: excelMaxRowsCellsCount,startColumn: 0,columnCount: excelMaxColumnsCellsCountWithSubtract,sheet: nSheet);


                    }, child: Text('Search and Replace')),
                    SizedBox(height: 10,),


                    //apply download and open
                    TextButton(onPressed: ()async{

                      try{
                        var excelApi = Provider.of<addinHelper.ExcelApiI>(context, listen: false);
                        await excelApi.onLoad();
                        await Excel.copySheetsFromAsset('assets/files/template.xlsx');
                      }catch(error){
                        await showDialog(context: context, builder: (BuildContext) {
                          return SelectableText('Error ${error}');
                        });
                      }



                    }, child: Text('Download And Open')),
                    SizedBox(height: 10,),





                  ],
                ),
              )
            ],
          ),
        ),
      ),
      floatingActionButton: FloatingActionButton(
        onPressed: () async {
          await _incrementCounter(context: context);
        },
        tooltip: 'Increment',
        child: const Icon(Icons.add),
      ), // This trailing comma makes auto-formatting nicer for build methods.
    );
  }
}
