import 'dart:ui';
import 'dart:async';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
// ignore: import_of_legacy_library_into_null_safe
import 'package:excel/excel.dart';
import 'package:path/path.dart';
import 'package:open_file/open_file.dart';

void main() {
  runApp(MaterialApp(
    home: Home(),
  ));
}

class Home extends StatefulWidget {
  @override
  _HomeState createState() => _HomeState();
}

class _HomeState extends State<Home> {
  String data = '';
  late Excel excelfile;
  List<String> nameList = [];
  int positionCounter = 0;
  int date = 0;

  bool nodate() {
    if (date == 0) {
      return true;
    } else {
      return false;
    }
  }

  Future<String> get _localPath async {
    final directory = await getExternalStorageDirectory();

    return directory!.path;
  }

  Future<File> get _localFile async {
    final path = await _localPath;
    return File('$path/2021.xlsx');
  }

  Future<Excel> readContent() async {
    var excel = Excel.createExcel();
    try {
      final file = await _localFile;

      // Read the file
      final contents = file.readAsBytesSync();
      // Returning the contents of file

      excel = Excel.decodeBytes(contents);
      return excel;
    } catch (e) {
      // if encountering an error, return
      return excel;
    }
  }

  // saving the file
  void writeContent({required Excel file}) async {
    // write the file
    final path = await _localPath;
    file.encode().then((onvalue) {
      File(join('$path/2021.xlsx'))
        ..createSync(recursive: true)
        ..writeAsBytesSync(onvalue);
    });
  }

  List<String> returnNameList({required Excel file}) {
    List<String> nameList = [];

    var nameIndex = 1;
    var cellName;

    Sheet sheetObject = file['Sheet1'];

    while (true) {
      cellName = sheetObject.cell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: nameIndex));

      if (cellName.value == null) {
        break;
      }
      nameList.add(cellName.value);
      nameIndex++;
    }

    return nameList;
  }

  List<int> returnDateList({required Excel file}) {
    List<int> dateList = [];
    var dateIndex = 1;
    var cellDate;

    Sheet sheetObject = file['Sheet1'];

    while (true) {
      cellDate = sheetObject.cell(
          CellIndex.indexByColumnRow(columnIndex: dateIndex, rowIndex: 0));

      if (cellDate.value == null) {
        break;
      }
      dateList.add(cellDate.value);
      dateIndex++;
    }

    return dateList;
  }

  // inserting the name
  Excel insertName({required Excel file, required String name}) {
    var i = 1;
    var cell;
    List<String> nameList = returnNameList(file: file);
    for (int i = 0; i < nameList.length; i++) {
      if (name == nameList[i] || name == '') {
        return file;
      }
    }
    Sheet sheetObject = file['Sheet1'];
    while (true) {
      cell = sheetObject
          .cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i));
      if (cell.value == null) {
        cell.value = name;
        break;
      }
      i++;
    }
    return file;
  }

  // inserting the name
  Excel insertDate({required Excel file, required int date}) {
    // Assuming the data is not there and inserting the value at the first null occured
    var i = 1;
    var cell;
    List<int> dateList = returnDateList(file: file);
    for (int i = 0; i < dateList.length; i++) {
      if (date == dateList[i]) {
        return file;
      }
    }
    Sheet sheetObject = file['Sheet1'];

    while (true) {
      cell = sheetObject
          .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0));
      if (cell.value == null) {
        cell.value = date;
        break;
      }
      i++;
    }
    return file;
  }

  // ignore: non_constant_identifier_names
  Excel InsertValue(
      {required int date,
      required String name,
      required Excel file,
      required int value}) {
    /* Appending all name and date to a list
      to get the index we want 
    */

    // ignore: non_constant_identifier_names
    List<String> NameList = [];
    // ignore: non_constant_identifier_names
    List<int> DateList = [];

    var nameIndex = 1;
    var dateIndex = 1;
    var cellName;
    var cellDate;
    var cell;

    Sheet sheetObject = file['Sheet1'];

    while (true) {
      cellName = sheetObject.cell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: nameIndex));

      if (cellName.value == null) {
        break;
      }
      NameList.add(cellName.value);
      nameIndex++;
    }

    while (true) {
      cellDate = sheetObject.cell(
          CellIndex.indexByColumnRow(columnIndex: dateIndex, rowIndex: 0));

      if (cellDate.value == null) {
        break;
      }
      DateList.add(cellDate.value);
      dateIndex++;
    }

    // finding the index of name
    for (var i = 0; i < NameList.length; i++) {
      if (name == NameList[i]) {
        nameIndex = i;
      }
    }

    // finding the index of date
    for (var i = 0; i < DateList.length; i++) {
      if (date == DateList[i]) {
        dateIndex = i;
      }
    }

    // adding the value
    cell = sheetObject.cell(CellIndex.indexByColumnRow(
        columnIndex: dateIndex + 1, rowIndex: nameIndex + 1));
    cell.value = value;

    return file;
  }

  Widget popupDialog(BuildContext context) {
    return AlertDialog(
      title: const Text(
        'Enter new name',
        style: TextStyle(color: Colors.grey),
      ),
      backgroundColor: Colors.amber,
      content: Column(
        mainAxisSize: MainAxisSize.min,
        crossAxisAlignment: CrossAxisAlignment.start,
        children: <Widget>[
          TextField(
            cursorColor: Colors.grey,
            decoration:
                InputDecoration(labelStyle: TextStyle(color: Colors.white)),
            onSubmitted: (value) {
              setState(() {
                String name = value;
                insertName(file: excelfile, name: name);
                writeContent(file: excelfile);
              });
            },
          )
        ],
      ),
      actions: <Widget>[
        ElevatedButton(
          onPressed: () {
            Navigator.of(context).pop();
          },
          style: ElevatedButton.styleFrom(
            onPrimary: Colors.deepOrange,
            primary: Colors.grey,
          ),
          child: const Text(
            'Close',
            style: TextStyle(
              color: Colors.amber,
            ),
          ),
        )
      ],
    );
  }

  @override
  void initState() {
    super.initState();
    readContent().then((value) {
      setState(() {
        excelfile = value;
      });
    });
  }

  @override
  Widget build(BuildContext context) {
    nameList = returnNameList(file: excelfile);
    if (nameList.length == 0) {
      nameList.add("No name added");
    }
    return Scaffold(
      backgroundColor: Colors.grey[800],
      appBar: AppBar(
        elevation: 0.0,
        centerTitle: true,
        title: Text(
          "OFFICE WORK",
          style: TextStyle(color: Colors.amber, fontSize: 30.0),
        ),
        backgroundColor: Colors.grey[900],
      ),
      body: SingleChildScrollView(
        child: Column(
          children: <Widget>[
            Container(
              padding: EdgeInsets.fromLTRB(35.0, 35.0, 35.0, 0.0),
              child: Text(
                "${nameList[positionCounter]}",
                style: TextStyle(fontSize: 40.0, color: Colors.amber[600]),
              ),
            ),
            Row(
              mainAxisAlignment: MainAxisAlignment.end,
              children: <Widget>[
                // add user button
                IconButton(
                    onPressed: () {
                      showDialog(
                        context: context,
                        builder: (BuildContext context) => popupDialog(context),
                      );
                    },
                    iconSize: 35.0,
                    icon: Icon(Icons.manage_accounts))
              ],
            ),
            SizedBox(
              height: 10.0,
            ),
            // input goes here
            Padding(
              padding:
                  const EdgeInsets.symmetric(horizontal: 120.0, vertical: 5.0),
              child: TextField(
                cursorColor: Colors.amber,
                keyboardType: TextInputType.number,
                maxLength: 2,
                style: TextStyle(
                  color: Colors.grey[400],
                  fontSize: 20.0,
                ),
                decoration: InputDecoration(
                  labelStyle: TextStyle(color: Colors.grey[400]),
                  border: OutlineInputBorder(
                    borderRadius: BorderRadius.circular(30.0),
                  ),
                  labelText: 'Date',
                ),
                onSubmitted: (value) {
                  setState(() {
                    date = int.parse(value);
                    insertDate(file: excelfile, date: date);
                    writeContent(file: excelfile);
                  });
                },
              ),
            ),
            Column(
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    Padding(
                      padding: const EdgeInsets.fromLTRB(40.0, 0.0, 40.0, 0.0),
                      child: ElevatedButton(
                        onPressed: nodate()
                            ? null
                            : () {
                                setState(() {
                                  InsertValue(
                                      date: date,
                                      name: nameList[positionCounter],
                                      file: excelfile,
                                      value: 8);
                                  writeContent(file: excelfile);
                                });
                              },
                        child: Text(
                          '08',
                          style: TextStyle(
                              fontSize: 35.0, color: Colors.grey[400]),
                        ),
                        style: ElevatedButton.styleFrom(
                          onPrimary: Colors.black87,
                          primary: Colors.grey[700],
                          minimumSize: Size(88, 36),
                          padding: EdgeInsets.all(35.0),
                          shape: const RoundedRectangleBorder(
                            borderRadius: BorderRadius.all(Radius.circular(2)),
                          ),
                        ),
                      ),
                    ),
                    Padding(
                      padding: const EdgeInsets.fromLTRB(0.0, 0.0, 40.0, 0.0),
                      child: ElevatedButton(
                        onPressed: nodate()
                            ? null
                            : () {
                                setState(() {
                                  InsertValue(
                                      date: date,
                                      name: nameList[positionCounter],
                                      file: excelfile,
                                      value: 16);
                                  writeContent(file: excelfile);
                                });
                              },
                        child: Text(
                          '16',
                          style: TextStyle(
                              fontSize: 35.0, color: Colors.grey[400]),
                        ),
                        style: ElevatedButton.styleFrom(
                          onPrimary: Colors.black87,
                          primary: Colors.grey[700],
                          minimumSize: Size(88, 36),
                          padding: EdgeInsets.all(35.0),
                          shape: const RoundedRectangleBorder(
                            borderRadius: BorderRadius.all(Radius.circular(2)),
                          ),
                        ),
                      ),
                    ),
                  ],
                ),
                SizedBox(
                  height: 10.0,
                ),
                // Nil Button
                Padding(
                  padding: const EdgeInsets.fromLTRB(0.0, 10.0, 0.0, 10.0),
                  child: ElevatedButton(
                    onPressed: nodate()
                        ? null
                        : () {
                            setState(() {
                              InsertValue(
                                  date: date,
                                  name: nameList[positionCounter],
                                  file: excelfile,
                                  value: 0);
                              writeContent(file: excelfile);
                            });
                          },
                    child: Text(
                      'Nil',
                      style: TextStyle(fontSize: 35.0, color: Colors.grey[400]),
                    ),
                    style: ElevatedButton.styleFrom(
                      onPrimary: Colors.black87,
                      primary: Colors.grey[700],
                      minimumSize: Size(100, 100),
                      padding: EdgeInsets.fromLTRB(0.0, 0.0, 0.0, 0.0),
                      shape: const RoundedRectangleBorder(
                        borderRadius: BorderRadius.all(Radius.circular(2)),
                      ),
                    ),
                  ),
                )
              ],
            ),
            SizedBox(
              height: 20.0,
            ),
            // next and back button
            Row(
              mainAxisAlignment: MainAxisAlignment.center,
              children: <Widget>[
                Padding(
                  padding: const EdgeInsets.fromLTRB(10.0, 0.0, 10.0, 0.0),
                  child: ElevatedButton(
                    onPressed: () {
                      setState(() {
                        if (positionCounter > 0) {
                          positionCounter--;
                        }
                      });
                    },
                    child: Text(
                      "<< Back",
                      style: TextStyle(color: Colors.grey[400], fontSize: 20.0),
                    ),
                    style: ElevatedButton.styleFrom(
                        onPrimary: Colors.black87,
                        primary: Colors.grey[700],
                        minimumSize: Size(150, 36),
                        padding: EdgeInsets.fromLTRB(20, 20, 20, 20),
                        shape: const RoundedRectangleBorder(
                          borderRadius: BorderRadius.all(Radius.circular(2)),
                        )),
                  ),
                ),
                Padding(
                  padding: const EdgeInsets.fromLTRB(0.0, 0.0, 10.0, 0.0),
                  child: ElevatedButton(
                    onPressed: () {
                      setState(() {
                        if (positionCounter < (nameList.length - 1)) {
                          positionCounter++;
                        }
                      });
                    },
                    child: Text(
                      "Next >>",
                      style: TextStyle(fontSize: 20.0, color: Colors.grey[400]),
                    ),
                    style: ElevatedButton.styleFrom(
                        onPrimary: Colors.black87,
                        primary: Colors.grey[700],
                        minimumSize: Size(150, 36),
                        padding: EdgeInsets.fromLTRB(20, 20, 20, 20),
                        shape: const RoundedRectangleBorder(
                          borderRadius: BorderRadius.all(Radius.circular(2)),
                        )),
                  ),
                ),
              ],
            ),
            SizedBox(height: 10.0),
            Row(
              mainAxisAlignment: MainAxisAlignment.center,
              children: <Widget>[
                // share button
                Padding(
                    padding: EdgeInsets.fromLTRB(0.0, 10.0, 10.0, 10.0),
                    child: ElevatedButton(
                        onPressed: () async {
                          writeContent(file: excelfile);
                          final path = await _localPath;
                          OpenFile.open('$path/2021.xlsx');
                        },
                        style: ElevatedButton.styleFrom(
                            onPrimary: Colors.deepOrange,
                            primary: Colors.amber,
                            minimumSize: Size(20, 20),
                            padding: EdgeInsets.all(20.0),
                            shape: const RoundedRectangleBorder(
                                borderRadius:
                                    BorderRadius.all(Radius.circular(2)))),
                        child: Text(
                          "Open & Share",
                          style: TextStyle(
                              color: Colors.grey[800],
                              fontWeight: FontWeight.bold),
                        )))
              ],
            ),
          ],
        ),
      ),
    );
  }
}
