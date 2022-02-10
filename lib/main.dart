// ignore_for_file: prefer_const_constructors

import 'dart:io';

import 'package:flutter/material.dart';
import 'package:path/path.dart';
import 'package:excel/excel.dart';
import 'package:flutter/services.dart' show ByteData, rootBundle;
import 'package:requests/requests.dart';

bool glogic = true;
String authHost = "/authenticate";
String sellersHost = "/sellers";
String histName = Requests.getHostname(authHost);
const String login = "@gmail.com";
const String password = "3PJKSGPV";
String bigToken = "";
const String cookie = "XSRF-TOKEN=eyJpdiI6ImlTbjhkaERuZDd3cDJmMXRMeWVxZGc9PSIsInZhbHVlIjoiNXZMXC9odWxReitDZisxcXZwQStYc1NmRVZ6T2JSbDhDbWlqengxZVVleWhVQmdUMlwvVVJEMFRIRFFFeEdDbWdnIiwibWFjIjoiYmYxMTNkMzY1NWM3MDU5Yjc2ZTZjYmVkYjQzMmUyNWY1MDIxNzVhOGMyNjE4ZDVhYTUwNGFlNjM3ZDNiMGU0NSJ9; myglo_session=eyJpdiI6IkVadExudGpXSnl6S3Nvcko1VXlZS0E9PSIsInZhbHVlIjoib2kxcis3aU1jRHNEaThJNWJhMFZcL000M2liMWxTZWxLaDdcL1RibGlFYWs3VkZuXC9iZmNaRnREYk9CclhXOUpKQyIsIm1hYyI6IjgwNzU0YjZiMTI4YWJmNGQzM2IwZmUwMzJhYzcxNmY3MDhkZjk0MWFmMTk0ZTk1NDMyZWQ3NTg5MTBjMTEzYjgifQ%3D%3D";
void main() {
  runApp(MyApp());
}

bool isOk = true;
var excel = Excel.createExcel();
int counter = 1;
Sheet sheet = excel["Filtred"];

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: Scaffold(
        backgroundColor: Colors.indigo,
        appBar: AppBar(
          title: Text('Sheets cleaner'),
          centerTitle: true,
        ),
        body: HomePage(),
      ),
    );
  }
}

class HomePage extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Column(
        crossAxisAlignment: CrossAxisAlignment.center,
        mainAxisAlignment: MainAxisAlignment.center,
        children: [ButtonOne(), ButtonTwo()]);
  }
}

class ButtonOne extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Center(
      child: ElevatedButton(
        onPressed: () {_prepareApi();},
        child: Text('Prepare api (press twice)'),
      ),
    );
  }
}

class ButtonTwo extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Center(
      child: ElevatedButton(
        onPressed: () {
          doWork();
        },
        child: Text('Lets filter our data Excel'),
      ),
    );
  }
}

Future<void> _prepareApi() async {

  Response r2 = await Requests.get(sellersHost);
  _setTokenFromRequest(r2.content());
  
  Response r3 = await Requests.post(authHost, body: {
    'email': login,
    'password': password,
    '_token': bigToken
  }, headers: {
    'cookie':
    cookie
  });
}

void _setTokenFromRequest(var request) {
  String token = request.toString().substring(152, 192);
  bigToken = token;
  debugPrint("Наш токен - " '$bigToken');
}

Future<void> doWork() async {
  ByteData data = await rootBundle.load('assets/file/leads.xlsx');
  var bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
  var excel = Excel.decodeBytes(bytes);
  for (var table in excel.tables.keys) {
    for (var row in excel.tables[table]!.rows) {
      await Future.delayed(Duration(seconds: 1));
      String newRow = row.toString();
       try {
         _splitString(newRow);
       } catch (e){
         debugPrint(e.toString());
       }
    }
  }
}

void _splitString(String row) {
  String name;
  String number;
  String newRow =
      row.replaceAll("[", "").replaceAll("]", "").replaceAll(" ", "");
  name = newRow.split(",")[0];
  number = newRow.split(",")[1];
  _testUsingApi(name, number);
}

void _testUsingApi(String name, String number) {
  String pretty = _getPrettyString(number);
  _searchNumber(pretty, name, number);

}

Future<void> _searchNumber(String number, String name, String normNumber) async {
  Response r2 = await Requests.get("/order/status?phone="+number);
  if (r2.content().length.toString() == "8656") {
    writeFile(name, normNumber);
    print("Подходит - " + name + " " + normNumber);
  } else {
    print("Не подходит - " + name + " " + normNumber);
  }
  }


String _getPrettyString(String number) {
  String first = "";
  String second = "";
  String third = "";
  String fourth = "";
  String shortNumber = number.substring(4, 13);
  first = shortNumber.substring(0,2);
  second = shortNumber.substring(2,5);
  third = shortNumber.substring(5,7);
  fourth = shortNumber.substring(7,9);
  return("%2B38%280"+first+"%29-"+second+"-"+third+"-"+fourth);
}


void writeFile(String name, number) {
    var name1 = sheet.cell(CellIndex.indexByString("A${counter}"));
    name1.value = name;
    var number1 = sheet.cell(CellIndex.indexByString("B${counter}"));
    number1.value = number;
    counter++;
    _saveFile();

}

void _saveFile() {
  print("Сохраняю");
  excel.encode().then((onValue) {
    File(join("/storage/emulated/0/Download/output.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });
}

//excel.tables[table]!.rows.length /*возвращает количество рядов*/
