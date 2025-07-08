# TableReader
Class library to read table.

![Release(.NET Framework)](https://img.shields.io/badge/release(dotnet_framwork/ClosedXml)-0.1.2.0-blue)  
![Release(.NET Framework)](https://img.shields.io/badge/release(dotnet_framwork/ExcelDataReader)-0.2.2.0-blue)  
![Release(.NET)](https://img.shields.io/badge/release(dotnet/ClosedXml)-0.2.0.1-darkblue)  
![Release(.NET)](https://img.shields.io/badge/release(dotnet/ClosedXml)-0.1.0.0-darkblue)  
![.NET Framework](https://img.shields.io/badge/.NET%20Framework-%3E%3D%204.7.2-red)
![.NET](https://img.shields.io/badge/.NET-8.0-red)  
![ClosedXml(.net framework)](https://img.shields.io/badge/ClosedXml(.net_framework)-0.97.0-lawngreen)
![ClosedXml(.NET)](https://img.shields.io/badge/ClosedXml(.NET)-0.102.0-LightSeaGreen)  
![ExceTableReader(.NET)](https://img.shields.io/badge/ExcelTableReader(.NET)-3.7.0-SpringGreen)
![ExceTableReader(.net framework)](https://img.shields.io/badge/ExcelTableReader(.net%20framework)-3.6.0-YellowGreen)

# ExcelTableReader

**ExcelTableReader** is a C# class library designed to easily read specified tables from Excel files. This library allows you to choose between two open-source libraries, ClosedXml or ExcelDataReader, depending on your environment. It supports both .NET Framework (4.7.2 and later) and .NET (8.0).

## Features

- **Flexible Library Selection**: You can choose to use either ClosedXml or ExcelDataReader for handling Excel files.
  - Uses **ClosedXml v0.97.0** for Excel file processing in .NET Framework environment.
  - Uses **ExcelDataReader v3.6** for Excel file processing in .NET Framework environment,
- **Multi-Environment Support**: Compatible with both .NET Framework 4.7.2+ and .NET 8.0.
- **Simple Interface**: Easily read specified tables from Excel files.

## Usage

### Using ClosedXml (C#)

Here is an example of how to use `TableReader.ClosedXml`. The result is stored in the `table` variable, which is an instance of `System.Data.DataTable` from .NET Framework or .NET:

```csharp
using System.Data;
using TableReader.ClosedXML;
using TableReader.Interface;

string testFilePath = "example.xlsx";
string sheetName = "Sheet1";
string tableName = "SampleTable"; // Specify the table name here

using var stream = new FileStream(testFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite) {
ITableReader reader = new ExcelTableReader(stream, sheetName);
DataTable table = reader.Read(tableName);

// Process the table data
```

### Using ExcelDataReader (C#)

Similarly, using `ExcelDataReader`, the result is also stored in the `table` variable, which is an instance of `System.Data.DataTable`:
```csharp
using System.Data;
using TableReader.ExcelDataReader;
using TableReader.Interface;

System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

string testFilePath = "example.xlsx";
string sheetName = "Sheet1";
string tableName = "SampleTable"; // Specify the table name here

using var stream = new FileStream(testFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
ITableReader reader = new ExcelTableReader(stream, sheetName);
DataTable table = reader.Read(tableName);
// Process the table data
```

## Exampels

This repository includes sample projects demonstrating the use of both TableReader.ClosedXml and TableReader.ExcelDataReader. Navigate to the `Sample` directory to see these examples.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
