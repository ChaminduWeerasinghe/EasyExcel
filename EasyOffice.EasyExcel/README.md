# EasyOffice.EasyExcel


Basic Excel Export and Import Library for .NET Core 7.0+ using NPOI Library.


## Usage

>Model is used for export and import data from excel. Maintain property names as same as excel 
 column name or use **HeaderName** attribute for mapping excel column with model property.
> Make sure to keep properties as public and have getter and setter.

```csharp
// Model
public class TObject
{
    [HeaderName("Id")] // HeaderName attribute is optional
    public string Column1 { get; set; }
    public string Name { get; set; }
}

```

### Excel Export

```csharp
// FileGenerateData

public class FileGenerateData
{
    public FileGenerateData(string fileName)
    {
        FileName = fileName;
        Stream = new MemoryStream();
    }
    public MemoryStream Stream { get;}
    public string FileName { get;}
}
```

```csharp
// Save to provided directory and return FilePath
string filePath = ExcelExport<TObject>.GenerateExcel("SheetName",new List<TObject>())
        .ExportIntoDirectory("FileName","DirectoryPath");

// Returns instance of FileGenerateData
FileGenerateData data = ExcelExport<TObject>.GenerateExcel("SheetName",new List<TObject>())
        .ExportAsStream("FileName");
```

```csharp
// Controller Level
[HttpPost("excel-export"), DisableRequestSizeLimit]
public IActionResult ExportExcel()
{
    FileGenerateData data = ExcelExport<TObject>.GenerateExcel("SheetName",new List<TObject>())
            .ExportAsStream("FileName");
    return File(model.Stream.ToArray(),ExportConst.ContentType, model.FileName);
}

```

---
### Excel Import

By using ImportOption.ImportFrom() method you can import excel data from File Path,IFormFile or Stream.

```csharp
// ImportOption
ImportOption importOption = ImportOption.ImportFrom("FilePath");
ImportOption importOption = ImportOption.ImportFrom(file);
ImportOption importOption = ImportOption.ImportFrom(new Stream());
```

**ReadExcel or ReadExcelInSafe**

These methods are the primary methods for importing data from excel. Both methods return List of TObject.

> **ReadExcel** - Read xlsx file and map data to List of TObject. Throws _InvalidValueException_ exception when try to map blank value in excel to not nullable property

> **ReadExcelInSafe** - This is more safe way of read and map data to List of TObject. Set default value when try to map blank value in excel to not nullable property.
```csharp
// ReadExcel
List<TObject> data = ExcelImport<TObject>.ReadExcel(ImportOption.ImportFrom(filePath), "SheetName");

// // ReadExcelInSafe
List<TObject> data = ExcelImport<TObject>.ReadExcelInSafe(ImportOption.ImportFrom(filePath), "SheetName");
```