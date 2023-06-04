**EasyOffice.EasyExcel**
------------------------

Basic Excel Generation Library for .NET Core 7.0+ and .NET Framework 4.8+ using NPOI Library.


Usage
-----

```csharp

// FileGenerateModel

public class FileGenerateModel
{
    public FileGenerateModel(string fileName)
    {
        FileName = fileName;
        Stream = new MemoryStream();
    }
    public MemoryStream Stream { get; set; }
    public string FileName { get; set; }
}

```

```csharp
// Model class for exporting excel file
public class TObject
{
    [HeaderName("Id")]
    public string Column1 { get; set; }
    [HeaderName("Name")]
    public string Column2 { get; set; }
}

```

```csharp
// Save to provided directory and return FilePath
string filePath = ExcelGenerat<TObject>.GenerateExcel("SheetName",new List<TObject>())
        .ExportIntoDirectory("FileName","DirectoryPath");
```

```csharp
// Returns instance of FileGenerateModel
FileGenerateModel model = ExcelGenerat<TObject>.GenerateExcel("SheetName",new List<TObject>())
        .ExportAsStream("FileName");
    
// Controller Level
[HttpPost("excel-export"), DisableRequestSizeLimit]
public IActionResult ExportExcel()
{
    FileGenerateModel model = ExcelGenerate<TObject>.GenerateExcel("SheetName",new List<TObject>())
            .ExportAsStream("FileName");
    return File(model.Stream.ToArray(),ExportConst.ContentType, model.FileName);
}


```

