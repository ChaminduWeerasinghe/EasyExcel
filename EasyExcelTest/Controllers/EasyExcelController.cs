using EasyOffice.EasyExcel;
using EasyOffice.EasyExcel.Attributes;
using EasyOffice.EasyExcel.Models;
using Microsoft.AspNetCore.Mvc;

namespace EasyExcelTest.Controllers;

[ApiController]
[Route("[controller]")]
public class EasyExcelController  : ControllerBase
{

    private readonly string _fileSaveDir;

    private readonly List<TestObject> _summaries = new()
    {
        new TestObject {Column1 = "1", Name = "ABC"},
        new TestObject {Column1 = "2", Name = "CDE"},
        new TestObject {Column1 = "3", Name = "EFG"},
        new TestObject {Column1 = "4", Name = "HIJ"},
    };
    
    private readonly List<TestObject2> _summaries2 = new()
    {
        new TestObject2(1, "ALL AVAILABLE", true, 4),
        new TestObject2(2, null, true, null),
        new TestObject2(3, "1 NULL", null, 4),
        new TestObject2(4, null, null, null),
    };


    public EasyExcelController(IWebHostEnvironment hostEnvironment)
    {
        _fileSaveDir = Path.Join(hostEnvironment.ContentRootPath, "ExcelFiles");
    }

    [HttpPost("export/into-stream")]
    public ActionResult ExportExcel()
    {
        // var fileGenerateModel = ExcelExport<TestObject>.GenerateExcel("SheetName", _summaries).ExportAsStream("FileName");
        var fileGenerateModel = ExcelExport<TestObject2>.GenerateExcel("SheetName", _summaries2).ExportAsStream("Test");
        return File(fileGenerateModel.Stream.ToArray(), ExcelConstant.ContentType, fileGenerateModel.FileName);
    }
    
    [HttpPost("export/into-directory")]
    public ActionResult ExportExcelIntoDirectory()
    {
        // var fileGenerateModel = ExcelExport<TestObject>.GenerateExcel("SheetName", _summaries).ExportIntoDirectory("FileName",_fileSaveDir);
        var fileGenerateModel = ExcelExport<TestObject2>.GenerateExcel("SheetName", _summaries2).ExportIntoDirectory("Test",_fileSaveDir);
        return Ok(fileGenerateModel);
    }
    
    [HttpPost("import/throw/from-file")]
    public ActionResult ImportExcel(IFormFile file)
    {
        // var fileGenerateModel = ExcelImport<TestObject>.ReadExcel(ImportOption.ImportFrom(file), "SheetName");
        // var fileGenerateModel = ExcelImport<TestObject2>.ReadExcel(ImportOption.ImportFrom(file), "SheetName");
        var fileGenerateModel = ExcelImport<TestObject3>.ReadExcel(ImportOption.ImportFrom(file), "SheetName");
        return Ok(fileGenerateModel);
    }
    
    [HttpPost("import/throw/from-directory")]
    public ActionResult ImportExcelFromDirectory()
    {
        var filePath = Path.Join(_fileSaveDir, "Test.xlsx");
        // var fileGenerateModel = ExcelImport<TestObject>.ReadExcel(ImportOption.ImportFrom(filePath), "SheetName");
        // var fileGenerateModel = ExcelImport<TestObject2>.ReadExcel(ImportOption.ImportFrom(filePath), "SheetName");
        var fileGenerateModel = ExcelImport<TestObject3>.ReadExcel(ImportOption.ImportFrom(filePath), "SheetName");
        return Ok(fileGenerateModel);
    }
    
    [HttpPost("import/default/from-file")]
    public ActionResult ImportExcelOrDefault(IFormFile file)
    {
        // var fileGenerateModel = ExcelImport<TestObject>.ReadExcel(ImportOption.ImportFrom(file), "SheetName");
        // var fileGenerateModel = ExcelImport<TestObject2>.ReadExcel(ImportOption.ImportFrom(file), "SheetName");
        var fileGenerateModel = ExcelImport<TestObject3>.ReadExcelInSafe(ImportOption.ImportFrom(file), "SheetName");
        return Ok(fileGenerateModel);
    }
    
    [HttpPost("import/default/from-directory")]
    public ActionResult ImportExcelOrDefaultFromDirectory()
    {
        var filePath = Path.Join(_fileSaveDir, "Test.xlsx");
        // var fileGenerateModel = ExcelImport<TestObject>.ReadExcel(ImportOption.ImportFrom(filePath), "SheetName");
        // var fileGenerateModel = ExcelImport<TestObject2>.ReadExcel(ImportOption.ImportFrom(filePath), "SheetName");
        var fileGenerateModel = ExcelImport<TestObject3>.ReadExcelInSafe(ImportOption.ImportFrom(filePath), "SheetName");
        return Ok(fileGenerateModel);
    }
}

internal class TestObject
{
    [HeaderName("Id")]
    public string Column1 { get; set; } = null!;
    public string Name { get; set; } = null!;
}

internal class TestObject2
{
    public TestObject2()
    {
        
    }
    public TestObject2(int id, string? name, bool? valid, int? column4)
    {
        Column1 = id;
        Name = name;
        Valid = valid;
        Column4 = column4;
    }

    [HeaderName("Id")] public int Column1 { get; set;}
    public string? Name { get; set;}
    public bool? Valid { get; set;}
    [HeaderName("Count")] public int? Column4 { get; set;}
}

internal class TestObject3
{
    [HeaderName("Id")] public int Column1 { get; set;}
    public string Name { get; set; } = null!;
    public bool Valid { get; set;}
}
