using EasyOffice.EasyExcel;
using EasyOffice.EasyExcel.Attributes;
using EasyOffice.EasyExcel.Models;


namespace UnitTest;

public class ExcelImport
{

    [Fact]
    public void ShouldImportFromFileStream()
    {
        var directoryPath = Path.Combine("C:\\Users\\Cube360\\RiderProjects\\EasyOffice\\UnitTest" , "ExcelExport");
        var filePath = Path.Combine(directoryPath, "FileName.xlsx");
        var file = File.OpenRead(filePath);
       
        var res = ExcelImport<TImport>.ReadExcel(new FileImportModel(file),"SheetName");
        file.Close();
        // Assert.NotEmpty(res);
    }
    
    [Fact]
    public void ShouldThrowStreamEmpty()
    {
        var directoryPath = Path.Combine("C:\\Users\\Cube360\\RiderProjects\\EasyOffice\\UnitTest" , "ExcelExport");
        var filePath = Path.Combine(directoryPath, "FileName.xlsx");
        var file = File.OpenRead(filePath);
        file.Close();
        var res = ExcelImport<TImport>.ReadExcel(new FileImportModel(file),"SheetName");
        Assert.NotNull(res);
    }
}

public class TImport
{
    [HeaderName("Id")] public int Column1 { get; set; }
    [HeaderName("Name")] public string Column2 { get; set; }= null!;
    [HeaderName("Valid")] public bool? Column3 { get; set; }
    // [HeaderName("Data")] public string[] Column4 { get; set; }
}