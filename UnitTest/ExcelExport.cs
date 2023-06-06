using EasyOffice.EasyExcel;
using EasyOffice.EasyExcel.Attributes;

namespace UnitTest;

public class ExcelExport
{
    
    private readonly List<TExport> _summaries = new List<TExport>
    {
        new TExport {Column1 = "1", Column2 = "ABC"},
        new TExport {Column1 = "2", Column2 = "CDE"},
        new TExport {Column1 = "3", Column2 = "EFG"},
        new TExport {Column1 = "4", Column2 = "HIJ"},
    };
    
    [Fact]
    public void ShouldExportAsStream()
    {
        var res = ExcelExport<TExport>.GenerateExcel("SheetName",_summaries).ExportAsStream("FileName");
        Assert.NotNull(res);
        Assert.NotNull(res.Stream);
        Assert.NotNull(res.FileName);
    }
    
    [Fact]
    public void ShouldExportIntoDirectory()
    {
        var directoryPath = Path.Combine("C:\\Users\\Cube360\\RiderProjects\\EasyOffice\\UnitTest" , "ExcelExport");
        var res = ExcelExport<TExport>.GenerateExcel("SheetName",_summaries).ExportIntoDirectory("FileName",directoryPath);
        Assert.NotNull(res);
    }
}


public class TExport
{
    [HeaderName("Id")] public string Column1 { get; set; } = null!;
    [HeaderName("Name")] public string Column2 { get; set; }= null!;
}