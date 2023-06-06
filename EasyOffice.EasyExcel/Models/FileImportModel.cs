using EasyOffice.EasyExcel.Exceptions;
using Microsoft.AspNetCore.Http;
using NPOI.SS.UserModel;

namespace EasyOffice.EasyExcel.Models;

public class FileImportModel
{
    public FileImportModel(IFormFile file)
    {
        if (file is {Length: > 0})
        {
            using var memoryStream = new MemoryStream();
            file.CopyToAsync(memoryStream).Wait();
            Workbook = WorkbookFactory.Create(memoryStream);
        }
        else
        {
            throw new ExcelFileNotFoundException("No file contained in data");
        }
    }
    public FileImportModel(FileStream file)
    {
        if (file is {Length: > 0})
        {
            using var memoryStream = new MemoryStream();
            file.CopyToAsync(memoryStream).Wait();
            Workbook = WorkbookFactory.Create(memoryStream);
        }
        else
        {
            throw new ExcelFileNotFoundException("No file contained in data");
        }
    }
    public FileImportModel(Stream file)
    {
        if (file is {Length: > 0})
        {
            using var memoryStream = new MemoryStream();
            file.CopyToAsync(memoryStream).Wait();
            Workbook = WorkbookFactory.Create(memoryStream);
        }
        else
        {
            throw new ExcelFileNotFoundException("No file contained in data");
        }
    }


    public IWorkbook Workbook { get; }
}