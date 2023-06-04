using Microsoft.AspNetCore.Http;

namespace EasyOffice.EasyExcel.Models;

public class FileImportModel
{
    public FileImportModel(IFormFile file)
    {
        File = file;
    }

    public IFormFile File { get;}
}