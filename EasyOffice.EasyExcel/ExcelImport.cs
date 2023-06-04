using System.Reflection;
using EasyOffice.EasyExcel.Attributes;
using EasyOffice.EasyExcel.Exceptions;
using EasyOffice.EasyExcel.Models;
using Microsoft.AspNetCore.Http;
using NPOI.SS.UserModel;

namespace EasyOffice.EasyExcel;

public sealed class ExcelImport<TObj> where TObj : class
{

    public static List<TObj> ReadExcel(FileImportModel file,string sheetName)
    {
        var workBook = WorkbookAccessor.GetWorkBookByFile(file.File);
        return RetrieveExcelData(workBook, sheetName);
    }

    public static List<TObj> ReadExcel(string filePath,string sheetName)
    {
        var workBook = WorkbookAccessor.GetWorkBookByFilePath(filePath);
        return RetrieveExcelData(workBook, sheetName);
    }

    private static List<TObj> RetrieveExcelData(IWorkbook workbook, string sheetName)
    {
        var sheet = workbook.GetSheet(sheetName);
        var rows = sheet.GetRowEnumerator();
        var mapping = new Dictionary<int, PropertyInfo>();
        var insertObjectList = new List<TObj>();

        while (rows.MoveNext())
        {
            if (rows.Current is null) continue;
            
            var curRow = (IRow)rows.Current;

            if (mapping.Count > 0)
            {
                var insertObject = Activator.CreateInstance<TObj>();
                foreach (var cell in curRow.Cells)
                {
                    if (cell.CellType is CellType.Blank) continue;
                    if (!mapping.TryGetValue(cell.ColumnIndex, out var propertyInfo)) continue;

                    var cellValue = cell.ToString();

                    var value = Convert.ChangeType(cellValue ?? default, propertyInfo.PropertyType);

                    propertyInfo.SetValue(insertObject,value );
                }
                insertObjectList.Add(insertObject);
            }
            else if (curRow.Cells.All(x => x.CellType != CellType.Blank))
            {
                mapping =  GetRowIndexPropertyInfoMapping(sheet.GetRow(curRow.RowNum));
            }
        }

        return insertObjectList;
    }

    private static Dictionary<int, PropertyInfo> GetRowIndexPropertyInfoMapping(IRow row)
    {
        var map = new Dictionary<int, PropertyInfo>();
        var properties = typeof(TObj).GetProperties().ToList();
        foreach (var cell in row.Cells)
        {
            var cellValue = cell.ToString();

            if (cellValue != null)
            {
                foreach (var property in properties)
                {
                    var propertyName = property.Name;
                    var propertyAttributeName = property.GetCustomAttribute<HeaderName>()?.Name;
                    
                    if (string.Equals(cellValue, propertyName, StringComparison.CurrentCultureIgnoreCase) ||
                        string.Equals(cellValue, propertyAttributeName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        map.Add(cell.ColumnIndex, property);
                    }
                }
            }
        }

        return map;
    }
}

public sealed class WorkbookAccessor
{
    
    internal static IWorkbook GetWorkBookByFile(IFormFile file)
    {
        if (file.Length > 0)
        {
            using var memoryStream = new MemoryStream();
            var task = file.CopyToAsync(memoryStream);
            task.Wait();

            try
            {
                return WorkbookFactory.Create(memoryStream);
            }
            catch (Exception)
            {
                throw new ExcelFileNotFoundException("Excel file contained in file");
            }
        }
        throw new ExcelFileNotFoundException("No data contained in file");
    }
    
    internal static IWorkbook GetWorkBookByFilePath(string filePath)
    {
        if (!File.Exists(filePath))
            throw new ExcelFileNotFoundException("Excel file not found in provided path");
        
        var task = File.ReadAllBytesAsync(filePath);
        task.Wait();
        var byteArray = task.Result;
        return WorkbookFactory.Create(new MemoryStream(byteArray));
    }
}


