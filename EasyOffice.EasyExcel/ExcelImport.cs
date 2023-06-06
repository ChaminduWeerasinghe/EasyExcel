using System.Reflection;
using EasyOffice.EasyExcel.Attributes;
using EasyOffice.EasyExcel.Exceptions;
using EasyOffice.EasyExcel.Models;
using NPOI.SS.UserModel;

namespace EasyOffice.EasyExcel;

public sealed class ExcelImport<TObj> where TObj : class
{

    // public static List<TObj> ReadReadFromFormFile(FileImportModel fileImportModel,string sheetName)
    // {
    //     var workBook = WorkbookAccessor.GetWorkBookFromFile(fileImportModel.FormFile);
    //     return RetrieveExcelData(workBook, sheetName);
    // }
    //
    // public static List<TObj> ReadExcel(string filePath,string sheetName)
    // {
    //     var workBook = WorkbookAccessor.GetWorkBookByFilePath(filePath);
    //     return RetrieveExcelData(workBook, sheetName);
    // }
    //
    // public static List<TObj> ReadFromFileStream(FileImportModel fileImportModel,string sheetName)
    // {
    //     var workBook = WorkbookAccessor.GetWorkBookFromStream(fileImportModel.FileStream);
    //     return RetrieveExcelData(workBook, sheetName);
    // }
    
    public static List<TObj> ReadExcel(FileImportModel fileImportModel,string sheetName)
    {
        return RetrieveExcelData(fileImportModel.Workbook, sheetName);
    }
    
    public static List<TObj> ReadExcel(string filePath,string sheetName)
    {
        var workBook = GetWorkBookByFilePath(filePath);
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
            IRow curRow;
            try
            {
                curRow = (IRow)rows.Current!;
            }
            catch (Exception)
            {
                continue;
            }
            

            if (mapping.Count > 0)
            {
                var insertObject = Activator.CreateInstance<TObj>();
                foreach (var cell in curRow.Cells)
                {
                    if (cell.CellType is CellType.Blank) continue;
                    if (!mapping.TryGetValue(cell.ColumnIndex, out var propertyInfo)) continue;

                    var cellValue = cell.ToString();
                    
                    var value = ValueConverter(propertyInfo,cellValue);
                    
                    propertyInfo.SetValue(insertObject,value);
                }
                insertObjectList.Add(insertObject);
            }
            else
            {
                mapping =  GetRowIndexPropertyInfoMapping(sheet.GetRow(curRow.RowNum));
            }
        }

        return insertObjectList;
    }


    private static object? ValueConverter(PropertyInfo propertyInfo,string? cellValue)
    {
        var propertyInfoPropertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
        return (cellValue == null) ? null : Convert.ChangeType(cellValue, propertyInfoPropertyType);
    }

    private static Dictionary<int, PropertyInfo> GetRowIndexPropertyInfoMapping(IRow row)
    {
        var map = new Dictionary<int, PropertyInfo>();
        var properties = typeof(TObj).GetProperties().ToList();
        foreach (var cell in row.Cells)
        {
            var cellValue = cell.ToString();

            if (!string.IsNullOrWhiteSpace(cellValue))
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
    private static IWorkbook GetWorkBookByFilePath(string filePath)
    {
        if (!File.Exists(filePath))
            throw new ExcelFileNotFoundException("Excel file not found in provided path");
        
        var task = File.ReadAllBytesAsync(filePath);
        task.Wait();
        var byteArray = task.Result;
        return WorkbookFactory.Create(new MemoryStream(byteArray));
    }
}



