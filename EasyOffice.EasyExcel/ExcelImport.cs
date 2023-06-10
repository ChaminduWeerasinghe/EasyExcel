using System.Reflection;
using EasyOffice.EasyExcel.Attributes;
using EasyOffice.EasyExcel.Exceptions;
using NPOI.SS.UserModel;
using ImportOption = EasyOffice.EasyExcel.Models.ImportOption;

namespace EasyOffice.EasyExcel;

public sealed class ExcelImport<TObj> where TObj : class
{

    
    public static List<TObj> ReadExcel(ImportOption importOption,string sheetName)
    {
        return RetrieveExcelData(importOption.Workbook, sheetName);
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
                foreach (var cell in curRow.Cells.Where(cell => cell.CellType is not CellType.Blank))
                {
                    if (!mapping.TryGetValue(cell.ColumnIndex, out var propertyInfo)) continue;

                    var cellValue = cell.ToString();
                    
                    var value = ValueConverter(propertyInfo,cellValue);
                    SetValueForProperty(propertyInfo,insertObject,value);
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

    private static void SetValueForProperty(PropertyInfo propertyInfo, TObj insertObject, object? value)
    {
        try
        {
            propertyInfo.SetValue(insertObject,value);
        }
        catch (ArgumentException e)
        {
            throw new PropertyInaccessibleException($"Error while setting value to property {propertyInfo.Name} in {typeof(TObj).Name}",e);
        }
    }

    private static object? ValueConverter(PropertyInfo propertyInfo,string? cellValue)
    {
        var propertyInfoPropertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
        return string.IsNullOrWhiteSpace(cellValue) ? null : Convert.ChangeType(cellValue, propertyInfoPropertyType);
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
}



