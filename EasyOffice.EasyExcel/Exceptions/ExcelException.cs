namespace EasyOffice.EasyExcel.Exceptions;

public class ExcelFileNotFoundException : Exception
{
    public ExcelFileNotFoundException(string message) : base(message) { }
    public ExcelFileNotFoundException(string message, Exception inner) : base(message,inner) { }
}