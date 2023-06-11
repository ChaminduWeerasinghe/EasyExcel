namespace EasyOffice.EasyExcel.Exceptions;

public class FileNotFoundException : Exception
{
    public FileNotFoundException(string message) : base(message) { }
    public FileNotFoundException(string message, Exception inner) : base(message,inner) { }
}

public class EmptyFileException : Exception
{
    public EmptyFileException(string message) : base(message) { }
    public EmptyFileException(string message, Exception inner) : base(message,inner) { }
}

public class PropertyInaccessibleException : Exception
{
    public PropertyInaccessibleException(string message) : base(message) { }
    public PropertyInaccessibleException(string message, Exception inner) : base(message,inner) { }
}

public class InvalidFileTypeException : Exception
{
    public InvalidFileTypeException(string message) : base(message) { }
    public InvalidFileTypeException(string message, Exception inner) : base(message,inner) { }
}

public class InvalidValueException : Exception
{
    public InvalidValueException(string message) : base(message) { }
    public InvalidValueException(string message, Exception inner) : base(message,inner) { }
}