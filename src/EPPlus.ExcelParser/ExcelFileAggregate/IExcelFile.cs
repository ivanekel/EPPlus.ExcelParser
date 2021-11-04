using System.IO;

namespace EPPlus.ExcelParser.ExcelFileAggregate
{
    public interface IExcelFile
    {
        Stream ExcelFileStream { get; }
    }
}