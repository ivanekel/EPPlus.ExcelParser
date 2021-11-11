using OfficeOpenXml;

namespace EPPlus.ExcelParser.Shared
{
    internal interface IExcelPropertyMapper<TObject>
    {
        int ColumnNumber { get; }
        void MapFromExcel(TObject target, ExcelRange cell);
        void MapToExcel(TObject target, ExcelWorksheet worksheet, int rowNumber);
    }
}
