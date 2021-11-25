using OfficeOpenXml;

namespace EPPlus.ExcelParser.Exporting.Mapping
{
    internal interface IObjectMapper<TObject>
    {
        void Map(TObject target, ExcelWorksheet worksheet, int rowNumber);
        void SetHeaders(ExcelWorksheet worksheet);
        void AutoFit(ExcelWorksheet worksheet);
    }
}
