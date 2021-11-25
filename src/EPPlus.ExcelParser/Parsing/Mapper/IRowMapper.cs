using OfficeOpenXml;

namespace EPPlus.ExcelParser.Parsing.Mapper
{
    internal interface IRowMapper<TObject>
    {
        TObject Map(ExcelWorksheet worksheet, int row);
    }
}
