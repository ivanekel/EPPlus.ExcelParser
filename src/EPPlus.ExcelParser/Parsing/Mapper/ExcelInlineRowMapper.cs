using System;
using OfficeOpenXml;

namespace EPPlus.ExcelParser.Parsing.Mapper
{
    public class ExcelInlineRowMapper
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly int _row;


        internal ExcelInlineRowMapper(ExcelWorksheet worksheet, int row)
        {
            _worksheet = worksheet;
            _row = row;
        }

        public TCustomProperty GetValue<TCustomProperty>(int column)
        {
            try
            {
                return _worksheet.Cells[_row, column].GetValue<TCustomProperty>();
            }
            catch
            {
                throw new InvalidCastException($"cannot cast {{row,column}}:{{{_row},{column}}}. " +
                                               $"Try make type nullable check type compatibility");
            }
        }
    }
}