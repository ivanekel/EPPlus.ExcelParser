using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace EPPlus.ExcelParser.Shared
{
    public class ExcelMapper<TObject> where TObject : class
    {
        private readonly Dictionary<int, IExcelPropertyMapper<TObject>> _mappings = new Dictionary<int, IExcelPropertyMapper<TObject>>();
        private readonly Dictionary<int, string> _headers = new Dictionary<int, string>();

        public ExcelMapper<TObject> MapPropertyAndCell<TProperty>(Expression<Func<TObject, TProperty>> property, int columnNumber, string headerName = null)
        {
            _mappings.Add(columnNumber, new ExcelPropertyMapper<TObject, TProperty>(columnNumber, property));
            if (headerName != null) _headers.Add(columnNumber, headerName);
            return this;
        }

        public TObject MapFromExcel(ExcelWorksheet worksheet, int rowNumber)
        {
            var target = Activator.CreateInstance<TObject>();
            foreach (var map in _mappings)
            {
                try
                {
                    map.Value.MapFromExcel(target, worksheet.Cells[rowNumber, map.Value.ColumnNumber]);
                }
                catch
                {
                    throw new InvalidCastException($"cannot cast {{row,column}}:{{{rowNumber},{map.Value.ColumnNumber}}}. " +
                                                   $"Try make type nullable check type compatibility");
                }
            }

            return target;
        }
    }
}
