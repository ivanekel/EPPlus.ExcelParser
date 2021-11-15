using EPPlus.ExcelParser.Exporting.Mapping;
using EPPlus.ExcelParser.Shared;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace EPPlus.ExcelParser.Exporting.Exporter
{
    public class ExcelExporter<TObject> where TObject : class
    {
        private readonly ExcelPackage _excelPackage;
        private readonly bool _hasHeaders;

        private IEnumerable<TObject> _data;
        private IObjectMapper<TObject> _rowMapper;

        public ExcelExporter(IEnumerable<TObject> data, bool hasHeaders = false)
        {
            _data = data;
            _hasHeaders = hasHeaders;
            _excelPackage = new ExcelPackage();
        }

        public ExcelPackage GetExcel()
        {
            if (_rowMapper == null)
                throw new Exception("Mapper is not set");

            var worksheet = _excelPackage.Workbook.Worksheets.Add("Sheet");

            var rowNumber = 1;

            if (_hasHeaders)
            {
                _rowMapper.SetHeaders(worksheet);
                rowNumber++;
            }

            foreach (var item in _data)
            {
                _rowMapper.Map(item, worksheet, rowNumber);
                rowNumber++;
            }

            _rowMapper.AutoFit(worksheet);

            return _excelPackage;
        }

        public ExcelExporter<TObject> SetMapping(ExcelMapper<TObject> mapper)
        {
            _rowMapper = new DefaultObjectMappingProvider<TObject>(mapper);
            return this;
        }

        public ExcelExporter<TObject> SetMapping<TMapper>() where TMapper : ExcelMapper<TObject>
            => SetMapping(Activator.CreateInstance<TMapper>());
    }
}
