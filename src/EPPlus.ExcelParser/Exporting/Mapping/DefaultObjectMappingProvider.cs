using EPPlus.ExcelParser.Shared;
using OfficeOpenXml;

namespace EPPlus.ExcelParser.Exporting.Mapping
{
    internal class DefaultObjectMappingProvider<TObject> : IObjectMapper<TObject> where TObject : class
    {
        private ExcelMapper<TObject> _mapper;

        public DefaultObjectMappingProvider(ExcelMapper<TObject> mapper)
        {
            _mapper = mapper;
        }

        public void AutoFit(ExcelWorksheet worksheet)
        {
            _mapper.AutoFit(worksheet);
        }

        public void Map(TObject target, ExcelWorksheet worksheet, int rowNumber)
        {
            _mapper.MapToExcel(target, worksheet, rowNumber);
        }

        public void SetHeaders(ExcelWorksheet worksheet)
        {
            _mapper.SetHeaders(worksheet);
        }
    }
}
