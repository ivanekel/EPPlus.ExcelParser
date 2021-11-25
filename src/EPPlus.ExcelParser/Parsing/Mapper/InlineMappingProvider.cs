using OfficeOpenXml;
using System;

namespace EPPlus.ExcelParser.Parsing.Mapper
{
    internal class InlineMappingProvider<TObject> : IRowMapper<TObject>
    {
        private Func<ExcelInlineRowMapper, TObject> _mapper;

        internal InlineMappingProvider(Func<ExcelInlineRowMapper, TObject> mapper)
        {
            _mapper = mapper;
        }

        public TObject Map(ExcelWorksheet worksheet, int row)
            => _mapper(new ExcelInlineRowMapper(worksheet, row));
    }
}
