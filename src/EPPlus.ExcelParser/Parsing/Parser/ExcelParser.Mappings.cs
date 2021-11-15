using EPPlus.ExcelParser.Parsing.Mapper;
using EPPlus.ExcelParser.Shared;
using System;

namespace EPPlus.ExcelParser.Parsing.Parser
{
    public partial class ExcelParser<TObject> where TObject : class
    {
        public ExcelParser<TObject> SetMapping(Func<ExcelInlineRowMapper, TObject> mapper)
        {
            _rowMapper = new InlineMappingProvider<TObject>(mapper);
            return this;
        }

        public ExcelParser<TObject> SetMapping(ExcelMapper<TObject> mapper)
        {
            _rowMapper = new DefaultRowMappingProvider<TObject>(mapper);
            return this;
        }

        public ExcelParser<TObject> SetMapping<TMapper>() where TMapper : ExcelMapper<TObject>
            => SetMapping(Activator.CreateInstance<TMapper>());
    }
}