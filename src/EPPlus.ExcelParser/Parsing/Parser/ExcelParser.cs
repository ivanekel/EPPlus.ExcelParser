using EPPlus.ExcelParser.Parsing.Mapper;
using EPPlus.ExcelParser.Parsing.Validator;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.ExcelParser.Parsing.Parser
{
    public partial class ExcelParser<TObject>
    {
        public bool IsValid { get => !_hasErrors; }

        private readonly ExcelPackage _excelPackage;
        private readonly Marking _marking = new Marking();
        private readonly int _firstRowNumber;

        private ExcelValidator<TObject> _validation;
        private IRowMapper<TObject> _rowMapper;
        private bool _hasErrors = false;

        internal ExcelParser(ExcelPackage excelPackage, bool hasHeaders)
        {
            _excelPackage = excelPackage;
            _firstRowNumber = hasHeaders ? 2 : 1;
        }

        public IEnumerable<TObject> GetRows(bool raiseCastExceptions = false)
        {
            if (_rowMapper == null)
                throw new Exception("Mapper is not set");

            var worksheet = _excelPackage.Workbook.Worksheets.First();

            for (var row = _firstRowNumber; row <= worksheet.Dimension.Rows; row++)
            {
                TObject mappedObject;

                try
                {
                    mappedObject = _rowMapper.Map(worksheet, row);
                }
                catch
                {
                    if (raiseCastExceptions) throw;

                    _marking.MarkRowAsInvalid(worksheet, row);
                    continue;
                }

                var validationResult = _validation.ValidateWithPersistence(mappedObject);
                _hasErrors = !validationResult.IsValid;

                if (validationResult.IsValid == false)
                {
                    _marking.MarkRowAsInvalidFromValidationResult(worksheet, row, validationResult);
                    continue;
                }

                _marking.MarkRowAsValid(worksheet, row);
                yield return mappedObject;
            }
        }

        public ExcelParserResult<TObject> GetResult(bool raiseCastExceptions = false)
        {          
            return ExcelParserResult<TObject>.CreateNew(
                new List<TObject>(GetRows(raiseCastExceptions)),
                _excelPackage,
                !_hasErrors);
        }
    }
}