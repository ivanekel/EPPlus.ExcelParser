using EPPlus.ExcelParser.Parsing.Validator;
using System;

namespace EPPlus.ExcelParser.Parsing.Parser
{
    public partial class ExcelParser<TObject>
    {
        public ExcelParser<TObject> SetValidation(Action<ExcelValidator<TObject>> validatorBuilder)
        {
            _validation = new ExcelValidator<TObject>();
            validatorBuilder(_validation);
            return this;
        }

        public ExcelParser<TObject> SetValidation(ExcelValidator<TObject> validator)
        {
            _validation = validator;
            return this;
        }

        public ExcelParser<TObject> SetValidation<TValidator>() where TValidator : ExcelValidator<TObject>
        {
            _validation = Activator.CreateInstance<TValidator>();
            return this;
        }
    }
}