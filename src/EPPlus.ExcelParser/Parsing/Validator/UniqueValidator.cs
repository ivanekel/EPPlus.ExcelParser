using FluentValidation;
using FluentValidation.Validators;
using System;
using System.Collections.Generic;

namespace EPPlus.ExcelParser.Parsing.Validator
{
    public class UniqueValidator<T, TProperty> : PropertyValidator<T, TProperty>, INotEmptyValidator
    {
        public override string Name => "UniqueValidator";

        public override bool IsValid(ValidationContext<T> context, TProperty value)
        {
            if (!context.RootContextData.ContainsKey("Persistence"))
                throw new Exception("No persistence");

            var uniqueValues = (HashSet<(string, object)>)context.RootContextData["Persistence"];

            if (uniqueValues.Contains((context.PropertyName, value)))
            {
                return false;
            }

            uniqueValues.Add((context.PropertyName, value));

            return true;
        }

        protected override string GetDefaultMessageTemplate(string errorCode)
        {
            return Localized(errorCode, Name);
        }
    }
}