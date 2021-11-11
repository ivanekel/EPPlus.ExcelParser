using EPPlus.ExcelParser.Parsing.Validator;
using FluentValidation;
using System.Drawing;

namespace EPPlus.ExcelParser.Parsing
{
    public static class ValidationExtensions
    {
        internal const string ColorCodePrefix = "Color-Code-";

        public static IRuleBuilderOptions<T, TProperty> Unique<T, TProperty>(
            this IRuleBuilder<T, TProperty> ruleBuilder)
        {
            return ruleBuilder.SetValidator(new UniqueValidator<T, TProperty>());
        }

        public static IRuleBuilderOptions<T, TProperty> WithRowErrorColor<T, TProperty>(
            this IRuleBuilderOptions<T, TProperty> rule, KnownColor invalidColor = KnownColor.Red)
        {
            return rule.WithErrorCode($"{ColorCodePrefix}{invalidColor}");
        }
    }
}