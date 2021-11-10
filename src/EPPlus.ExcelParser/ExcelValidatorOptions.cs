using System.Drawing;
using FluentValidation;

namespace EPPlus.ExcelParser
{
    public static class ExcelValidatorOptions
    {
        public static IRuleBuilderOptions<T, TProperty> WithRowColor<T, TProperty>(
            this IRuleBuilderOptions<T, TProperty> rule, KnownColor invalidColor = KnownColor.Red)
        {
            return rule.WithMessage("InvalidColorDefined").WithErrorCode(invalidColor.ToString());
        }
    }
}