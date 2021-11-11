using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using FluentValidation;
using FluentValidation.Internal;
using FluentValidation.Results;

namespace EPPlus.ExcelParser.Parsing.Validator
{
    public class ExcelValidator<TCustomObject> : AbstractValidator<TCustomObject>
    {
        private readonly HashSet<(string, object)> _persistence = new HashSet<(string, object)>();

        public ValidationResult ValidateWithPersistence(TCustomObject instance)
        {
            return Validate(new ValidationContext<TCustomObject>(instance, new PropertyChain(),
                ValidatorOptions.Global.ValidatorSelectors.DefaultValidatorSelectorFactory())
            {
                RootContextData = { new KeyValuePair<string, object>("Persistence", _persistence) }
            });
        }

        public Task<ValidationResult> ValidateWithPersistenceAsync(TCustomObject instance,
            CancellationToken cancellation = new CancellationToken())
        {
            return ValidateAsync(
                new ValidationContext<TCustomObject>(instance, new PropertyChain(),
                    ValidatorOptions.Global.ValidatorSelectors.DefaultValidatorSelectorFactory())
                {
                    RootContextData = { new KeyValuePair<string, object>("Persistence", this) }
                }, cancellation);
        }
    }
}