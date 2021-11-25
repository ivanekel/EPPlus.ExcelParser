using FluentValidation;
using FluentValidation.Internal;
using System.Collections.Generic;

namespace EPPlus.ExcelParser.Parsing.Validator
{
    public class PersistentValidationContext<TObject> : ValidationContext<TObject>
    {
        public PersistentValidationContext(TObject instanceToValidate) : base(instanceToValidate)
        {
        }

        public PersistentValidationContext(TObject instanceToValidate, PropertyChain propertyChain, IValidatorSelector validatorSelector) : base(instanceToValidate, propertyChain, validatorSelector)
        {
        }
    }
}