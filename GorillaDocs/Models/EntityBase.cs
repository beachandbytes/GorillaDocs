using GorillaDocs.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace GorillaDocs.Models
{
    public class EntityBase : Notify
    {
        protected virtual void ValidateProperty(string propertyName, object value)
        {
            var context = new ValidationContext(this, null, null) { MemberName = propertyName };
            Validator.ValidateProperty(value, context);
        }
    }
}
