using GorillaDocs.Word.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using GDP = GorillaDocs.Properties;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word
{
    public static class DeliveryHelpers
    {
        public static bool IsDelivery(this Wd.ContentControl control)
        {
            return (!string.IsNullOrEmpty(control.Tag) && control.Tag.EndsWith(GDP.Resources.ContentControl_DeliveryEnding)) ||
                (!string.IsNullOrEmpty(control.Title) && control.Title.EndsWith(GDP.Resources.ContentControl_DeliveryEnding));
        }
        public static bool IsRecipient(this Wd.ContentControl control)
        {
            return (!string.IsNullOrEmpty(control.Tag) && control.Tag.StartsWith(GDP.Resources.ContentControl_Recipient)) ||
                (!string.IsNullOrEmpty(control.Title) && control.Title.StartsWith(GDP.Resources.ContentControl_Recipient));
        }
        public static bool IsCC(this Wd.ContentControl control)
        {
            return (!string.IsNullOrEmpty(control.Tag) && control.Tag.StartsWith(GDP.Resources.ContentControl_CC)) ||
                (!string.IsNullOrEmpty(control.Title) && control.Title.StartsWith(GDP.Resources.ContentControl_CC));
        }

        public static Delivery AsDelivery(this Wd.ContentControl control)
        {
            if (control.IsRecipient())
                return new RecipientDelivery(control);
            else if (control.IsCC())
                return new CcDelivery(control);
            else
                throw new InvalidOperationException("The Content Control is not a Recipient or a CC control.");
        }
    }
}
