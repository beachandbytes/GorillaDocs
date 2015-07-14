using System;
using GDP = GorillaDocs.Properties;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Delivery
{
    public static class DeliveryHelpers
    {
        public static bool IsDelivery1(this Wd.ContentControl control)
        {
            return !string.IsNullOrEmpty(control.Title) && control.Title.EndsWith(GDP.Resources.ContentControl_DeliveryEnding);
        }
        public static bool IsRecipient1(this Wd.ContentControl control)
        {
            return !string.IsNullOrEmpty(control.Title) && control.Title.StartsWith(GDP.Resources.ContentControl_Recipient);
        }
        public static bool IsCC1(this Wd.ContentControl control)
        {
            return !string.IsNullOrEmpty(control.Title) && control.Title.StartsWith(GDP.Resources.ContentControl_CC);
        }

        public static Delivery1 AsDelivery1(this Wd.ContentControl control)
        {
            if (control.IsRecipient1())
                return new RecipientDelivery1(control);
            else if (control.IsCC1())
                return new CcDelivery1(control);
            else
                throw new InvalidOperationException("The Content Control is not a Recipient or a CC control.");
        }
    }
}
