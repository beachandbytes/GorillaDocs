using System;
using System.Collections.Generic;
using System.Linq;
using GDP = GorillaDocs.Properties;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Models
{
    public class RecipientDelivery : Delivery
    {
        private const string spacer = ": ";
        public RecipientDelivery(Wd.ContentControl control) : base(control) { }

        protected override Wd.ContentControl DetailsControl
        {
            get
            {
                string tag = control.Tag.Substring(0, control.Tag.IndexOf(GDP.Resources.ContentControl_DeliveryEnding)) + "Details";
                var range = control.Range.Paragraphs[1].Range;
                return range.ContentControls(x => x.Tag == tag).FirstOrDefault();
            }
        }

        public override void Update()
        {
            if (IsEmail)
                UpdateEmail();
            else if (IsFacsimile)
                UpdateFacsimile();
            else
                RemoveDeliveryDetails();
        }

        protected override void AddDetailsControl()
        {
            Wd.Range range = control.Range.CollapseEnd();
            range.Move(Wd.WdUnits.wdCharacter, 1);
            range.Text = spacer;
            range = range.CollapseEnd();
            AddDetailsControl(range);
        }

        void RemoveDeliveryDetails()
        {
            if (DetailsControl != null)
            {
                var range = DetailsControl.Range;
                DetailsControl.Delete(true);
                range.MoveStart(Wd.WdUnits.wdCharacter, -2);
                if (range.Text == spacer)
                    range.Delete();
            }
        }
    }
}
