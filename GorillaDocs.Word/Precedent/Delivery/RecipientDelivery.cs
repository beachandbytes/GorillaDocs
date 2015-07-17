using System.Linq;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Delivery
{
    public class RecipientDelivery1 : Delivery1
    {
        const string spacer = ": ";
        public RecipientDelivery1(Wd.ContentControl control) : base(control) { }

        protected override Wd.ContentControl DetailsControl
        {
            get
            {
                var range = control.Range;
                range.MoveEndUntil("\r\v\a");
                return range.ContentControls(x => x.Title == control.Title + " Details").FirstOrDefault();
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
