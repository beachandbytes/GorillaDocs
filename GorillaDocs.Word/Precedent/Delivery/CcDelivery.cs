using GDP = GorillaDocs.Properties;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Delivery
{
    public class CcDelivery1 : Delivery1
    {
        const string spacer = ": ";
        public CcDelivery1(Wd.ContentControl control) : base(control) { }

        public override void Update()
        {
            if (IsEmail)
                UpdateEmail();
            else if (IsFacsimile)
                UpdateFacsimile();
            else if (IsOther)
                UpdateAddress();
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

        bool IsOther { get { return this.control.Range.Text.Contains(GDP.Resources.Delivery_Other); } }

        void UpdateAddress()
        {
            if (DetailsControl == null)
                AddDetailsControl();
            SetMapping("Address");

            var range = control.Range;
            control.Delete(true);
            range.Delete();
            range.Delete();
            range.Delete();
        }
    }
}
