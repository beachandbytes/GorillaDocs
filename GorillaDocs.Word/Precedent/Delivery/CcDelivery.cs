using GDP = GorillaDocs.Properties;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Precedent.Delivery
{
    public class CcDelivery1 : Delivery1
    {
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
            Wd.Range range = control.Range;
            range.MoveOutOfContentControl();
            range.Text = ": ";
            range.Collapse(Wd.WdCollapseDirection.wdCollapseEnd);
            AddDetailsControl(range);
        }

        void RemoveDeliveryDetails()
        {
            if (DetailsControl != null)
                DetailsControl.Delete(true);
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
