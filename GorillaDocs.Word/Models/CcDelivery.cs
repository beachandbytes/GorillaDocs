using System;
using System.Collections.Generic;
using System.Linq;
using GDP = GorillaDocs.Properties;
using Wd = Microsoft.Office.Interop.Word;

namespace GorillaDocs.Word.Models
{
    public class CcDelivery : Delivery
    {
        public CcDelivery(Wd.ContentControl control) : base(control) { }

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
            range.Text = " ";
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
        }
    }
}
