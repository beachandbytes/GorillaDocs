using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GorillaDocs
{
    public class ImageHelper
    {
        public static stdole.IPictureDisp GetImage(string imageName)
        {
            switch (imageName)
            {
                case "BT":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.BT);
                case "BTI":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.BTI);
                case "DiscardChanges":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.DiscardChanges);
                case "H1":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.H1);
                case "H2":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.H2);
                case "H3":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.H3);
                case "H4":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.H4);
                case "H5":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.H5);
                case "H6":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.H6);
                case "SH":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.SH);
                case "DP":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.DP);
                case "D":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.D);
                case "D1":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.D1);
                case "D2":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.D2);
                case "D3":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.D3);
                case "S1":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.S1);
                case "S2":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.S2);
                case "S3":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.S3);
                case "S4":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.S4);
                case "S5":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.S5);
                case "S6":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.S6);
                case "L1":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.L1);
                case "L2":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.L2);
                case "L3":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.L3);
                case "L4":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.L4);
                case "B1":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.B1);
                case "B2":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.B2);
                case "B3":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.B3);
                case "EditDocument":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.EditDocument);
                case "ToggleLogo":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.ToggleLogo);
                case "IN":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.IN);
                case "Landscape":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.Landscape);
                case "Portrait":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.Portrait);
                default:
                    throw new ArgumentException(string.Format("'{0}' does not exist.", imageName));
            }
        }

    }
}
