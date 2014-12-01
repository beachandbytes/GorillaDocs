using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;
using O = Microsoft.Office.Core;
using System.IO;
using System.Linq.Expressions;

namespace GorillaDocs.Word
{
    public static class FirmAddressHelper
    {
        const string TagAndBookmark = "FirmAddress";

        public static void UpdateFirmAddressesControls<T>(this Wd.Document doc, T Office, string Namespace, FileInfo FirmAddress)
        {
            try
            {
                doc.UpdateFirmAddressPart(Office, Namespace);
                doc.FirmAddressControls(x => x.ReinsertFromFile(FirmAddress, TagAndBookmark));
                doc.FirmAddressControls(x => x.Range.ContentControls.DeleteUnMapped());
            }
            catch (Exception ex)
            {
                Message.LogWarning("If the following error is related to Protected Memory, then the error occured when iterating the Shapes collection. No idea why it only happens sometimes..");
                Message.LogError(ex);
            }
        }

        static void FirmAddressControls(this Wd.Document doc, Action<Wd.ContentControl> action)
        {
            foreach (Wd.ContentControl control in doc.ContentControls)
                if (control.Tag == TagAndBookmark)
                    action(control);
        }

        static void ReinsertFromFile(this Wd.ContentControl control, FileInfo FirmAddress, string Bookmark)
        {
            Wd.Range range = control.Range;
            range.Delete();
            range.InsertFromFile(FirmAddress.FullName, Bookmark);
        }

        public static void UpdateFirmAddressPart<T>(this Wd.Document doc, T Office, string Namespace)
        {
            try
            {
                var parts = doc.CustomXMLParts.SelectByNamespace(Namespace);
                foreach (O.CustomXMLPart part in parts)
                    part.Delete();
                doc.CustomXMLParts.Add(Serializer.SerializeToString<T>(Office));
            }
            catch (Exception ex)
            {
                Message.LogWarning("If the following error is related to Protected Memory, then the error occured when iterating the Shapes collection. No idea why it only happens sometimes..");
                Message.LogError(ex);
            }
        }
   }
}
