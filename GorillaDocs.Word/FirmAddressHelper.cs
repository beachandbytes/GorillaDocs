using System;
using System.Collections.Generic;
using System.Linq;
using Wd = Microsoft.Office.Interop.Word;
using O = Microsoft.Office.Core;
using System.IO;

namespace GorillaDocs.Word
{
    public static class FirmAddressHelper
    {
        const string TagAndBookmark = "FirmAddress";

        public static void UpdateFirmAddressesControls(this Wd.HeadersFooters headersfooters, FileInfo FirmAddress)
        {
            foreach (Wd.HeaderFooter hf in headersfooters)
                hf.Range.UpdateFirmAddressesControls(FirmAddress);
        }

        public static void UpdateFirmAddressesControls(this Wd.Shapes shapes, FileInfo FirmAddress)
        {
            foreach (Wd.Shape shape in shapes)
                if (shape.Type == O.MsoShapeType.msoTextBox)
                    shape.TextFrame.TextRange.UpdateFirmAddressesControls(FirmAddress);
        }

        public static void UpdateFirmAddressesControls(this Wd.Range range, FileInfo FirmAddress)
        {
            try
            {
                range.FirmAddressControls(x => x.ReinsertFromFile(FirmAddress, TagAndBookmark));
                range.FirmAddressControls(x => x.Range.ContentControls.DeleteEmptyMappedControls());
            }
            catch (Exception ex)
            {
                Message.LogWarning("If the following error is related to Protected Memory, then the error occured when iterating the Shapes collection. No idea why it only happens sometimes..");
                Message.LogError(ex);
            }
        }

        static void FirmAddressControls(this Wd.Range range, Action<Wd.ContentControl> action)
        {

            foreach (Wd.ContentControl control in range.ContentControls.FindAll(TagAndBookmark))
                action(control);
        }

        static void ReinsertFromFile(this Wd.ContentControl control, FileInfo FirmAddress, string Bookmark)
        {
            Wd.Range range = control.Range;
            control.DeleteParagraphIfEmpty();
            range.InsertFile_Safe(FirmAddress.FullName, Bookmark);
            if (range.Bookmarks.Exists(Bookmark))
                range.Bookmarks[Bookmark].Delete();
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
