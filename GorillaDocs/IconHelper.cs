using System.Drawing;
using System.IO;
using System.Windows.Media.Imaging;

namespace GorillaDocs
{
    public class IconHelper
    {
        readonly string fullname;
        public IconHelper(string fullname)
        {
            this.fullname = fullname;
        }

        public void SaveAsPng(string outputPath)
        {
            var icon = Icon.ExtractAssociatedIcon(fullname);
            using (Bitmap bmp = icon.ToBitmap())
                bmp.Save(outputPath, System.Drawing.Imaging.ImageFormat.Png);
        }

        public BitmapFrame AsBitmapFrame()
        {
            var icon = Icon.ExtractAssociatedIcon(fullname);
            using (Bitmap bmp = icon.ToBitmap())
            {
                var stream = new MemoryStream();
                bmp.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                return BitmapFrame.Create(stream);
            }
        }
    }
}
