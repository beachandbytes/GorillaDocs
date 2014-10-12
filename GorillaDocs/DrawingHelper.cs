using System.Drawing;
using System.IO;
using System.Windows.Media.Imaging;

namespace GorillaDocs
{
    public static class DrawingHelper
    {
        public static BitmapFrame AsBitmapFrame(this Bitmap bitmap)
        {
            var stream = new MemoryStream();
            bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
            return BitmapFrame.Create(stream);
        }
    }
}
