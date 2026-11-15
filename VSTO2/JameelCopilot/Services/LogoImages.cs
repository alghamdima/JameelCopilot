using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace JameelCopilot.Services
{
    /// <summary>
    /// The Jameel Copilot mark for the ribbon and the About dialog: one 256 px picture
    /// carried in the add-in (made from WebUI\jc-mark.svg by
    /// tools\Render-Logo.ps1) and scaled to the size asked for.
    /// </summary>
    internal static class LogoImages
    {
        private const int MasterSize = 256;

        /// <summary>A new 256 x 256 copy with transparent corners. The caller disposes it.</summary>
        public static Bitmap Master()
        {
            using (var stream = typeof(LogoImages).Assembly.GetManifestResourceStream("JameelCopilot.Resources.jc-256.png"))
            using (var loaded = new Bitmap(stream))
            {
                // A copy that does not need the stream kept open.
                return new Bitmap(loaded);
            }
        }

        /// <summary>The mark at <paramref name="pixels"/> square (16 to 256). The caller disposes it.</summary>
        public static Bitmap Sized(int pixels)
        {
            var size = Math.Max(16, Math.Min(MasterSize, pixels));
            using (var master = Master())
            {
                var scaled = new Bitmap(size, size, PixelFormat.Format32bppArgb);
                using (var graphics = Graphics.FromImage(scaled))
                using (var edges = new ImageAttributes())
                {
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    // Without this the picture's border is blended with nothing and comes out dark.
                    edges.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(master, new Rectangle(0, 0, size, size),
                        0, 0, master.Width, master.Height, GraphicsUnit.Pixel, edges);
                }
                return scaled;
            }
        }
    }
}
