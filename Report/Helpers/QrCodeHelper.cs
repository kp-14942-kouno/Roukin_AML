using QRCoder;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace MyTemplate.Report.Helpers
{
    public class QrCoderHelper
    {
        /// <summary>
        /// テキストからQRコード画像を生成し、BitmapImageで返す
        /// </summary>
        public static BitmapImage GenerateQrCode(string text, int pixelsPerModule = 20)
        {
            using var qrGenerator = new QRCodeGenerator();
            using var qrData = qrGenerator.CreateQrCode(text, QRCodeGenerator.ECCLevel.Q);
            using var qrCode = new QRCode(qrData);
            using Bitmap qrBitmap = qrCode.GetGraphic(pixelsPerModule);

            using Bitmap mono = qrBitmap.Clone(
                new Rectangle(0, 0, qrBitmap.Width, qrBitmap.Height),
                System.Drawing.Imaging.PixelFormat.Format1bppIndexed);

            using var ms = new MemoryStream();
            qrBitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Tiff);
            ms.Position = 0;

            var image = new BitmapImage();
            image.BeginInit();
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.StreamSource = ms;
            image.EndInit();
            image.Freeze();
            return image;
        }
    }
}
