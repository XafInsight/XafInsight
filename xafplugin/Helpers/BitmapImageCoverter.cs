using System.IO;
using System.Windows.Media.Imaging;

namespace xafplugin.Helpers
{
    public class BitmapImageCoverter
    {
        /// <summary>
        /// Converts a byte array to an ImageSource that can be used as a Window Icon
        /// </summary>
        /// <param name="byteArray">The byte array containing the image data</param>
        /// <returns>An ImageSource object that can be assigned to Window.Icon</returns>
        public static System.Windows.Media.ImageSource ByteArrayToIcon(byte[] byteArray)
        {
            if (byteArray == null || byteArray.Length == 0)
                return null;

            var image = new BitmapImage();
            using (var stream = new MemoryStream(byteArray))
            {
                stream.Position = 0;
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.StreamSource = stream;
                image.EndInit();
                image.Freeze(); // Makes the image usable across threads
            }
            return image;
        }
    }
}
