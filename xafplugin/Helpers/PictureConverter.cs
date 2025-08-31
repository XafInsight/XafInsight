namespace xafplugin.Helpers
{
    using System.Drawing;
    using System.Windows.Forms;

    public class PictureConverter : AxHost
    {
        private PictureConverter() : base("") { }

        public static stdole.IPictureDisp ImageToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }

}
