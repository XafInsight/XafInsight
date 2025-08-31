using xafplugin.Interfaces;

namespace xafplugin.Modules
{
    public class EnvironmentService : IEnvironmentService
    {

        public string DatabasePath
        {
            get
            {
                return Globals.ThisAddIn?.TempDbPath;
            }
            set
            {
                if (Globals.ThisAddIn != null)
                {
                    Globals.ThisAddIn.TempDbPath = value;
                }
            }
        }

        public string FileHash
        {
            get
            {
                return Globals.ThisAddIn?.FileHash;
            }
            set
            {
                if (Globals.ThisAddIn != null)
                {
                    Globals.ThisAddIn.FileHash = value;
                }
            }
        }
    }

}
