using System;
using WindowsInstaller;

namespace msi_get
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Indicate path to MSI as an argument");
                return;
            }

            var filename = args[0];

            Console.WriteLine("{0}", GetMsiProperty(filename, "ProductCode"));
        }

        // This method is taken from http://www.alteridem.net/2008/05/20/read-properties-from-an-msi-file/
        public static string GetMsiProperty(string msiFile, string property)
        {
            string retVal = string.Empty;

            // Create an Installer instance
            Type classType = Type.GetTypeFromProgID( "WindowsInstaller.Installer" );
            Object installerObj = Activator.CreateInstance(classType);
            Installer installer = installerObj as Installer;

            // Open the msi file for reading
            // 0 - Read, 1 - Read/Write
            Database database = installer.OpenDatabase(msiFile, 0);

            // Fetch the requested property
            string sql = String.Format( "SELECT Value FROM Property WHERE Property ='{0}'", property );
            View view = database.OpenView(sql);
            view.Execute(null);

            // Read in the fetched record
            Record record = view.Fetch();
            if (record != null)
                retVal = record.get_StringData(1);

            return retVal;
        }
    }
}
