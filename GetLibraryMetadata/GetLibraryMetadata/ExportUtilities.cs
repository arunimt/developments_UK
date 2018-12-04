using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace GetLibraryMetadata
{
    class ExportUtilities
    {
        private string Filename;

        public ExportUtilities(bool append = false)
        {
            Filename = ConfigurationSettings.AppSettings["exportfile"];

            // Log file header line
            string logHeader = Filename + " is created.";
        }

        public void WriteLine(string text, bool append = true)
        {
            try
            {
                using (StreamWriter Writer = new StreamWriter(Filename, append, Encoding.UTF8))
                {
                    if (text != "") Writer.WriteLine(text);
                }
            }
            catch
            {
                throw;
            }
        }

    }
}
