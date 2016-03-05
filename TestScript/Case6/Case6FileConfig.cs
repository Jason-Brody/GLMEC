using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript.Case6
{
    public class Case6FileConfig
    {
        private string _workDir;
        public Case6FileConfig(string BaseDirectory)
        {
            this._workDir = BaseDirectory;
        }

        private const string _source = "SourceOutput";

        public string WorkFolder { get { return _workDir; } }

        public string SourceOutputFolder { get { return Path.Combine(_workDir, _source); } }

        public string EDIDCFileName { get; } = "EDIDC.txt";


        public string EDIDSFileName { get; } = "EDIDS.txt";

        public string EDP21Folder { get; } = "EDP21";

        public string EDP21FileName { get; } = "EDP21.txt";

        public string HRP1001FileName { get; } = "HRP1001.txt";

        public string GetFullPath(string fileName)
        {
            return Path.Combine(_workDir, fileName);
        }

        public string ReportCSVFile { get; } = "Report.csv";

        public string ReportExcelFile { get; } = "Report.xlsx";

        public string GetSourcePath(string fileName)
        {
            return Path.Combine(_workDir, _source, fileName);
        }
    }
}
