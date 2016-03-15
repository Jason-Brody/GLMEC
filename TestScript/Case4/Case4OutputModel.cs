using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestScript.Case4
{
    class Case4OutputModel:FileConfig
    {
        public Case4OutputModel() : base("Case4") { }

        public List<CSKSDataModel> CSKS { get; set; }

        public string CSKSFile { get { return getFile("CSKS.txt"); } }

        public List<COASDataModel> COAS { get; set; }

        public string COASFile { get { return getFile("COAS.txt"); } }

        public List<string> GLAccounts { get; set; }

    }
}
