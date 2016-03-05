using CaseRunnerModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;

namespace CaseRunner
{
    class Utils
    {
        public static List<IRunner> GetCases()
        {
            List<Assembly> assembs = new List<Assembly>();
            AppDomain.CurrentDomain.AssemblyResolve += (s, e) =>
            {
                return assembs.Where(a => a.FullName == e.Name).First();
            };
            List<IRunner> cases = new List<IRunner>();
           
            DirectoryInfo di = new DirectoryInfo("TestCases");
            foreach (var file in di.GetFiles())
            {
                
                Assembly asm = Assembly.LoadFile(file.FullName);
                
                cases.AddRange(getCaseInfo(asm));
                assembs.Add(asm);
                
            }
            return cases;
        }

       

        private static List<IRunner> getCaseInfo(Assembly asm)
        {
            List<IRunner> runners = new List<IRunner>();
            foreach (var t in asm.GetTypes())
            {
                if (t.GetInterface(nameof(IRunner)) != null)
                {
                   
                    object instance = Activator.CreateInstance(t);
                   
                    if (instance is IRunner)
                        runners.Add((IRunner)instance);
                }
            }
            return runners;
        }
    }
}
