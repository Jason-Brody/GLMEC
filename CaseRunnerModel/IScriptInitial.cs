using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunnerModel
{
    public interface IScriptRunner1
    {
        void SetInputData(object data, IProgress<ProcessInfo1> MyProgress);
    }
}
