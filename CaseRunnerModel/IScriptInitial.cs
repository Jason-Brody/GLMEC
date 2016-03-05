using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunnerModel
{
    public interface IScriptRunner
    {
        void SetInputData(object data, IProgress<ProcessInfo> MyProgress);
    }
}
