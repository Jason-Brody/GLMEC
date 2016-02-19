using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunnerModel
{

    public delegate void ProcessHander(StepInfo step);
    public interface IRunner
    {
        event ProcessHander OnProcess;

        void Run();

        List<StepInfo> GetSteps { get; }

        CaseInfo Info { get; }

    }
}
