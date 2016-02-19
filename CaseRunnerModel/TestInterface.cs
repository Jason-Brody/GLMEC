using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunnerModel
{
    //public class TestInterface : IRunner
    //{
    //    List<StepInfo> _testSteps = new List<StepInfo>();

    //    public TestInterface()
    //    {
    //        _testSteps.Add(new StepInfo() { IsProcessKnown = true, Description = "Step001" });
    //        _testSteps.Add(new StepInfo() { IsProcessKnown = true, Description = "Step002" });
    //        _testSteps.Add(new StepInfo() { IsProcessKnown = true, Description = "Step003" });
    //        _testSteps.Add(new StepInfo() { IsProcessKnown = true, Description = "Step004" });
    //        _testSteps.Add(new StepInfo() { IsProcessKnown = true, Description = "Step005" });
    //        _testSteps.Add(new StepInfo() { IsProcessKnown = true, Description = "Step006" });
    //        _testSteps.Add(new StepInfo() { IsProcessKnown = true, Description = "Step007" });
    //    }

    //    public List<StepInfo> GetSteps
    //    {
    //        get
    //        {
    //            return _testSteps;
    //        }
    //    }

    //    public CaseInfo Info
    //    {
    //        get
    //        {
    //            return new CaseInfo() { Name = "Test001" };
    //        }
    //    }

    //    public void Run()
    //    {
    //        foreach(var s in _testSteps)
    //        {
    //            Random rd = new Random();
    //            s.TotalProcess = rd.Next(100);
    //            for(int i = 0;i < s.TotalProcess; i++)
    //            {
    //                Task.Delay(50).Wait();
    //                s.CurrentProcess++;
    //            }

    //        }

            
    //    }
    //}
}
