using CaseRunnerModel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunnerModel
{
    public class ScriptRunner<T,T1> where T :class,new()
    {
        private Dictionary<int, StepAttribute> _stepDic;


        public void Start()
        {
            addMethod();
            
        }
       

        private void addMethod()
        {
            _stepDic = new Dictionary<int, StepAttribute>();
            foreach (var method in typeof(T).GetMethods().Where(m => m.IsPublic))
            {
                var stepAttr = method.GetCustomAttribute<StepAttribute>(true);
                if (stepAttr != null)
                {
                    _stepDic.Add(stepAttr.Order, stepAttr);
                }
            }
        }
    }
}
