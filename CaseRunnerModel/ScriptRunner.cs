﻿using CaseRunnerModel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunnerModel
{
    public class ScriptRunner<T> where T : class, IScriptInitial, new()
    {
        private Dictionary<int, Tuple<StepAttribute, MethodInfo>> _stepDic;

        private T _obj = null;

        public ScriptRunner() { }

        public ScriptRunner(T obj)
        {
            this._obj = obj;
        }

        public void Run(object data)
        {
            if (_stepDic == null)
                addMethod();

            if (_obj == null)
                _obj = new T();
            _obj.InitialData(data);


            foreach (var item in _stepDic.OrderBy(o => o.Key))
            {
                item.Value.Item2.Invoke(_obj, null);
            }
        }

        private void addMethod()
        {
            _stepDic = new Dictionary<int, Tuple<StepAttribute, MethodInfo>>();
            foreach (var method in typeof(T).GetMethods().Where(m => m.IsPublic))
            {
                var stepAttr = method.GetCustomAttribute<StepAttribute>(true);
                if (stepAttr != null)
                {
                    _stepDic.Add(stepAttr.Order, new Tuple<StepAttribute, MethodInfo>(stepAttr, method));
                }
            }
        }
    }
}