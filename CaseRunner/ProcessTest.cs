using CaseRunnerModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunner
{
    class ProcessTest: INotifyPropertyChanged
    {

        private static int _current;

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
