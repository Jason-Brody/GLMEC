using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseRunnerModel
{
    public class ProcessInfo1 : WPFNotify
    {
        public int Id { get; set; }

        private string _name;

        private string _des;

        private bool _isComplete;

        private static int _currentProcess;

        private static int _TotalProcess;

        private bool _isProcessKnown;

        public bool IsProcessKnown
        {
            get { return _isProcessKnown; }
            set { SetProperty<bool>(ref _isProcessKnown, value); }
        }

        public int CurrentProcess
        {
            get { return _currentProcess; }

            set { SetProperty<int>(ref _currentProcess, value); }
        }

        public int TotalProcess
        {
            get { return _TotalProcess; }

            set { SetProperty<int>(ref _TotalProcess, value); }
        }

        public string Name
        {
            get { return _name; }
            set { SetProperty<string>(ref _name, value); }
        }

        public string Description
        {
            get { return _des; }
            set { SetProperty<string>(ref _des, value); }
        }


    }
}
