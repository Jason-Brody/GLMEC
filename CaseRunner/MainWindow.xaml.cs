using CaseRunnerModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TestScript.Case1;
using System.Timers;

namespace CaseRunner
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IRunner _currentCase;
        public MainWindow()
        {
            InitializeComponent();

        }

        private void btn_Load_Click(object sender, RoutedEventArgs e)
        {
            //cb_Cases.DataContext = Utils.GetCases();
        }

        private async void btn_Run_Click(object sender, RoutedEventArgs e)
        {
            _currentCase = new Case1_MTD_Analysis();


            lv_Items.DataContext = _currentCase.GetSteps;
            _currentCase.OnProcess += (s) =>
            {
                pb.Dispatcher.BeginInvoke(new Action(() =>
                {
                    pb.DataContext = s;
                }));
                lv_Items.Dispatcher.BeginInvoke(new Action(() =>
                {
                    lv_Items.SelectedItem = s;
                }));
                tb_Process.Dispatcher.BeginInvoke(new Action(() =>
                {
                    tb_Process.DataContext = s;
                }));
            };

            DateTime start = DateTime.Now;
            Timer t = new Timer();
            t.Elapsed += (o, e1) => {
                tb_Period.Dispatcher.BeginInvoke(new Action(()=> {
                    tb_Period.Text = DateTime.Now.Subtract(start).ToString();
                }));
            };
            t.Interval = TimeSpan.FromSeconds(1).Seconds;


            btn_Run.IsEnabled = false;
            t.Start();

            try
            {
                await Task.Run(() => _currentCase.Run());
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                btn_Run.IsEnabled = true;
                pb.IsIndeterminate = false;
                t.Stop();
            }
            

            

        }

        

        private void cb_Cases_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _currentCase = ((sender as ComboBox).SelectedItem as IRunner);
        }
    }
}
