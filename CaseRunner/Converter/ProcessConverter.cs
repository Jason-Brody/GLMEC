using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace CaseRunner.Converter
{
    class ProcessConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if(values[0]!= DependencyProperty.UnsetValue)
            {
                bool isProcessKnow = bool.Parse(values[0].ToString());
                if (!isProcessKnow)
                    return "";
                return $"{values[1]}/{values[2]}";
            }
            return "";
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
