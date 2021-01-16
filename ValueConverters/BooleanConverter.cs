using AutoRenewal.Models;
using System;
using System.Globalization;
using System.Windows.Data;

namespace AutoRenewal.ValueConverters
{
    public class BooleanConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            var org = (Organization)values[0];
            var inputPath = values[1] as string;
            return (org != null && !string.IsNullOrWhiteSpace(inputPath));
        }

        public object[] ConvertBack(object value, Type[] targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
