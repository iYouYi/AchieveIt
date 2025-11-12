using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace AchieveItSystemUI
{
    public class UniversalBoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool boolValue = value is bool b && b;

            if (parameter?.ToString().ToLower() == "inverse")
                boolValue = !boolValue;

            return boolValue ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}