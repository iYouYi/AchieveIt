using System;
using System.Globalization;
using System.Windows.Data;

namespace AchieveItSystemUI
{
    public class FocusStateToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (value is bool boolValue && boolValue) ? "暂停" : "开始";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}