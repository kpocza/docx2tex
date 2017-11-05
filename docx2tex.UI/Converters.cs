using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Data;

namespace docx2tex.UI
{
    public class TextBoxStringBindingConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value != null && value is string && (string)value == string.Empty)
            {
                return null;
            }

            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value!= null && value is string && (string)value == string.Empty)
            {
                return null;
            }

            return value;
        }

        #endregion
    }
    public class NullStringBindingConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value != null && value is string && (string)value == string.Empty)
            {
                return null;
            }

            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value != null && value is string && (string)value == string.Empty)
            {
                return null;
            }

            return value;
        }

        #endregion
    }
}
