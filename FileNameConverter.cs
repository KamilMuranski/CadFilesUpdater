using System;
using System.Globalization;
using System.IO;
using System.Windows.Data;

namespace CadFilesUpdater
{
    public class FileNameConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string filePath)
            {
                return Path.GetFileNameWithoutExtension(filePath);
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
