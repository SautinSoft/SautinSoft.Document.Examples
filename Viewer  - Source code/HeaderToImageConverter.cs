using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace Viewer
{
    #region HeaderToImageConverter

    [ValueConversion(typeof(string), typeof(bool))]
    public class HeaderToImageConverter : IValueConverter
    {
        public static HeaderToImageConverter Instance = new HeaderToImageConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            PathNode pn = (value as PathNode);
            if (pn == null) return null;

            switch (pn.Type)
            {
                case 0:
                    {
                        Uri uri = new Uri("pack://application:,,,/Images/diskdrive.png");
                        BitmapImage source = new BitmapImage(uri);
                        return source;
                    }

                case 1:
                    {
                        Uri uri = new Uri("pack://application:,,,/Images/folder.png");
                        BitmapImage source = new BitmapImage(uri);
                        return source;
                    }

                default:
                    {
                        if (pn.FullPath.ToLower().EndsWith("pdf"))
                        {
                            Uri uri = new Uri("pack://application:,,,/Images/pdf.jpg");
                            BitmapImage source = new BitmapImage(uri);
                            return source;
                        }
                        else if (pn.FullPath.ToLower().EndsWith("docx") ||
                            pn.FullPath.ToLower().EndsWith("docm") ||
                            pn.FullPath.ToLower().EndsWith("dotx") ||
                            pn.FullPath.ToLower().EndsWith("dotm") ||
                            pn.FullPath.ToLower().EndsWith("rtf"))
                        {
                            Uri uri = new Uri("pack://application:,,,/Images/docx.jpg");
                            BitmapImage source = new BitmapImage(uri);
                            return source;
                        }
                        else 
                        {
                            Uri uri = new Uri("pack://application:,,,/Images/html.jpg");
                            BitmapImage source = new BitmapImage(uri);
                            return source;
                        }
                    }
            }

        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException("Cannot convert back");
        }
    }

    #endregion // DoubleToIntegerConverter
}