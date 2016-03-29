using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

    class MyMethod
    {
        public static string date2str(DateTime date)
        {
            if (date == DateTimePicker.MinimumDateTime)
            {
                return "";
            }
            return date.ToString("yyyy-MM");
        }
        public static DateTime str2Date(string str)
        {
            if (null == str || str.Trim() =="")
            {
                return DateTimePicker.MinimumDateTime;
            }
            return DateTime.Parse(str);
        }
        public static DateTime obj2Date(Object obj)
        {
            if (null == obj)
            {
                return DateTime.MinValue;
            }
            return (DateTime)obj;
        }

        public static double str2double(string str)
        {
            try
            {
                double result = double.Parse(str);
                return result;
            }
            catch
            {
                return 0;
            }
        }
        public static double obj2double(Object obj)
        {
            try
            {
                double result = double.Parse(obj.ToString());
                return result;
            }
            catch
            {
                return 0;
            }
        }
        public static int str2int(string str)
        {
            try
            {
                int result = int.Parse(str);
                return result;
            }
            catch
            {
                return 0;
            }
        }
    }