using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace DG.XrmOrg.XrmSolution.ConsoleJobs.Helpers
{
    internal static class CsvHelper
    {
        public static void WriteToCsv<T>(string filePath, IEnumerable<T> items, bool appendToFileIfExists = false) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path is empty");
            }

            StreamWriter writer = new StreamWriter(filePath, appendToFileIfExists, Encoding.UTF8);
            try
            {
                if (!appendToFileIfExists)
                {
                    writer.WriteLine(GetCsvHeader<T>());
                }
                foreach (var item in items)
                {
                    writer.WriteLine(RowToCsv(item));
                }
            }
            finally
            {
                writer.Close();
            }
        }

        public static List<T> ReadFromCsv<T>(string filePath, bool hasHeader = true) where T : class, new()
        {
            StreamReader reader = new StreamReader(filePath, Encoding.UTF8);

            var returnList = new List<T>();

            try
            {
                var res = reader.ReadLine();
                if (hasHeader && res != null)
                {
                    res = reader.ReadLine();
                }
                while (res != null)
                {
                    returnList.Add(RowFromCsv<T>(res));
                    res = reader.ReadLine();
                }
            }
            finally
            {
                reader.Close();
            }

            return returnList;
        }

        public static string GetCsvHeader<T>() where T : class, new()
        {
            var properties = typeof(T).GetProperties();
            var strings = new string[properties.Length];
            for (int i = 0; i < properties.Length; ++i)
            {
                strings[i] = properties[i].Name;
            }

            return string.Join(ConfigurationManager.AppSettings["CsvSeparator"], strings);
        }

        public static string RowToCsv<T>(T item) where T : class, new()
        {
            var properties = typeof(T).GetProperties();
            var strings = new string[properties.Length];
            for (int i = 0; i < properties.Length; ++i)
            {
                strings[i] = properties[i].GetValue(item)?.ToString() ?? "";
            }

            return string.Join(ConfigurationManager.AppSettings["CsvSeparator"], strings);
        }

        public static T RowFromCsv<T>(string item) where T : class, new()
        {
            var attributeStrings = item.Split(ConfigurationManager.AppSettings["CsvSeparator"].ToArray());
            var properties = typeof(T).GetProperties();

            T returnval = new T();
            for (int i = 0; i < attributeStrings.Length; ++i)
            {
                var property = properties[i];
                object value = FromString(property.PropertyType, attributeStrings[i]);
                property.SetValue(returnval, value);
            }
            return returnval;
        }

        private static object FromString(Type type, string s)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Int32:
                    if (type.IsEnum)
                    {
                        return Enum.Parse(type, s);
                    }
                    return int.Parse(s);
                case TypeCode.Int64:
                    return long.Parse(s);
                case TypeCode.String:
                    return s;
                case TypeCode.Boolean:
                    return bool.Parse(s);
                case TypeCode.DateTime:
                    return DateTime.Parse(s);
                case TypeCode.Double:
                    return double.Parse(s);
                case TypeCode.Decimal:
                    return decimal.Parse(s);
                case TypeCode.Object:
                    if (Nullable.GetUnderlyingType(type) != null && !string.IsNullOrEmpty(s))
                    {
                        return FromString(Nullable.GetUnderlyingType(type), s);
                    }

                    if (Guid.TryParse(s, out Guid res))
                    {
                        return res;
                    }
                    else
                    {
                        return TypeDescriptor.GetConverter(type).ConvertFromString(s);

                    }
                default:
                    return TypeDescriptor.GetConverter(type).ConvertFromString(s);
            }
        }
    }
}