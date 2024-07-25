using System.Text;
using System.Reflection;
using System.ComponentModel;

namespace DG.XrmOrg.XrmSolution.ConsoleJobs.Helpers
{
    internal static class CsvHelper<T> where T : class, new()
    {
        public static void WriteToCsv(string filePath, IEnumerable<T> items, bool appendToFileIfExists = false)
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
                    writer.WriteLine(GetCsvHeader());
                }
                foreach (var item in items)
                {
                    writer.WriteLine(ToCsv(item));
                }
            }
            finally
            {
                writer.Close();
            }
        }

        public static IEnumerable<T> ReadFromCsv(string filePath, bool hasHeader = true)
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
                    returnList.Add(FromCsv(res));
                    res = reader.ReadLine();
                }
            }
            finally
            {
                reader.Close();
            }

            return returnList;
        }

        public static string GetCsvHeader()
        {
            var properties = typeof(T).GetProperties();
            var strings = new string[properties.Length];
            for (int i = 0; i < properties.Length; ++i)
            {
                strings[i] = properties[i].Name;
            }

            return string.Join(ConfigurationManager.AppSettings["CsvSeparator"], strings);
        }

        public static string ToCsv(T item)
        {
            var properties = typeof(T).GetProperties();
            var strings = new string[properties.Length];
            for (int i = 0; i < properties.Length; ++i)
            {
                strings[i] = properties[i].GetValue(item)?.ToString() ?? "";
            }

            return string.Join(ConfigurationManager.AppSettings["CsvSeparator"], strings);
        }

        public static T FromCsv(string item)
        {
            var attributeStrings = item.Split(ConfigurationManager.AppSettings["CsvSeparator"]);
            var properties = typeof(T).GetProperties();

            T returnval = new();
            for (int i = 0; i < attributeStrings.Length; ++i)
            {
                var property = properties[i];
                object? value = TypeConverter.FromString(property.PropertyType, attributeStrings[i]);
                property.SetValue(returnval, value);
            }
            return returnval;
        }

        private static object? FromString(Type type, string s)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Int32:
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