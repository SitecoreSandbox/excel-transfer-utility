#region Namespaces

using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.IO;
using System.Text; 

#endregion

namespace ExcelTransferUtility.sitecore_modules.Web.Excel_Transfer_Utility.Export
{
    /// <summary>
    /// Exports Sitecore items from Sitecore to .csv file
    /// </summary>
    public class CsvExport
    {
        private readonly List<string> _fields = new List<string>();
        private readonly List<Dictionary<string, object>> _rows = new List<Dictionary<string, object>>();
        private Dictionary<string, object> CurrentRow => _rows[_rows.Count - 1];

        public object this[string field]
        {
            set
            {
                if (!_fields.Contains(field))
                {
                    _fields.Add(field);
                }
                CurrentRow[field] = value;
            }
        }

        public void AddRow()
        {
            _rows.Add(new Dictionary<string, object>());
        }

        public string Export()
        {
            var builder = new StringBuilder();
            foreach (var str in _fields)
            {
                builder.Append(str).Append(",");
            }
            builder.AppendLine();
            foreach (var dictionary in _rows)
            {
                foreach (var str2 in _fields)
                {
                    builder.Append(MakeValueCsvFriendly(dictionary[str2])).Append(",");
                }
                builder.AppendLine();
            }
            return builder.ToString();
        }

        public byte[] ExportToBytes() =>
            Encoding.UTF8.GetBytes(Export());

        public void ExportToFile(string path)
        {
            File.WriteAllText(path, Export());
        }

        private static string MakeValueCsvFriendly(object value)
        {
            if (value == null)
            {
                return "";
            }

            var nullable = value as INullable;

            if (nullable != null && nullable.IsNull)
            {
                return "";
            }

            if (value is DateTime)
            {
                var time = (DateTime)value;
                if (time.TimeOfDay.TotalSeconds == 0.0)
                {
                    var time2 = (DateTime)value;
                    return time2.ToString("yyyy-MM-dd");
                }
                var time3 = (DateTime)value;
                return time3.ToString("yyyy-MM-dd HH:mm:ss");
            }

            var str = value.ToString();

            if (!str.Contains(",") && !str.Contains("\""))
            {
                return str;
            }

            return '"' + str.Replace("\"", "\"\"") + '"';
        }
    }
}