using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace orderCreateFinal
{
    public class DataSetter
    {
        public DataSet DataSet { get; set; }

        public DataSetter(DataSet dataSet)
        {
            DataSet = dataSet;
        }

        public DataTable GetDataTableByName(string tableName)
        {
            if (DataSet.Tables.Contains(tableName))
            {
                return DataSet.Tables[tableName];
            }
            else
            {
                throw new ArgumentException($"Таблица с именем '{tableName}' не найдена.");
            }
        }

        public DataRow[] GetRowsByColumnValue(string tableName, string columnName, object value)
        {
            DataTable dt = GetDataTableByName(tableName);

            if (!dt.Columns.Contains(columnName))
            {
                throw new ArgumentException($"Столбец '{columnName}' не найден в таблице '{tableName}'.");
            }
            return dt.AsEnumerable()
                     .Where(row => row[columnName].Equals(value))
                     .ToArray();
        }
        public DataRow[] GetRowsByColumnValue(DataTable dt, string columnName, object value)
        {
            if (!dt.Columns.Contains(columnName))
            {
                throw new ArgumentException($"Столбец '{columnName}' не найден.");
            }
            return dt.AsEnumerable()
                     .Where(row => row[columnName].Equals(value))
                     .ToArray();
        }

        public DataRow GetRowByColumnValue(DataTable dt, string columnName, object value)
        {
            if (!dt.Columns.Contains(columnName))
            {
                throw new ArgumentException($"Столбец '{columnName}' не найден.");
            }

            // Ищем первую строку, где значение в столбце совпадает с переданным значением
            var row = dt.AsEnumerable()
                        .FirstOrDefault(r => r[columnName].Equals(value));

            return row;
        }

        public object GetValueByColumnCondition(DataTable dt, string searchColumn, object searchValue)
        {
            if (!dt.Columns.Contains(searchColumn))
            {
                throw new ArgumentException($"Столбец '{searchColumn}' не найден.");
            }

            var row = dt.AsEnumerable()
                        .FirstOrDefault(r => r[searchColumn].Equals(searchValue));

            return row?[searchColumn];
        }

        // Поиск значения из стобца resultColumn при условии searchColumn = 'searchValue'
        public object GetValueByColumnCondition(DataTable dt, string searchColumn, object searchValue, string resultColumn)
        {
            if (!dt.Columns.Contains(searchColumn))
            {
                throw new ArgumentException($"Столбец '{searchColumn}' не найден.");
            }

            if (!dt.Columns.Contains(resultColumn))
            {
                throw new ArgumentException($"Столбец '{resultColumn}' не найден.");
            }

            // Поиск строки по условию
            var row = dt.AsEnumerable()
                        .FirstOrDefault(r => r[searchColumn].Equals(searchValue));

            return row?[resultColumn];
        }

        // Заглушка для получения данных
        public object GiveData(string name)
        {
            return new object();
        }

        // Заглушка для сохранения данных
        public void ToData(object obj) 
        {
            // Save Data
            return;
        }
    }
}
