using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace xlsEngineApp
{
    //reference link : https://github.com/txstudio/DataMapUtility

    public class DataMapUtility
    {
        /// <summary>
        /// 將泛型物件集合換成 DataTable 物件
        /// </summary>
        /// <typeparam name="T">集合物件型態</typeparam>
        /// <param name="items">要轉換的集合物件</param>
        /// <returns>轉換後的 DataTable 物件</returns>
        public static DataTable MapToTable<T>(IEnumerable<T> items)
        {

            //取得 T 類別中的成員名稱和屬性型別
            PropertyDescriptorCollection _propertyDescriptors;
            DataTable _table;

            _propertyDescriptors = TypeDescriptor.GetProperties(typeof(T));
            _table = new DataTable();

            //新增欄位名稱和型別
            for (int i = 0; i <= _propertyDescriptors.Count - 1; i++)
            {
                PropertyDescriptor _propertyDescriptor = _propertyDescriptors[i];

                //Nullable(Of ..) 型態的屬性必須要取得指定型別

                if (_propertyDescriptor.PropertyType.ToString().Contains("Nullable") == true)
                {
                    _table.Columns.Add(_propertyDescriptor.Name
                                        , Nullable.GetUnderlyingType(_propertyDescriptor.PropertyType));
                }
                else
                {
                    _table.Columns.Add(_propertyDescriptor.Name
                                        , _propertyDescriptor.PropertyType);
                }
            }

            //逐筆新增資料
            object[] values = new object[_propertyDescriptors.Count];

            //物件為空值的話，回傳僅欄位設定的 DataTable 物件
            if (items == null)
            {
                return _table;
            }

            foreach (T item in items)
            {
                for (int i = 0; i <= values.Length - 1; i++)
                {
                    values[i] = _propertyDescriptors[i].GetValue(item);
                }
                _table.Rows.Add(values);
            }

            return _table;
        }
    }
}
