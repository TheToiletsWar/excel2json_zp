using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel2json_zpmc
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    class JsonExporter
    {
        string mContext = "";
        int mHeaderRows = 0;
        List<Dto.TagDefinition> tDataCache= new List<Dto.TagDefinition>();
        public string context
        {
            get
            {
                return mContext;
            }
        }
        public JsonExporter(ExcelLoader excel, int headerRows, string dateFormat)
        {
            
            if (headerRows > 0)
            {
                mHeaderRows = headerRows - 1;
            }
            else
            {
                mHeaderRows = headerRows;
            }
            List<DataTable> validSheets = new List<DataTable>();
            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                DataTable sheet = excel.Sheets[i];

                //string sheetName = sheet.TableName;

                if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0)
                    validSheets.Add(sheet);
            }

            var jsonSettings = new JsonSerializerSettings
            {
                DateFormatString = dateFormat,
                Formatting = Formatting.Indented
            };

            // mutiple sheet
            // 遍历每个sheet，每个sheetname为key，数据为object存入data
            Dictionary<string, object> data = new Dictionary<string, object>();
            foreach (var sheet in validSheets)
            {
                switch (sheet.TableName)
                {
                    case "tag_list":
                        var tagdata = convertSheetToArray(sheet);
                        tagDataSolve(tagdata);
                        break;
                    case "general":

                        break;
                    case "runlogic":

                        break;
                    case "drive":
                        var result = JsonConvert.SerializeObject(convertSpecialSheetToDict(sheet, 2));

                        break;
                    case "tablist_struction":

                        break;
                    default:

                        break;
                }
            }

            //-- convert to json string
            mContext = JsonConvert.SerializeObject(data, jsonSettings);

        }
        /// <summary>
        /// 点位数据处理
        /// </summary>
        /// <param name="tagdata"></param>
        /// <returns></returns>
        private void tagDataSolve(Object tagdata)
        {
          
            for (int i = 0; i < ((List<object>)tagdata).Count; i++)
            {
                var tag = ((List<object>)tagdata)[i];
                Dto.TagDefinition t = new Dto.TagDefinition();
                Dictionary<string, object> subtagDic = ((Dictionary<string, object>)tag);
                foreach (var key in subtagDic.Keys)
                {
                    t.GetType().GetProperty(key).SetValue(t,subtagDic[key]);
                
                }
                this.tDataCache.Add(t);
            }
        }
        private object convertSheetToArray(DataTable sheet)
        {
            List<object> values = new List<object>();

            int firstDataRow = mHeaderRows;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];

                values.Add(
                    convertRowToDict(sheet, row)
                    );
            }

            return values;
        }
        private object convertSpecialSheetToDict(DataTable sheet, int skiprow)
        {
            Dictionary<string, object> importData =
               new Dictionary<string, object>();
            int firstDataRow = mHeaderRows;
            List<string> heads = new List<string>();
            string headfieldName = "heads";
            string cardtitle = sheet.Columns[0].ToString();
         
            //记录title的名称，body下的二维数组数量
            List<string> titleNames = new List<string>();

            List<int> startIndexs = new List<int>();
            List<int> endIndexs = new List<int>();
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];

                //前两行header和type数据处理

                //heads
                if ("header" == row[sheet.Columns[0]].ToString())
                {
                    for (int j = 1; j < sheet.Columns.Count; j++)
                    {
                        string value = row[sheet.Columns[j]].ToString();
                        value = getDefaultColumn(value);
                        heads.Add(value);
                    }
                    continue;
                }
                else if ("headtype" == row[sheet.Columns[0]].ToString())
                {
                    headfieldName = row[sheet.Columns[1]].ToString();
                    continue;
                }
                //body
                else if ("title" == row[sheet.Columns[0]].ToString())
                {
                    //start
                
                    titleNames.Add(row[sheet.Columns[1]].ToString());
                    startIndexs.Add(i);
                }
                else if ("type" == row[sheet.Columns[0]].ToString())
                {
                    //end
                    endIndexs.Add(i);
                }
            }
            importData.Add("data", getObjectsByIndex(sheet,startIndexs,endIndexs,titleNames));
            importData.Add("card_title", cardtitle);
            importData.Add(headfieldName, heads);
            return importData;
        }

        //表达式解析到点名
        private string expressionToTagName(string expression)
        {
            string value = expression;
            //除号之前的内容
            if (value.Contains('/'))
            {
                value = value.Substring(0,value.IndexOf('/'));
            }
            //包含函数括号并且不包含not，and，or关键字
            if (value.Contains('(') && value.Contains(')')&&!(value.Contains("or")&&value.Contains("and")&&value.Contains("not")))
            {
                value = value.Substring(value.LastIndexOf('('));
                value = value.Substring(value.IndexOf(')'));
            }
            if (value.Contains("not")||value.Contains("and"))
            {
                value = value.Replace("not", "");
                value = value.Replace("and", "");
            }
            //包含函数括号并且包括关键字
            if ((value.Contains("or") && value.Contains("and") && value.Contains("not"))&&value.Contains('(')&&value.Contains(')'))
            {
                value = value.Replace("or",",");
                var arraystr = value.Split(',');
                for (int i = 0; i < arraystr.Length; i++)
                {
                    arraystr[i] = arraystr[i].Replace('(', ' ').Replace(')', ' ');
                } 
            }
            return value.Trim();
        }
        private Dictionary<string, List<Dto.DisplayWayDefinition>> getObjectsByIndex(DataTable sheet,List<int> startindex,List<int> endindex,List<string> titlenames)
        {
            Dictionary<string, List<Dto.DisplayWayDefinition>> bodyData = new Dictionary<string, List<Dto.DisplayWayDefinition>>();
            
            //遍历出需要拿数据的行数
            for (int j = 0; j < startindex.Count; j++)
            {
                List<DataRow> datarows = new List<DataRow>();
                DataRow startrow = sheet.Rows[startindex[j]];
                DataRow endrow = sheet.Rows[endindex[j]];
                datarows.Add(startrow);
                datarows.Add(endrow);
                //centerrow
                for (int n = startindex[j] + 1; n < endindex[j]; n++)
                {
                    DataRow otherrow = sheet.Rows[n];
                    datarows.Add(otherrow);
                }
                //合并
                bodyData.Add(titlenames[j],getObjectsByRows(sheet, datarows));
            }
            return bodyData;
        }

        private List<Dto.DisplayWayDefinition> getObjectsByRows(DataTable sheet,List<DataRow> rows)
        {
            List<Dto.DisplayWayDefinition> data = new List<Dto.DisplayWayDefinition>();
            for (int i = 1; i < sheet.Columns.Count; i++)
            {
                //每一列实例化一个对象
                Dto.DisplayWayDefinition d = new Dto.DisplayWayDefinition();
                foreach (DataRow row in rows)
                {
                    //利用反射把第一列的字符串内容作为属性名获取属性,
                        d.GetType().GetProperty(row[0].ToString()).SetValue(d, row[i].ToString());
                }
                data.Add(d);
            }

            return data;
        }
      


        private static string getDefaultColumn(string value)
        {
            if (value.Equals("\"\""))
            {
                value = "";
            }

            return value;
        }




        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object convertSheetToDict(DataTable sheet)
        {
            Dictionary<string, object> importData =
                new Dictionary<string, object>();

            int firstDataRow = mHeaderRows;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];
                string ID = row[sheet.Columns[0]].ToString();
                if (ID.Length <= 0)
                    ID = string.Format("row_{0}", i);

                var rowObject = convertRowToDict(sheet, row);
                // 多余的字段
                // rowObject[ID] = ID;
                importData[ID] = rowObject;
            }

            return importData;
        }
        /// <summary>
        /// 把一行数据转换成一个对象，每一列是一个属性
        /// </summary>
        private Dictionary<string, object> convertRowToDict(DataTable sheet, DataRow row)
        {
            var rowData = new Dictionary<string, object>();
            int col = 0;
            foreach (DataColumn column in sheet.Columns)
            {
                // 过滤掉包含指定前缀的列
                string columnName = column.ToString();
                //if (excludePrefix.Length > 0 && columnName.StartsWith(excludePrefix))
                //    continue;

                object value = row[column];

                // 尝试将单元格字符串转换成 Json Array 或者 Json Object
                if (true)
                {
                    string cellText = value.ToString().Trim();
                    if (cellText.StartsWith("[") || cellText.StartsWith("{"))
                    {
                        try
                        {
                            object cellJsonObj = JsonConvert.DeserializeObject(cellText);
                            if (cellJsonObj != null)
                                value = cellJsonObj;
                        }
                        catch (Exception exp)
                        {
                        }
                    }
                }

                if (value.GetType() == typeof(System.DBNull))
                {
                    value = "";
                    //value = getColumnDefault(sheet, column, firstDataRow);

                }
                else if (value.GetType() == typeof(double))
                { // 去掉数值字段的“.0”
                    double num = (double)value;
                    if ((int)num == num)
                        value = (int)num;
                }

                //全部转换为string
                //方便LitJson.JsonMapper.ToObject<List<Dictionary<string, string>>>(textAsset.text)等使用方式 之后根据自己的需求进行解析
                //if (allString && !(value is string))
                //{
                //    value = value.ToString();
                //}

                string fieldName = column.ToString();
                // 表头自动转换成小写
                //if (lowcase)
                //    fieldName = fieldName.ToLower();

                if (string.IsNullOrEmpty(fieldName))
                    fieldName = string.Format("col_{0}", col);

                rowData[fieldName] = value;
                col++;
            }

            return rowData;
        }
        /// <summary>
        /// 对于表格中的空值，找到一列中的非空值，并构造一个同类型的默认值
        /// </summary>
        private object getColumnDefault(DataTable sheet, DataColumn column, int firstDataRow)
        {
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                object value = sheet.Rows[i][column];
                Type valueType = value.GetType();
                if (valueType != typeof(System.DBNull))
                {
                    if (valueType.IsValueType)
                        return Activator.CreateInstance(valueType);
                    break;
                }
            }
            return "";
        }
        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public void SaveToFile(string filePath, Encoding encoding)
        {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mContext);
            }
        }
    }
}
