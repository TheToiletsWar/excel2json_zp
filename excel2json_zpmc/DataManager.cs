using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace excel2json_zpmc
{
    class DataManager
    {
        // 数据导入设置
        private Encoding mEncoding;

        // 导出数据
        private JsonExporter mJson;

        /// <summary>
        /// 导出的Json文本
        /// </summary>
        public string JsonContext
        {
            get
            {
                if (mJson != null)
                    return mJson.context;
                else
                    return "";
            }
        }

        /// <summary>
        /// 保存Json
        /// </summary>
        /// <param name="filePath">保存路径</param>
        public void saveJson(string filePath)
        {
            if (mJson != null)
            {
                mJson.SaveToFile(filePath, mEncoding);
            }
        }

        //public void saveCSharp(string filePath)
        //{
        //    if (mCSharp != null)
        //        mCSharp.SaveToFile(filePath, mEncoding);
        //}


        /// <summary>
        /// 加载Excel文件
        /// </summary>
        /// <param name="options">导入设置</param>
        public void loadExcel(string filepath,string encoding)
        {

            //-- Excel File
            string excelPath = filepath;
            string excelName = Path.GetFileNameWithoutExtension(excelPath);

            //-- Header
            int header = 0;

            //-- Encoding
            Encoding cd = new UTF8Encoding(false);
            if (encoding != "utf8-nobom")
            {
                foreach (EncodingInfo ei in Encoding.GetEncodings())
                {
                    Encoding e = ei.GetEncoding();
                    if (e.HeaderName == encoding)
                    {
                        cd = e;
                        break;
                    }
                }
            }
            mEncoding = cd;

            //-- Load Excel
            ExcelLoader excel = new ExcelLoader(excelPath, header);

         
            //-- 导出JSON
            mJson = new JsonExporter(excel,header, "yyyy/MM/dd");
        }
    }
}
