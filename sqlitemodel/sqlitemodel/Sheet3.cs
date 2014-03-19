/*
The MIT License (MIT)

Copyright (c) <2013-2020> <wenshengming zhujiangping>

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

.
*/

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Data.SQLite;

namespace sqlitemodel
{
    public partial class Sheet3
    {
        struct _Map_Col_Info
        {
            public string sColOrgin;
            public string sColMap;
            public string sMapTypeName;           
        };
        
        public DataSet ds;
        public SQLiteDataAdapter adpater;
        public  Microsoft.Office.Tools.Excel.ListObject list1;


        private string[] m_strArrayFieldName;
        private string[] m_strArrayFieldDBName;
        private string[] m_strArrayMapTypeName;
        private int m_nFieldCount;

        private bool m_bIsLoading = false;
        //  保存映射值
        private Dictionary<string, Dictionary<int, string> > m_DicMapType = new Dictionary<string,Dictionary<int,string>>();

        //private Dictionary<string, _Map_Col_Info> m_OrginToMapInfo = new Dictionary<string,_Map_Col_Info>();     //  原始值列对应信息
        private Dictionary<string, _Map_Col_Info> m_MapToMapInfo = new Dictionary<string,_Map_Col_Info>();       //  映射值对应信息

        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet3_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void clear()
        {
            //  移除Valid信息
            foreach (KeyValuePair<string, _Map_Col_Info> keyValues in m_MapToMapInfo)
            {
                string sColName = keyValues.Key;
                string sColNameOrgin = keyValues.Value.sColOrgin;

                Excel.Range xRan;
                xRan = this.Range[sColName + Convert.ToString(7), sColName + Convert.ToString(this.list1.ListRows.Count + 7)];

                MessageBox.Show("解决0x800A03EC错误，目前没有找到可以退出编辑模式的方法，只能出这个提示框，才能正确设置！");
                xRan.Validation.Delete();
                xRan.Interior.ColorIndex = 0;


                xRan = this.Range[sColNameOrgin + Convert.ToString(7), sColNameOrgin + Convert.ToString(this.list1.ListRows.Count + 7)];
                xRan.Interior.ColorIndex = 0;
            }

            if (list1 != null)
            {
                this.Controls.Remove(list1);
            }
            if (adpater != null)
            {
                adpater = null;
            }
            if (ds != null)
            {
                ds = null;
            }
            m_nFieldCount = 0;
            m_strArrayFieldName = null;
            m_strArrayFieldDBName = null;
            m_strArrayMapTypeName = null;
            m_DicMapType.Clear();
            //m_OrginToMapInfo.Clear();
            m_MapToMapInfo.Clear();            
        }


        // Convert A, B, .... AAA, AAB to number 1,2, ...          
        int StringToNumber(string s)
        {
            int r = 0;
            for (int i = 0; i < s.Length; i++)
            {
                r = r * 26 + s[i] - 'A' + 1;
            }
            return r;
        }

        // Convert number 1, 2, ... to string A, B, ...   
        static string NumbertoString(int n)
        {
            string s = "";     // result   
            int r = 0;         // remainder   

            while (n != 0)
            {
                r = n % 26;
                char ch = ' ';
                if (r == 0)
                    ch = 'Z';
                else
                    ch = (char)(r - 1 + 'A');
                s = ch.ToString() + s;
                if (s[0] == 'Z')
                    n = n / 26 - 1;
                else
                    n /= 26;
            }
            return s;
        } 

        /*
 * 连接数据库的代码
 */
        public void BindData(string strTableName, int nFieldCount, string[] strArrayFieldName, string[] strArrayFieldDBName, string[] strArrayMapTypeName)
        {
            m_bIsLoading = true;
            m_nFieldCount                           = nFieldCount;
            m_strArrayFieldName                 = strArrayFieldName;
            m_strArrayFieldDBName             = strArrayFieldDBName;
            m_strArrayMapTypeName          = strArrayMapTypeName;

            //  确定listobject里面显示的列数
            int nColumnCount = 0;
            for (int i = 0; i < nFieldCount; ++i)
            {
                if (m_strArrayMapTypeName[i] != "")
                {
                    nColumnCount++;

                    //  读取映射值
                    SQLiteCommand sCommand = Globals.Sheet1.connection.CreateCommand();
                    sCommand.CommandText = "select mapoldvalue, mapvalue from mapdefine where maptype='" + m_strArrayMapTypeName[i] + "'";
            
                    SQLiteDataReader reader = sCommand.ExecuteReader();

                    Dictionary<int, string> dic = null;
                    if (!m_DicMapType.ContainsKey(m_strArrayMapTypeName[i]))
                    {
                        dic = new Dictionary<int, string>();
                        m_DicMapType.Add(m_strArrayMapTypeName[i], dic);
                    }
                    else
                    {
                        dic = m_DicMapType[m_strArrayMapTypeName[i]];
                    }                   
                  
                    while (reader.Read())
                    {                       
                        int nOldValue = reader.GetInt32(0);
                        string sValue = reader.GetString(1);
                        dic[nOldValue] = sValue;
                    }
                    reader.Close();
                }
            }
            //  RecordOrder字段不在定义表中出现
            nColumnCount += m_nFieldCount + 1;
            //  生成数据列的绑定字段字符串
            string[] mappedColumn = new string[nColumnCount];
            int nColumnIndex = 0;
            mappedColumn[nColumnIndex] = "RecordOrder";
            nColumnIndex++;

            for (int i = 0; i < m_nFieldCount; ++i)
            {              
                    mappedColumn[nColumnIndex] = m_strArrayFieldDBName[i];
                    nColumnIndex++;
                    if (m_strArrayMapTypeName[i] != "")
                    {
                        mappedColumn[nColumnIndex] = "";
                        nColumnIndex++;
                    }      
            }
            
            //  读取记录集
            String sql = @"select * from [" + strTableName +"]"; 
            ds = new DataSet();
            adpater = new SQLiteDataAdapter(sql,  Globals.Sheet1.connection);
            adpater.Fill(ds);

            int nRecordCount = ds.Tables[0].Rows.Count;

            // 创建listobject对象
            list1 =
                this.Controls.AddListObject(
                this.Range["A6", missing], strTableName);


            list1.AutoSetDataBoundColumnHeaders = false;

            //  设置绑定字段
            list1.SetDataBinding(ds, ds.Tables[0].TableName, mappedColumn);


            
            //  设置列名
            nColumnIndex = 1;
            this.Range[NumbertoString(nColumnIndex) + Convert.ToString(1), System.Type.Missing].ColumnWidth = 25;

            list1.ListColumns.Item[nColumnIndex].Name = "RecordOrder";
            nColumnIndex++;
            this.Range[NumbertoString(nColumnIndex) + Convert.ToString(1), System.Type.Missing].ColumnWidth = 25;

            for (int i = 0; i < m_nFieldCount; ++i)
            {                
                list1.ListColumns.Item[nColumnIndex].Name = m_strArrayFieldName[i];
                this.Range[NumbertoString(nColumnIndex) + Convert.ToString(1), System.Type.Missing].ColumnWidth = 25;
                nColumnIndex++;
                

                if (m_strArrayMapTypeName[i] != "")
                {
                    //string sOrginValue = "";
                    string sMapValue = "";

                    list1.ListColumns.Item[nColumnIndex].Name = m_strArrayFieldName[i] + "的映射值";
                    this.Range[NumbertoString(nColumnIndex) + Convert.ToString(1), System.Type.Missing].ColumnWidth = 25;
                    nColumnIndex++;
                     Dictionary<int, string> dic = m_DicMapType[m_strArrayMapTypeName[i]];
                     if (dic != null)
                     {
                         //int[] nkeys1 = dic.Keys.ToArray<int>();
                         //for (int m = 0; m < nkeys1.Length; ++m)
                         //{
                         //    if (sOrginValue == "")
                         //    {
                         //        sOrginValue += nkeys1[m].ToString();
                         //    }
                         //    else
                         //    {
                         //        sOrginValue += "," + nkeys1[m].ToString();
                         //    }
                         //}
                         string[] sMapValue1 = dic.Values.ToArray<string>();
                         for (int m = 0; m < sMapValue1.Length; ++m)
                         {
                             if (sMapValue == "")
                             {
                                 sMapValue += sMapValue1[m];
                             }
                             else
                             {
                                 sMapValue += "," + sMapValue1[m];
                             }
                         }
                     }

                    //  设置伙伴颜色   
                    string sColNameMap = NumbertoString(nColumnIndex - 1);
                    string sColNameOrign = NumbertoString(nColumnIndex - 2);
                    //char cColNameMap = (char)('A' + nColumnIndex - 2);
                    //char cColNameOrign = (char)('A' + nColumnIndex - 3);

                    Excel.Range xRan;
                    xRan = this.Range[sColNameMap + Convert.ToString(7), sColNameMap + Convert.ToString(nRecordCount + 7)];
                    xRan.Interior.ColorIndex = 46;

                    MessageBox.Show("解决0x800A03EC错误，目前没有找到可以退出编辑模式的方法，只能出这个提示框，才能正确设置！");
                    xRan.Validation.Delete();
                    xRan.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertInformation, Excel.XlFormatConditionOperator.xlBetween, sMapValue, System.Type.Missing);
                    xRan.Validation.IgnoreBlank = true;
                    xRan.Validation.InCellDropdown = true;
                    

                    xRan = this.Range[sColNameOrign + Convert.ToString(7), sColNameOrign + Convert.ToString(nRecordCount + 7)];
                    xRan.Interior.ColorIndex = 47;

                    //MessageBox.Show("解决0x800A03EC错误，目前没有找到可以退出编辑模式的方法，只能出这个提示框，才能正确设置！");
                    //xRan.Validation.Delete();
                    //xRan.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertInformation, Excel.XlFormatConditionOperator.xlBetween, sOrginValue, System.Type.Missing);
                    //xRan.Validation.IgnoreBlank = true;
                    //xRan.Validation.InCellDropdown = true;      
                  
                    
                    _Map_Col_Info mci;
                    mci.sColOrgin = sColNameOrign;
                    mci.sColMap = sColNameMap;
                    mci.sMapTypeName = m_strArrayMapTypeName[i];

                    m_MapToMapInfo.Add(sColNameMap, mci);
                    //m_MapToMapInfo.Add(sColNameMap, mci);

                    //  更新内容
                    //Dictionary<int, string> dic = m_DicMapType[m_strArrayMapTypeName[i]];
                    if (dic != null)
                    {
                        for (int nRIndex = 0; nRIndex < ds.Tables[0].Rows.Count; ++nRIndex)
                        {                            
                            string strOrginValue = ((Excel.Range)this.Cells[nRIndex + 7, nColumnIndex - 2]).Text;
                            int nValue = Convert.ToInt32(strOrginValue);

                            string sMappedValue = "错误映射类型";
                            if (dic.ContainsKey(nValue))
                            {
                                sMappedValue = dic[nValue];
                            }
                            this.Cells[nRIndex + 7, nColumnIndex - 1] = sMappedValue;
                        }
                    }
                }
            }
            //this.Range[NumbertoString(nColumnIndex) + Convert.ToString(1), System.Type.Missing].ColumnWidth = 25;
            MessageBox.Show("数据装载成功！！");
            m_bIsLoading = false;
        }
        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Sheet3_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet3_Shutdown);
            this.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(this.Sheet3_Change);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            if (adpater == null || ds == null)
            {
                return;
            }
            SQLiteTransaction sqTrans = Globals.Sheet1.connection.BeginTransaction();
            try
            {                               
                 SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adpater);
                adpater.DeleteCommand = builder.GetDeleteCommand();
                adpater.UpdateCommand = builder.GetUpdateCommand();
                adpater.Update(ds.Tables[0]);
                ds.AcceptChanges();
                sqTrans.Commit();
                MessageBox.Show("数据保存成功", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(SystemException err)
            {
                sqTrans.Rollback();
                MessageBox.Show(err.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }            
        }
        //  sheet2更新完之后，更新这里存在的映射
        public void UpdateMapTypeDefine(string strMapTypeName)
        {
            if (!m_DicMapType.ContainsKey(strMapTypeName))
            {
                return;
            }

            //  更新映射信息
            Dictionary<int, string> dic = m_DicMapType[strMapTypeName];
            dic.Clear();


            SQLiteCommand sCommand = Globals.Sheet1.connection.CreateCommand();
            sCommand.CommandText = "select mapoldvalue, mapvalue from mapdefine where maptype='" + strMapTypeName + "'";

            SQLiteDataReader reader = sCommand.ExecuteReader();       

            while (reader.Read())
            {
                int nOldValue = reader.GetInt32(0);
                string sValue = reader.GetString(1);
                dic[nOldValue] = sValue;
            }
            reader.Close();

            //  更新数据有效性信息         
            foreach (KeyValuePair<string, _Map_Col_Info> keyValues in m_MapToMapInfo)
            {
                string sColName = keyValues.Key;
                if (keyValues.Value.sMapTypeName == strMapTypeName)
                {
                    string sOrginValue = "";
                    string[] nkeys1 = dic.Values.ToArray<string>();
                    for (int m = 0; m < nkeys1.Length; ++m)
                    {
                        if (sOrginValue == "")
                        {
                            sOrginValue += nkeys1[m].ToString();
                        }
                        else
                        {
                            sOrginValue += "," + nkeys1[m].ToString();
                        }
                    }


                    Excel.Range xRan;
                    xRan = this.Range[sColName + Convert.ToString(7), sColName + Convert.ToString(this.list1.ListRows.Count + 7)];                    

                    MessageBox.Show("解决0x800A03EC错误，目前没有找到可以退出编辑模式的方法，只能出这个提示框，才能正确设置！");
                    xRan.Validation.Delete();
                    xRan.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertInformation, Excel.XlFormatConditionOperator.xlBetween, sOrginValue, System.Type.Missing);
                    xRan.Validation.IgnoreBlank = true;
                    xRan.Validation.InCellDropdown = true;      
                }
            }
        }

        private void Sheet3_Change(Excel.Range Target)
        {
            if (!m_bIsLoading)
            {
                foreach (Excel.Range rng in Target)
                {
                    int nCol = rng.Column;
                    int nRow = rng.Row;

                    string sColName = NumbertoString(nCol);
                    if (m_MapToMapInfo.ContainsKey(sColName))
                    {
                        _Map_Col_Info sMapColInfo = m_MapToMapInfo[sColName];
                        if (m_DicMapType.ContainsKey(sMapColInfo.sMapTypeName))
                        {
                            Dictionary<int, string> dic = m_DicMapType[sMapColInfo.sMapTypeName];

                            string sValue = (string)rng.Value2;
                            foreach (KeyValuePair<int, string> kvalue in dic)
                            {
                                if (kvalue.Value == sValue)
                                {
                                    int nColToModify = StringToNumber(sMapColInfo.sColOrgin);
                                    Excel.Range rngToModify = this.Cells[nRow, nColToModify];
                                    rngToModify.Value2 = kvalue.Key;
                                    break;
                                }
                            }            
                        }
                        continue;
                    }
                    
                    //MessageBox.Show(
                    //   rng.get_Address(System.Type.Missing, System.Type.Missing,
                    //   Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing));
                }
            }
        }
        public void OnMapTypeDeleted(string sMapTypeName)
        {
            //  移除映射类型信息
            if (!m_DicMapType.ContainsKey(sMapTypeName))
            {
                return;
            }
            m_DicMapType.Remove(sMapTypeName);

            //  移除Valid信息
            foreach (KeyValuePair<string, _Map_Col_Info> keyValues in m_MapToMapInfo)
            {
                string sColName = keyValues.Key;
                if (keyValues.Value.sMapTypeName == sMapTypeName)
                {
                    Excel.Range xRan;
                    xRan = this.Range[sColName + Convert.ToString(7), sColName + Convert.ToString(this.list1.ListRows.Count + 7)];

                    MessageBox.Show("解决0x800A03EC错误，目前没有找到可以退出编辑模式的方法，只能出这个提示框，才能正确设置！");
                    xRan.Validation.Delete();
                }
            }
        }
    }
}
