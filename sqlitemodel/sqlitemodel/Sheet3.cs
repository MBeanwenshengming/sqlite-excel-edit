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

        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet3_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void clear()
        {
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
            list1.ListColumns.Item[nColumnIndex].Name = "RecordOrder";
            nColumnIndex++;

            for (int i = 0; i < m_nFieldCount; ++i)
            {
                list1.ListColumns.Item[nColumnIndex].Name = m_strArrayFieldName[i];
                nColumnIndex++;
                if (m_strArrayMapTypeName[i] != "")
                {
                    list1.ListColumns.Item[nColumnIndex].Name = m_strArrayFieldName[i] + "的映射值";
                    nColumnIndex++;

                    //  设置伙伴颜色
                    char cColNameMap = (char)('A' + nColumnIndex - 2);
                    char cColNameOrign = (char)('A' + nColumnIndex - 3);
                    Excel.Range xRan;
                    xRan = this.Range[cColNameMap.ToString() + Convert.ToString(6), cColNameMap.ToString() + Convert.ToString(nRecordCount + 6)];
                    xRan.Interior.ColorIndex = 46;

                    xRan = this.Range[cColNameOrign.ToString() + Convert.ToString(6), cColNameOrign.ToString() + Convert.ToString(nRecordCount + 6)];
                    xRan.Interior.ColorIndex = 47;

                    //  更新内容
                    Dictionary<int, string> dic = m_DicMapType[m_strArrayMapTypeName[i]];
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

        private void Sheet3_Change(Excel.Range Target)
        {
            if (!m_bIsLoading)
            {
                MessageBox.Show("sf");
            }
        }        
    }
}
