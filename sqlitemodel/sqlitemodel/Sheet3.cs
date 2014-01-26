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
        }
        /*
 * 连接数据库的代码
 */
        public void BindData(string strTableName)
        {
            // 创建连接数据库
            /*
            String connectionString = @"Data Source=D:\Dev\All-Totorial\SQLites\OrderWater.db;Version=3;";
             * 
             */
            String sql = @"select * from [" + strTableName +"]";
            //connection = new SQLiteConnection(connectionString);
            //SQLiteCommand cmd = new SQLiteCommand(sql, connection);
            //connection.Open();
            ds = new DataSet();
            adpater = new SQLiteDataAdapter(sql,  Globals.Sheet1.connection);
            adpater.Fill(ds);
            /*
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adpater);
            adpater.DeleteCommand = builder.GetDeleteCommand();
            adpater.UpdateCommand = builder.GetUpdateCommand();
            adpater.Update(ds.Tables[0]);
            ds.AcceptChanges();
             * 
             */

            // Create a list object.
            list1 =
                this.Controls.AddListObject(
                this.Range["A6", missing], strTableName);

            //System.Diagnostics.Trace.WriteLine("{1}", ds.Tables.Count.ToString());
            // Bind the list object to the Customers table.
            list1.AutoSetDataBoundColumnHeaders = true;
            list1.DataSource = ds.Tables[0];

            
            /*
            foreach (DataTable tb in ds.Tables)
            {
                System.Diagnostics.Debug.WriteLine(tb.ToString());
                System.Diagnostics.Debug.WriteLine(tb.TableName);
                foreach (DataColumn col in tb.Columns)
                {
                    System.Diagnostics.Debug.WriteLine(col.ColumnName);
                }
            }
             */
            //list1.DataMember = "All_Customer";
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

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            if (adpater == null || ds == null)
            {
                return;
            }
            try
            {
                 SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adpater);
                adpater.DeleteCommand = builder.GetDeleteCommand();
                adpater.UpdateCommand = builder.GetUpdateCommand();
                adpater.Update(ds.Tables[0]);
                ds.AcceptChanges();
                MessageBox.Show("数据保存成功", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(SystemException err)
            {
                MessageBox.Show(err.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
    }
}
