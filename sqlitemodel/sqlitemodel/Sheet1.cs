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
    public partial class Sheet1
    {
        public Microsoft.Office.Tools.Excel.ListObject sheet1ListObject;        
        public SQLiteConnection connection;
        public DataSet ds;
        public SQLiteDataAdapter adpater;

        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.btnopendb.Click += new System.EventHandler(this.btnopendb_Click);
            this.dgvTableDefine.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            BindData();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnopendb_Click(object sender, EventArgs e)
        {

        }
        /*
         * 连接数据库的代码
         */
        public void BindData()
        {
            // 创建连接数据库
            String connectionString = @"Data Source=D:\Dev\All-Totorial\SQLites\OrderWater.db;Version=3;";
            String sql = @"select * from All_Customer Where   (1=1)   order by LastUpdated DESC LIMIT 0,20";
            connection = new SQLiteConnection(connectionString);
            //SQLiteCommand cmd = new SQLiteCommand(sql, connection);

            connection.Open();
            ds = new DataSet();
            adpater = new SQLiteDataAdapter(sql, connectionString);
            adpater.Fill(ds);

            // Create a list object.
            Microsoft.Office.Tools.Excel.ListObject list1 =
                this.Controls.AddListObject(
                this.Range["A1", missing], "Customers");

            System.Diagnostics.Trace.WriteLine("{1}", ds.Tables.Count.ToString());
            // Bind the list object to the Customers table.
            list1.AutoSetDataBoundColumnHeaders = true;
            list1.DataSource = ds.Tables[0];

            foreach (DataTable tb in ds.Tables)
            {
                System.Diagnostics.Debug.WriteLine(tb.ToString());
                System.Diagnostics.Debug.WriteLine(tb.TableName);
                foreach (DataColumn col in tb.Columns)
                {
                    System.Diagnostics.Debug.WriteLine(col.ColumnName);
                }
            }
            //list1.DataMember = "All_Customer";
        }
    }
}
