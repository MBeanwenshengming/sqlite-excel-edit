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
using System.Text.RegularExpressions;
using System.Data.SQLite;

namespace sqlitemodel
{
    public partial class Sheet2
    {
        private void Sheet2_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet2_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.btnSaveToDB.Click += new System.EventHandler(this.btnSaveToDB_Click);
            this.btnGetFromDB.Click += new System.EventHandler(this.btnGetFromDB_Click);
            this.btnBeginCreate.Click += new System.EventHandler(this.btnBeginCreate_Click);
            this.Startup += new System.EventHandler(this.Sheet2_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet2_Shutdown);

        }

        #endregion

        private void btnSaveToDB_Click(object sender, EventArgs e)
        {
            if (Globals.Sheet1.connection == null)
            {
                MessageBox.Show("当前的数据库不处于打开状态，无法创建映射类型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Globals.Sheet1.connection.State != ConnectionState.Open)
            {
                MessageBox.Show("当前的数据库不处于打开状态，无法创建映射类型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (this.cboMapType.Text == "")
            {
                MessageBox.Show("映射名为空，无法将映射保存到数据库，请先填写映射名", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (this.dgvMapInfoOfCurType.RowCount == 1)
            {
                MessageBox.Show("不存在映射定义，请先填写映射信息，然后保存到数据库", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //  首先，检查映射名称是否存在，如果存在，不能创建映射类型
            System.Data.SQLite.SQLiteCommand sCommand = Globals.Sheet1.connection.CreateCommand();
            sCommand.CommandText = "select * from mapdefine where maptype='" + this.cboMapType.Text + "'"; 
            int nResult = sCommand.ExecuteNonQuery();
            if (nResult != 0)
            {
                MessageBox.Show("已经存在映射类型，无法创建统一名称的映射类型", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //  检查datagridview中设置的映射数据是否合法
            for (int i = 1; i < this.dgvMapInfoOfCurType.RowCount; ++i)
            {
                for (int j = 1; j < this.dgvMapInfoOfCurType.ColumnCount; ++j)
                {              
                    if (this.dgvMapInfoOfCurType.Rows[i - 1].Cells[j - 1].Value == null)
                    {
                        MessageBox.Show("映射定义存在非法值，请检查后再保存", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (j == 1)
                    {
                        string strValue = this.dgvMapInfoOfCurType.Rows[i - 1].Cells[j - 1].Value.ToString();
                        try
                        {
                            Int64.Parse(strValue);
                        }
                        catch (Exception excep)
                        {
                            MessageBox.Show("映射原值有不为数字的错误", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }                        
                    }
                }
            }

            

            //  保存到数据库中
            //  开启一个事务
            System.Data.SQLite.SQLiteTransaction trans = Globals.Sheet1.connection.BeginTransaction();

            for (int i = 1; i < this.dgvMapInfoOfCurType.RowCount; ++i)
            {
                string sInsertSql = "insert into mapdefine values('" + this.cboMapType.Text + "'";
                for (int j = 1; j <= this.dgvMapInfoOfCurType.ColumnCount; ++j)
                {   
                    if (j == 1)
                    {
                        sInsertSql += "," + this.dgvMapInfoOfCurType.Rows[i - 1].Cells[j - 1].Value.ToString();
                    }
                    else
                    {
                        sInsertSql += ",'" + this.dgvMapInfoOfCurType.Rows[i - 1].Cells[j - 1].Value.ToString() + "'";
                    }
                }
                sInsertSql += ")";

                sCommand.CommandText = sInsertSql;
                int nResult1 = sCommand.ExecuteNonQuery();
            }
            //  提交事务
            trans.Commit();

            //  sheet1可用映射类型重新获取
            DataGridViewRow dr = new DataGridViewRow();
            dr.CreateCells(Globals.Sheet1.dgvAvailableMapType);
            dr.Cells[0].Value = this.cboMapType.Text;

            Globals.Sheet1.dgvAvailableMapType.Rows.Add(dr);


            this.dgvMapInfoOfCurType.ReadOnly = true;
            this.btnSaveToDB.Enabled = false;
            this.cboMapType.DropDownStyle = ComboBoxStyle.DropDownList;           
        }

        private void btnBeginCreate_Click(object sender, EventArgs e)
        {
            this.dgvMapInfoOfCurType.ReadOnly = false;
            this.btnSaveToDB.Enabled = true;
            this.cboMapType.DropDownStyle = ComboBoxStyle.DropDown;
        }

        private void btnGetFromDB_Click(object sender, EventArgs e)
        {
            if (Globals.Sheet1.connection == null)
            {
                MessageBox.Show("当前的数据库不处于打开状态，无法创建映射类型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Globals.Sheet1.connection.State != ConnectionState.Open)
            {
                MessageBox.Show("当前的数据库不处于打开状态，无法创建映射类型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (this.cboMapType.Text == "")
            {
                MessageBox.Show("请选择一个要检索的映射类型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            this.dgvMapInfoOfCurType.Rows.Clear();

            SQLiteCommand sqCommand = Globals.Sheet1.connection.CreateCommand();
            sqCommand.CommandText = "select * from mapdefine where maptype='" + this.cboMapType.Text + "'";
            SQLiteDataReader sqReader =  sqCommand.ExecuteReader();
            while (sqReader.Read())
            {
                DataGridViewRow dr = new DataGridViewRow();
                dr.CreateCells(this.dgvMapInfoOfCurType);
                dr.Cells[0].Value = sqReader.GetValue(1).ToString();
                dr.Cells[1].Value = sqReader.GetValue(2).ToString();
                dr.Cells[2].Value = sqReader.GetValue(3).ToString();

                this.dgvMapInfoOfCurType.Rows.Add(dr);
            }           
        }
    }
}
