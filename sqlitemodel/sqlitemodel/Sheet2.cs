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
            SQLiteDataReader sqReader = sCommand.ExecuteReader();
            if (sqReader.HasRows)
            {
                if (MessageBox.Show("已经存在映射类型，是否覆盖?", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Error) == DialogResult.Cancel)
                {
                    return;
                }                
            }
            sqReader.Close();

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
            for (int i = 1; i < this.dgvMapInfoOfCurType.RowCount; ++i)
            {
                for (int j = i + 1; j < this.dgvMapInfoOfCurType.RowCount; ++j)
                {
                    if (this.dgvMapInfoOfCurType.Rows[i - 1].Cells[3].Value.ToString() == this.dgvMapInfoOfCurType.Rows[j - 1].Cells[3].Value.ToString())
                    {
                        MessageBox.Show("字段序号重复，无法保存", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            

            //  保存到数据库中
            //  开启一个事务
            System.Data.SQLite.SQLiteTransaction trans = Globals.Sheet1.connection.BeginTransaction();
            try
            {
                //  先 删除 原有的 数据 ，然后 保存 新的 
                sCommand.CommandText = "delete from mapdefine where maptype='" + this.cboMapType.Text + "'";
                sCommand.ExecuteNonQuery();

                for (int i = 1; i < this.dgvMapInfoOfCurType.RowCount; ++i)
                {
                    string sInsertSql = "insert into mapdefine values('" + this.cboMapType.Text + "'";
                    for (int j = 1; j <= this.dgvMapInfoOfCurType.ColumnCount; ++j)
                    {
                        if (j == 1)
                        {
                            sInsertSql += "," + this.dgvMapInfoOfCurType.Rows[i - 1].Cells[j - 1].Value.ToString();
                        }
                        else if (j == 4)
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
                MessageBox.Show("保存成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString(), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //  sheet1可用映射类型重新获取
            //DataGridViewRow dr = new DataGridViewRow();
            //dr.CreateCells(Globals.Sheet1.dgvAvailableMapType);
            //dr.Cells[0].Value = this.cboMapType.Text;

            //Globals.Sheet1.dgvAvailableMapType.Rows.Add(dr);

            FreshAvailableMapType();

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
                dr.Cells[3].Value = sqReader.GetInt32(4);


                this.dgvMapInfoOfCurType.Rows.Add(dr);
            }           
        }
        public void ClearContent()
        {
            this.cboMapType.Items.Clear();
            this.dgvMapInfoOfCurType.Rows.Clear();
        }
        private void FreshAvailableMapType()
        {
            this.cboMapType.Items.Clear();
            Globals.Sheet1.dgvAvailableMapType.Rows.Clear();

            SQLiteCommand sqliteCommand = Globals.Sheet1.connection.CreateCommand();

            sqliteCommand.CommandText = "select maptype from mapdefine group by maptype";
            SQLiteDataReader smapreader = sqliteCommand.ExecuteReader();
            if (smapreader.HasRows)
            {
                while (smapreader.Read())
                {
                    string sName = smapreader.GetString(0);

                    DataGridViewRow dr = new DataGridViewRow();
                    dr.CreateCells(Globals.Sheet1.dgvAvailableMapType);
                    dr.Cells[0].Value = sName;

                    Globals.Sheet1.dgvAvailableMapType.Rows.Add(dr);

                    this.cboMapType.Items.Add(sName);
                }
                //MessageBox.Show("可用映射已经装载成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
