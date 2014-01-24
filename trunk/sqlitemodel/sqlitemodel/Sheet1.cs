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
using System.IO;
using sqlitemodel;

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
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            this.btnopendb.Click += new System.EventHandler(this.btnopendb_Click);
            this.dgvTableDefine.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.btnlistMapType.Click += new System.EventHandler(this.btnlistMapType_Click);
            this.btnCreateDB.Click += new System.EventHandler(this.btnCreateDB_Click);
            this.btnstartCreate.Click += new System.EventHandler(this.btnstartCreate_Click);
            this.btnModifyTable.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void btnCreate_Click(object sender, EventArgs e)
        {
            if (this.connection == null)
            {
                MessageBox.Show("数据库未打开", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (this.connection.State != ConnectionState.Open)
            {
                MessageBox.Show("数据库未打开", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (this.dgvTableDefine.RowCount == 1)
            {
                MessageBox.Show(@"您需要先设定字段信息，然后才可以创建", @"提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (this.txtNewTableName.Text == "")
            {
                MessageBox.Show("表名为空，无法创建新表", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }           
            if (this.txtTableDBName.Text == "")
            {
                MessageBox.Show("数据库表名为空，无法创建新表", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!isABC(this.txtTableDBName.Text))
            {
                MessageBox.Show("表的数据库名称必须为字母", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //  检查数据表是否存在
            if (CheckDBTableExist(this.txtNewTableName.Text, this.txtTableDBName.Text))
            {
                MessageBox.Show("表已经存在于数据库中，无法创建表，请重新命名", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //  检查数据合法性            
            for (int i = 1; i < this.dgvTableDefine.RowCount; ++i)
            {
                for (int j = 1; j <= this.dgvTableDefine.ColumnCount; ++j)
                {
                    if (j != 3 && j != 5 && j != 6)       //  字段说明和字段映射类型可以为空
                    {
                        if (this.dgvTableDefine.Rows[i - 1].Cells[j - 1].Value == null)
                        {
                            MessageBox.Show("所有的字段必须设置信息", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }
                    if (j == 2)
                    {
                        string svalue = (string)this.dgvTableDefine.Rows[i - 1].Cells[j - 1].Value;
                        if (!isABC(svalue))
                        {
                                MessageBox.Show("字段数据库名称必须为字母", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;                            
                        }
                    }
                    if (j == 5)
                    {
                        string svalue = (string)this.dgvTableDefine.Rows[i - 1].Cells[j - 1].Value;
                        
                        if (svalue != null && svalue != "")
                        {
                            SQLiteCommand sqliteCom = this.connection.CreateCommand();
                            sqliteCom.CommandText = "select * from mapdefine where maptype='" + svalue + "'";
                            SQLiteDataReader sqReader = sqliteCom.ExecuteReader();
                            if (!sqReader.HasRows)
                            {
                                MessageBox.Show("你选择的映射类型不存在，无法创建数据表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                        }
                    }
                    if (j == 7)
                    {
                        string sValue = (string)this.dgvTableDefine.Rows[i - 1].Cells[j - 1].Value;
                        if (!isNum(sValue))
                        {
                            MessageBox.Show("字段序号必须为数字，无法创建数据表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }
                }
            }
            //  检查字段合法性,不能存在重复字段
            for (int i = 1; i < this.dgvTableDefine.RowCount; ++i)
            {
                for (int j = i + 1; j < this.dgvTableDefine.RowCount; ++j)
                {
                    string svalue = (string)this.dgvTableDefine.Rows[i - 1].Cells[1].Value;
                    string svalue1 = (string)this.dgvTableDefine.Rows[j - 1].Cells[1].Value;
                    if (svalue == svalue1)
                    {
                        MessageBox.Show("存在重复字段，无法建表", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    svalue = (string)this.dgvTableDefine.Rows[i - 1].Cells[6].Value;
                    svalue1 = (string)this.dgvTableDefine.Rows[j - 1].Cells[6].Value;
                    if (svalue == svalue1)
                    {
                        MessageBox.Show("存在重复字段序号，无法建表", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }
            }

            string[] sInsertIntoTabledefine = new string[this.dgvTableDefine.RowCount - 1];
            string sCreateTableInDB         = "";
            string[] sCreateIndexOnTable = new string[this.dgvTableDefine.RowCount - 1];
            //  插入一条记录到tabledefine
            
            for (int i = 1; i < this.dgvTableDefine.RowCount; ++i)
            {
                    sInsertIntoTabledefine[i - 1] = "insert into tabledefine (dbtablename, tablename, fieldname, dbfieldname, fielddesc, fieldtype,dbmaptype, iscreateindex,RecordOrder) values('" + this.txtTableDBName.Text + "','" + this.txtNewTableName.Text + "'";
                   
                    for (int j = 1; j <= this.dgvTableDefine.ColumnCount; ++j)
                    {
                        if (j == 6)
                        {
                            if (this.dgvTableDefine.Rows[i - 1].Cells[j - 1].Value == null)
                            {
                                sInsertIntoTabledefine[i - 1] += ",0";
                            }
                            else
                            {
                                bool bValue = (bool)this.dgvTableDefine.Rows[i - 1].Cells[j - 1].Value;
                                if (bValue)
                                {
                                    sInsertIntoTabledefine[i - 1] += ",1";
                                }
                                else
                                {
                                    sInsertIntoTabledefine[i - 1] += ",0";
                                }
                            }
                        }
                        else if (j == 7)
                        {
                            string sValue = (string)this.dgvTableDefine.Rows[i - 1].Cells[j - 1].Value;                            
                            sInsertIntoTabledefine[i - 1] += "," + sValue;
                        }
                        else
                        {
                            string sValue = (string)this.dgvTableDefine.Rows[i - 1].Cells[j - 1].Value;
                            if (sValue == null)
                            {
                                sInsertIntoTabledefine[i - 1] += ",''";
                            }
                            else
                            {
                                sInsertIntoTabledefine[i - 1] += ",'" + (string)this.dgvTableDefine.Rows[i - 1].Cells[j - 1].Value + "'";
                            }
                        }
                    }
                    sInsertIntoTabledefine[i - 1] += ")";                    
             }

            sCreateTableInDB = "create table " + this.txtTableDBName.Text + "(RecordOrder int PRIMARY KEY";
            for (int i = 1; i < this.dgvTableDefine.Rows.Count; ++i)
            {
                //if (i == 1)
                //{
                //    sCreateTableInDB += (string)this.dgvTableDefine.Rows[i - 1].Cells[1].Value + " " + (string)this.dgvTableDefine.Rows[i - 1].Cells[3].Value + " not null";
                //}
                //else
                //{
                    sCreateTableInDB += "," + (string)this.dgvTableDefine.Rows[i - 1].Cells[1].Value + " " + (string)this.dgvTableDefine.Rows[i - 1].Cells[3].Value + " not null";
                //}
                if (this.dgvTableDefine.Rows[i - 1].Cells[5].Value != null )     //  需要创建索引，那么创建
                {
                    if ((bool)this.dgvTableDefine.Rows[i - 1].Cells[5].Value == true)
                    {
                        sCreateIndexOnTable[i - 1] = "create index " + this.txtTableDBName.Text + "_" + (string)this.dgvTableDefine.Rows[i - 1].Cells[1].Value + " on " + this.txtTableDBName.Text + "(" + (string)this.dgvTableDefine.Rows[i - 1].Cells[1].Value + ")";
                    }
                }
            }            
            sCreateTableInDB += ")";


            //  用事务提交所有更改
            SQLiteTransaction sqTrans = this.connection.BeginTransaction();
            try
            {
                SQLiteCommand sqCommand = this.connection.CreateCommand();
                for (int i = 0 ; i < sInsertIntoTabledefine.Count(); ++i)
                {
                    if (sInsertIntoTabledefine[i] != null &&  sInsertIntoTabledefine[i] != "")
                    {
                        sqCommand.CommandText = sInsertIntoTabledefine[i];
                        sqCommand.ExecuteNonQuery();
                    }
                }
                sqCommand.CommandText = sCreateTableInDB;
                sqCommand.ExecuteNonQuery();

                for (int i = 0; i < sCreateIndexOnTable.Count(); ++i)
                {
                    if (sCreateIndexOnTable[i] != null && sCreateIndexOnTable[i] != "")
                    {
                        sqCommand.CommandText = sCreateIndexOnTable[i];
                        sqCommand.ExecuteNonQuery();
                    }
                }
                sqTrans.Commit();
            }
            catch (Exception E)
            {                
                sqTrans.Rollback();
            }
            
            //  在数据库创建该表
            this.txtNewTableName.Enabled = false;
            this.txtTableDBName.Enabled = false;
            this.dgvTableDefine.ReadOnly = true;
            this.btnCreate.Enabled = false;
            //this.cboTable.DropDownStyle = ComboBoxStyle.DropDownList;
        }
        private bool isABC(string sValue)
        {
            for (int m = 0; m < sValue.Length; ++m)
            {
                if (sValue[m] < 'a' || sValue[m] > 'z')
                {
                    return false;
                }
            }
            return true;
        }
        private bool isNum(string sValue)
        {
            for (int m = 0; m < sValue.Length; ++m)
            {
                if (sValue[m] < '1' || sValue[m] > '7')
                {
                    return false;
                }
            }
            return true;
        }
        private bool CheckDBTableExist(string sTableName, string sDBTableName)
        {
            return false;
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnopendb_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择要打开的数据库文件";
            openFileDialog.InitialDirectory = "c://";
            openFileDialog.Filter = "所有文件|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fName = openFileDialog.FileName;
                

                try
                {
                    //  打开数据库
                    string sDataSource = "data source=" + fName;
                    connection = new SQLiteConnection(sDataSource);
                    connection.Open();

                    this.db.Text = fName;
                    MessageBox.Show("数据库打开成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //  读取已经有的数据库定义表
                    SQLiteCommand sqliteCommand = connection.CreateCommand();
                    sqliteCommand.CommandText = "select tablename from tabledefine group by tablename";
                    
                    SQLiteDataReader s = sqliteCommand.ExecuteReader();
                    int nFieldCount = s.FieldCount;
                    //s.
                    if (s.HasRows)
                    {
                        while (s.Read())
                        {
                            string sName = s.GetString(0);                            
                            this.cboTable.Items.Add(sName);
                        }
                        MessageBox.Show("已存在的表名已经装载成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    s.Close();

                    sqliteCommand.CommandText = "select maptype from mapdefine group by maptype";
                    SQLiteDataReader smapreader = sqliteCommand.ExecuteReader();
                    if (smapreader.HasRows)
                    {                        
                        while (smapreader.Read())
                        {
                            string sName = smapreader.GetString(0);

                            DataGridViewRow dr = new DataGridViewRow();
                            dr.CreateCells(dgvAvailableMapType);
                            dr.Cells[0].Value = sName;

                            this.dgvAvailableMapType.Rows.Add(dr);
                        }
                        MessageBox.Show("可用映射已经装载成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception excep)
                {
                    MessageBox.Show("数据库打开出错", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            } 
        }


        private void btnCreateDB_Click(object sender, EventArgs e)
        {
            DialogResult dlgResult = MessageBox.Show(@"是否创建一个新的数据库！", @"创建！", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (dlgResult == DialogResult.Cancel)
            {
                return;
            }
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择一个目录，并且输入数据库文件名";
            openFileDialog.InitialDirectory = "c://";
            openFileDialog.Filter = "所有文件|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            openFileDialog.CheckFileExists = false;
            openFileDialog.CheckPathExists = true;
            openFileDialog.FileName = "InputDBName";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fName = openFileDialog.FileName;
                this.db.Text = fName;

                if (File.Exists(fName))
                {
                    MessageBox.Show("文件已经存在无法创建数据库！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                try
                {
                    //  创建一个数据库
                    System.Data.SQLite.SQLiteConnection.CreateFile(fName);

                    //  创建数据表定义表格
                    string sConString = "data source=" + fName;
                    SQLiteConnection sqlCon = new SQLiteConnection(sConString);
                    sqlCon.Open();                   
                    
                    SQLiteCommand sqliteConmand = sqlCon.CreateCommand();
                    sqliteConmand.CommandText = "create table tabledefine(dbtablename varchar(30) not null, tablename varchar(30) not null, fieldname varchar(100), fielddesc varchar(100), fieldtype varchar(30), dbfieldname varchar(30)  not null, dbmaptype varchar(50), iscreateindex int default(0),RecordOrder int not null)";
                    int nResult = sqliteConmand.ExecuteNonQuery();

                    sqliteConmand.CommandText = "create index tabledefine_dbtablename_tablename_fieldname on tabledefine(dbtablename,tablename,fieldname)";
                    sqliteConmand.ExecuteNonQuery();

                    sqliteConmand.CommandText = "create table mapdefine(maptype varchar(50) not null, mapoldvalue int not null, mapvalue varchar(100) not null default(''), mapdesc varchar(100),RecordOrder int not null)";
                    nResult = sqliteConmand.ExecuteNonQuery();

                    sqliteConmand.CommandText = "create index mapdefine_maptype on mapdefine(maptype)";
                    sqliteConmand.ExecuteNonQuery();

                    sqlCon.Close();

                    MessageBox.Show("数据库创建成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);                    
                }
                catch (Exception excep)
                {
                    MessageBox.Show("创建数据库出错", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            } 
        }

        private void btnstartCreate_Click(object sender, EventArgs e)
        {
            this.txtNewTableName.Enabled = true;
            this.txtTableDBName.Enabled = true;
            this.dgvTableDefine.ReadOnly = false;
            this.btnCreate.Enabled = true;
            this.dgvTableDefine.Rows.Clear();
            //this.cboTable.DropDownStyle = ComboBoxStyle.DropDown;
        }

        private void btnlistMapType_Click(object sender, EventArgs e)
        {
            if (this.connection == null)
            {
                MessageBox.Show("当前的数据库不处于打开状态，无法获取映射类型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (this.connection.State != ConnectionState.Open)
            {
                MessageBox.Show("当前的数据库不处于打开状态，无法获取映射类型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (this.txtMapType.Text == "")
            {
                MessageBox.Show("请输入映射名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            SQLiteCommand sCommand = this.connection.CreateCommand();
            sCommand.CommandText = "select mapoldvalue, mapvalue, mapdesc, RecordOrder from mapdefine where maptype='" + this.txtMapType.Text + "'";
            
            SQLiteDataReader reader = sCommand.ExecuteReader();


            while (reader.Read())
            {
                DataGridViewRow dr = new DataGridViewRow();
                dr.CreateCells(this.dgvMapTypeInfo);
                dr.Cells[0].Value = reader.GetValue(0).ToString();
                dr.Cells[1].Value = reader.GetValue(1).ToString();
                dr.Cells[2].Value = reader.GetValue(2).ToString();
                dr.Cells[3].Value = reader.GetValue(3).ToString();
                this.dgvMapTypeInfo.Rows.Add(dr);
            }
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (this.connection == null)
            {
                MessageBox.Show("数据库未打开", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (this.connection.State != ConnectionState.Open)
            {
                MessageBox.Show("数据库未打开", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (this.cboTable.Text == "")
            {
                MessageBox.Show("请选择一个数据表", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            this.dgvTableDefine.Rows.Clear();
            SQLiteCommand sqCommand = this.connection.CreateCommand();
            sqCommand.CommandText = "select * from tabledefine where tablename='" + this.cboTable.Text + "'";
            SQLiteDataReader sqReader =  sqCommand.ExecuteReader();

            while (sqReader.Read())
            {
                this.txtNewTableName.Text = sqReader.GetValue(1).ToString();
                this.txtTableDBName.Text = sqReader.GetValue(0).ToString();

                DataGridViewRow dr = new DataGridViewRow();
                dr.CreateCells(this.dgvTableDefine);
                dr.Cells[0].Value = sqReader.GetValue(2).ToString();
                dr.Cells[1].Value = sqReader.GetValue(5).ToString();
                dr.Cells[2].Value = sqReader.GetValue(3).ToString();
                dr.Cells[3].Value = sqReader.GetValue(4).ToString();
                dr.Cells[4].Value = sqReader.GetValue(6).ToString();                    
                dr.Cells[5].Value = sqReader.GetInt32(7) == 1 ? true : false;
                dr.Cells[6].Value = sqReader.GetInt32(8);
                if (dr.Cells[4].Value != "")
                {
                    Globals.Sheet2.cboMapType.Items.Add(dr.Cells[4].Value);
                }
                this.dgvTableDefine.Rows.Add(dr);
            }
            btnModifyTable.Enabled = true;

            Globals.Sheet3.BindData(this.txtTableDBName.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.dgvTableDefine.ReadOnly = false;
        }  
    }
}
