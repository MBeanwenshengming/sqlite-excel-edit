﻿//------------------------------------------------------------------------------
// <auto-generated>
//     此代码由工具生成。
//     运行时版本:4.0.30319.1022
//
//     对此文件的更改可能会导致不正确的行为，并且如果
//     重新生成代码，这些更改将会丢失。
// </auto-generated>
//------------------------------------------------------------------------------

#pragma warning disable 414
namespace sqlitemodel {
    
    
    /// 
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(2)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class Sheet2 : Microsoft.Office.Tools.Excel.WorksheetBase {
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        private global::System.Object missing = global::System.Type.Missing;
        
        internal Microsoft.Office.Tools.Excel.Controls.Label label1;
        
        internal Microsoft.Office.Tools.Excel.Controls.ComboBox cboMapType;
        
        internal Microsoft.Office.Tools.Excel.Controls.DataGridView dgvMapInfoOfCurType;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnSaveToDB;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnGetFromDB;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btnBeginCreate;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button btndelmaptype;
        
        internal System.Windows.Forms.DataGridViewTextBoxColumn mapoldvalue;
        
        internal System.Windows.Forms.DataGridViewTextBoxColumn mapvalue;
        
        internal System.Windows.Forms.DataGridViewTextBoxColumn mapdesc;
        
        internal System.Windows.Forms.DataGridViewTextBoxColumn colRecordOrder;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        public Sheet2(global::Microsoft.Office.Tools.Excel.Factory factory, global::System.IServiceProvider serviceProvider) : 
                base(factory, serviceProvider, "Sheet2", "Sheet2") {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void Initialize() {
            base.Initialize();
            Globals.Sheet2 = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization() {
            this.InternalStartup();
            this.OnStartup();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void InitializeDataBindings() {
            this.BeginInitialization();
            this.BindToData();
            this.EndInitialization();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeCachedData() {
            if ((this.DataHost == null)) {
                return;
            }
            if (this.DataHost.IsCacheInitialized) {
                this.DataHost.FillCachedData(this);
            }
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BindToData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StartCaching(string MemberName) {
            this.DataHost.StartCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StopCaching(string MemberName) {
            this.DataHost.StopCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool IsCached(string MemberName) {
            return this.DataHost.IsCached(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BeginInitialization() {
            this.BeginInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void EndInitialization() {
            this.EndInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeControls() {
            this.label1 = new Microsoft.Office.Tools.Excel.Controls.Label(Globals.Factory, this.ItemProvider, this.HostContext, "1C54A06D51378F14F9919B86157E269FEABDD1", "1C54A06D51378F14F9919B86157E269FEABDD1", this, "label1");
            this.cboMapType = new Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, this.ItemProvider, this.HostContext, "25D353ED0275E2249FA2BCD32763C5A5D4CC62", "25D353ED0275E2249FA2BCD32763C5A5D4CC62", this, "cboMapType");
            this.dgvMapInfoOfCurType = new Microsoft.Office.Tools.Excel.Controls.DataGridView(Globals.Factory, this.ItemProvider, this.HostContext, "3476C46BF3B304340293BF9632168E88947C73", "3476C46BF3B304340293BF9632168E88947C73", this, "dgvMapInfoOfCurType");
            this.btnSaveToDB = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "410274FE6432E84491B49F7C45173A5C69D1B4", "410274FE6432E84491B49F7C45173A5C69D1B4", this, "btnSaveToDB");
            this.btnGetFromDB = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "58FAA37445CFBD545F55BEA45AACBC1A50ECF5", "58FAA37445CFBD545F55BEA45AACBC1A50ECF5", this, "btnGetFromDB");
            this.btnBeginCreate = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "665AC892960B7F64D566985B60FF984365BB96", "665AC892960B7F64D566985B60FF984365BB96", this, "btnBeginCreate");
            this.btndelmaptype = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "7CDBF2A1F78D7C7423C7AEE8769C3C4622A1D7", "7CDBF2A1F78D7C7423C7AEE8769C3C4622A1D7", this, "btndelmaptype");
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "10.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeComponents() {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.mapoldvalue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mapvalue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mapdesc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colRecordOrder = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMapInfoOfCurType)).BeginInit();
            // 
            // label1
            // 
            this.label1.Name = "label1";
            this.label1.Text = "映射类型";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cboMapType
            // 
            this.cboMapType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMapType.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboMapType.Name = "cboMapType";
            // 
            // dgvMapInfoOfCurType
            // 
            this.dgvMapInfoOfCurType.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dgvMapInfoOfCurType.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
                        this.mapoldvalue,
                        this.mapvalue,
                        this.mapdesc,
                        this.colRecordOrder});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvMapInfoOfCurType.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgvMapInfoOfCurType.Name = "dgvMapInfoOfCurType";
            this.dgvMapInfoOfCurType.ReadOnly = true;
            this.dgvMapInfoOfCurType.RowTemplate.Height = 23;
            this.dgvMapInfoOfCurType.Text = "dataGridView1";
            // 
            // mapoldvalue
            // 
            this.mapoldvalue.HeaderText = "映射原始值";
            this.mapoldvalue.Name = "mapoldvalue";
            this.mapoldvalue.ReadOnly = true;
            this.mapoldvalue.Width = 150;
            // 
            // mapvalue
            // 
            this.mapvalue.HeaderText = "映射值";
            this.mapvalue.Name = "mapvalue";
            this.mapvalue.ReadOnly = true;
            this.mapvalue.Width = 250;
            // 
            // mapdesc
            // 
            this.mapdesc.HeaderText = "映射说明";
            this.mapdesc.Name = "mapdesc";
            this.mapdesc.ReadOnly = true;
            this.mapdesc.Width = 300;
            // 
            // colRecordOrder
            // 
            this.colRecordOrder.HeaderText = "字段序号";
            this.colRecordOrder.Name = "colRecordOrder";
            this.colRecordOrder.ReadOnly = true;
            // 
            // btnSaveToDB
            // 
            this.btnSaveToDB.BackColor = System.Drawing.SystemColors.Control;
            this.btnSaveToDB.Enabled = false;
            this.btnSaveToDB.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnSaveToDB.Name = "btnSaveToDB";
            this.btnSaveToDB.Text = "保存当前值到数据库";
            this.btnSaveToDB.UseVisualStyleBackColor = false;
            // 
            // btnGetFromDB
            // 
            this.btnGetFromDB.BackColor = System.Drawing.SystemColors.Control;
            this.btnGetFromDB.Enabled = false;
            this.btnGetFromDB.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnGetFromDB.Name = "btnGetFromDB";
            this.btnGetFromDB.Text = "从数据库获取";
            this.btnGetFromDB.UseVisualStyleBackColor = false;
            // 
            // btnBeginCreate
            // 
            this.btnBeginCreate.BackColor = System.Drawing.SystemColors.Control;
            this.btnBeginCreate.Enabled = false;
            this.btnBeginCreate.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnBeginCreate.Name = "btnBeginCreate";
            this.btnBeginCreate.Text = "开始编辑";
            this.btnBeginCreate.UseVisualStyleBackColor = false;
            // 
            // btndelmaptype
            // 
            this.btndelmaptype.BackColor = System.Drawing.SystemColors.Control;
            this.btndelmaptype.ForeColor = System.Drawing.Color.Red;
            this.btndelmaptype.Name = "btndelmaptype";
            this.btndelmaptype.Text = "删除当前映射类型";
            this.btndelmaptype.UseVisualStyleBackColor = false;
            // 
            // Sheet2
            // 
            ((System.ComponentModel.ISupportInitialize)(this.dgvMapInfoOfCurType)).EndInit();
            this.label1.BindingContext = this.BindingContext;
            this.cboMapType.BindingContext = this.BindingContext;
            this.dgvMapInfoOfCurType.BindingContext = this.BindingContext;
            this.btnSaveToDB.BindingContext = this.BindingContext;
            this.btnGetFromDB.BindingContext = this.BindingContext;
            this.btnBeginCreate.BindingContext = this.BindingContext;
            this.btndelmaptype.BindingContext = this.BindingContext;
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName) {
            return this.DataHost.NeedsFill(this, MemberName);
        }
    }
    
    internal sealed partial class Globals {
        
        private static Sheet2 _Sheet2;
        
        internal static Sheet2 Sheet2 {
            get {
                return _Sheet2;
            }
            set {
                if ((_Sheet2 == null)) {
                    _Sheet2 = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
    }
}
