namespace updk
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.mainDatagridView = new System.Windows.Forms.DataGridView();
            this.FIOColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GenderColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GroupColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TestNameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TestResultNameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TestResultValueColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.mainDatagridView)).BeginInit();
            this.SuspendLayout();
            // 
            // mainDatagridView
            // 
            this.mainDatagridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.mainDatagridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.mainDatagridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.FIOColumn,
            this.GenderColumn,
            this.GroupColumn,
            this.TestNameColumn,
            this.TestResultNameColumn,
            this.TestResultValueColumn});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.mainDatagridView.DefaultCellStyle = dataGridViewCellStyle1;
            this.mainDatagridView.Location = new System.Drawing.Point(12, 12);
            this.mainDatagridView.Name = "mainDatagridView";
            this.mainDatagridView.Size = new System.Drawing.Size(1245, 559);
            this.mainDatagridView.TabIndex = 0;
            this.mainDatagridView.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.mainDatagridView_CellFormatting);
            this.mainDatagridView.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.mainDatagridView_CellPainting);
            // 
            // FIOColumn
            // 
            this.FIOColumn.HeaderText = "ФИО";
            this.FIOColumn.Name = "FIOColumn";
            this.FIOColumn.Width = 300;
            // 
            // GenderColumn
            // 
            this.GenderColumn.HeaderText = "Пол";
            this.GenderColumn.Name = "GenderColumn";
            // 
            // GroupColumn
            // 
            this.GroupColumn.HeaderText = "Группа";
            this.GroupColumn.Name = "GroupColumn";
            // 
            // TestNameColumn
            // 
            this.TestNameColumn.HeaderText = "Тест";
            this.TestNameColumn.Name = "TestNameColumn";
            this.TestNameColumn.Width = 300;
            // 
            // TestResultNameColumn
            // 
            this.TestResultNameColumn.HeaderText = "Результат";
            this.TestResultNameColumn.Name = "TestResultNameColumn";
            this.TestResultNameColumn.Width = 300;
            // 
            // TestResultValueColumn
            // 
            this.TestResultValueColumn.HeaderText = "Значение";
            this.TestResultValueColumn.Name = "TestResultValueColumn";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1269, 575);
            this.Controls.Add(this.mainDatagridView);
            this.Name = "MainForm";
            this.Text = "MainForm";
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.mainDatagridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView mainDatagridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn FIOColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn GenderColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn GroupColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn TestNameColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn TestResultNameColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn TestResultValueColumn;
    }
}