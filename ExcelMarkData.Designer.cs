namespace Excel_Mark_Data
{
    partial class ExcelMarkData
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
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.buttonStart = new System.Windows.Forms.Button();
            this.textBoxBrowse = new System.Windows.Forms.TextBox();
            this.dataGridViewresult = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.buttonInfo = new System.Windows.Forms.Button();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxLayout = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxName = new System.Windows.Forms.TextBox();
            this.btn_export = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewresult)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.BackColor = System.Drawing.Color.Yellow;
            this.buttonBrowse.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.buttonBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonBrowse.Location = new System.Drawing.Point(3, 3);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(157, 37);
            this.buttonBrowse.TabIndex = 1;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.UseVisualStyleBackColor = false;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // buttonStart
            // 
            this.buttonStart.BackColor = System.Drawing.Color.Red;
            this.buttonStart.Enabled = false;
            this.buttonStart.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.buttonStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonStart.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.buttonStart.Location = new System.Drawing.Point(931, 1);
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.Size = new System.Drawing.Size(181, 33);
            this.buttonStart.TabIndex = 1;
            this.buttonStart.Text = "Start";
            this.buttonStart.UseVisualStyleBackColor = false;
            this.buttonStart.Click += new System.EventHandler(this.buttonStart_Click);
            // 
            // textBoxBrowse
            // 
            this.textBoxBrowse.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxBrowse.Location = new System.Drawing.Point(618, 3);
            this.textBoxBrowse.Name = "textBoxBrowse";
            this.textBoxBrowse.ReadOnly = true;
            this.textBoxBrowse.Size = new System.Drawing.Size(1116, 22);
            this.textBoxBrowse.TabIndex = 2;
            this.textBoxBrowse.TextChanged += new System.EventHandler(this.textBoxBrowse_TextChanged);
            // 
            // dataGridViewresult
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            this.dataGridViewresult.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewresult.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader;
            this.dataGridViewresult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewresult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewresult.Location = new System.Drawing.Point(90, 452);
            this.dataGridViewresult.Name = "dataGridViewresult";
            this.dataGridViewresult.RowHeadersWidth = 51;
            this.dataGridViewresult.RowTemplate.Height = 24;
            this.dataGridViewresult.Size = new System.Drawing.Size(1743, 443);
            this.dataGridViewresult.TabIndex = 4;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Blue;
            this.button1.Enabled = false;
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(3, 0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(157, 33);
            this.button1.TabIndex = 5;
            this.button1.Text = "Read Excel";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 4.545455F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 90.90909F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 4.545455F));
            this.tableLayoutPanel1.Controls.Add(this.buttonInfo, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.dataGridViewresult, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1924, 919);
            this.tableLayoutPanel1.TabIndex = 6;
            // 
            // buttonInfo
            // 
            this.buttonInfo.Location = new System.Drawing.Point(1839, 3);
            this.buttonInfo.Name = "buttonInfo";
            this.buttonInfo.Size = new System.Drawing.Size(81, 33);
            this.buttonInfo.TabIndex = 10;
            this.buttonInfo.Text = "วิธีใช้";
            this.buttonInfo.UseVisualStyleBackColor = true;
            this.buttonInfo.Click += new System.EventHandler(this.buttonInfo_Click);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel3, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.panel3, 0, 1);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(90, 3);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 21.42857F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 78.57143F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(1743, 443);
            this.tableLayoutPanel2.TabIndex = 5;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 2;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35.46344F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 64.53656F));
            this.tableLayoutPanel3.Controls.Add(this.panel1, 0, 1);
            this.tableLayoutPanel3.Controls.Add(this.buttonBrowse, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.panel2, 1, 1);
            this.tableLayoutPanel3.Controls.Add(this.textBoxBrowse, 1, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 2;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 42F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(1737, 88);
            this.tableLayoutPanel3.TabIndex = 4;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 49);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(609, 36);
            this.panel1.TabIndex = 6;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.textBoxLayout);
            this.panel2.Controls.Add(this.buttonStart);
            this.panel2.Location = new System.Drawing.Point(618, 49);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1115, 36);
            this.panel2.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 17);
            this.label2.TabIndex = 9;
            this.label2.Text = "Layout Name";
            // 
            // textBoxLayout
            // 
            this.textBoxLayout.Location = new System.Drawing.Point(101, 5);
            this.textBoxLayout.Name = "textBoxLayout";
            this.textBoxLayout.Size = new System.Drawing.Size(824, 22);
            this.textBoxLayout.TabIndex = 2;
            this.textBoxLayout.Text = "MARK";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.panel4);
            this.panel3.Controls.Add(this.dataGridView);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(3, 97);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1737, 343);
            this.panel3.TabIndex = 5;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.label1);
            this.panel4.Controls.Add(this.textBoxName);
            this.panel4.Controls.Add(this.btn_export);
            this.panel4.Location = new System.Drawing.Point(617, 3);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1117, 70);
            this.panel4.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 17);
            this.label1.TabIndex = 9;
            this.label1.Text = "Save As";
            // 
            // textBoxName
            // 
            this.textBoxName.Location = new System.Drawing.Point(103, 8);
            this.textBoxName.Name = "textBoxName";
            this.textBoxName.Size = new System.Drawing.Size(824, 22);
            this.textBoxName.TabIndex = 8;
            this.textBoxName.Text = "DATA";
            // 
            // btn_export
            // 
            this.btn_export.BackColor = System.Drawing.Color.Lime;
            this.btn_export.Enabled = false;
            this.btn_export.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.btn_export.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_export.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_export.Location = new System.Drawing.Point(933, 3);
            this.btn_export.Name = "btn_export";
            this.btn_export.Size = new System.Drawing.Size(181, 33);
            this.btn_export.TabIndex = 7;
            this.btn_export.Text = "Export Excel";
            this.btn_export.UseVisualStyleBackColor = false;
            this.btn_export.Click += new System.EventHandler(this.btn_export_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(3, 3);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowHeadersWidth = 51;
            this.dataGridView.RowTemplate.Height = 24;
            this.dataGridView.Size = new System.Drawing.Size(608, 337);
            this.dataGridView.TabIndex = 4;
            // 
            // ExcelMarkData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1924, 919);
            this.Controls.Add(this.tableLayoutPanel1);
            this.DoubleBuffered = true;
            this.Name = "ExcelMarkData";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExcelMarkData";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewresult)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel3.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.Button buttonStart;
        private System.Windows.Forms.TextBox textBoxBrowse;
        private System.Windows.Forms.DataGridView dataGridViewresult;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.Button btn_export;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxName;
        private System.Windows.Forms.Button buttonInfo;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.TextBox textBoxLayout;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Label label2;
    }
}

