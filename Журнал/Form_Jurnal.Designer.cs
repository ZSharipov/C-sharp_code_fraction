namespace Журнал
{
    partial class Form_Jurnal
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_Jurnal));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.воридотToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.таъғиротToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.тозакунӣToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.шартномаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.оидиБарномаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.status_date = new System.Windows.Forms.ToolStripLabel();
            this.status_rows = new System.Windows.Forms.ToolStripLabel();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToOrderColumns = true;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.AliceBlue;
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.BackgroundColor = System.Drawing.Color.AliceBlue;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.GradientActiveCaption;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.AliceBlue;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.dataGridView1.Location = new System.Drawing.Point(0, 31);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersWidth = 10;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(995, 606);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.Color.Lavender;
            this.menuStrip1.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.menuStrip1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.воридотToolStripMenuItem,
            this.таъғиротToolStripMenuItem,
            this.тозакунӣToolStripMenuItem,
            this.шартномаToolStripMenuItem,
            this.оидиБарномаToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(8, 3, 0, 3);
            this.menuStrip1.Size = new System.Drawing.Size(995, 28);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // воридотToolStripMenuItem
            // 
            this.воридотToolStripMenuItem.Name = "воридотToolStripMenuItem";
            this.воридотToolStripMenuItem.Size = new System.Drawing.Size(76, 22);
            this.воридотToolStripMenuItem.Text = "Воридот";
            this.воридотToolStripMenuItem.Click += new System.EventHandler(this.воридотToolStripMenuItem_Click);
            // 
            // таъғиротToolStripMenuItem
            // 
            this.таъғиротToolStripMenuItem.Name = "таъғиротToolStripMenuItem";
            this.таъғиротToolStripMenuItem.Size = new System.Drawing.Size(82, 22);
            this.таъғиротToolStripMenuItem.Text = "Таъғирот";
            this.таъғиротToolStripMenuItem.Click += new System.EventHandler(this.таъғиротToolStripMenuItem_Click);
            // 
            // тозакунӣToolStripMenuItem
            // 
            this.тозакунӣToolStripMenuItem.Name = "тозакунӣToolStripMenuItem";
            this.тозакунӣToolStripMenuItem.Size = new System.Drawing.Size(81, 22);
            this.тозакунӣToolStripMenuItem.Text = "Тозакунӣ";
            this.тозакунӣToolStripMenuItem.Click += new System.EventHandler(this.тозакунӣToolStripMenuItem_Click);
            // 
            // шартномаToolStripMenuItem
            // 
            this.шартномаToolStripMenuItem.Name = "шартномаToolStripMenuItem";
            this.шартномаToolStripMenuItem.Size = new System.Drawing.Size(89, 22);
            this.шартномаToolStripMenuItem.Text = "Шартнома";
            this.шартномаToolStripMenuItem.Click += new System.EventHandler(this.шартномаToolStripMenuItem_Click);
            // 
            // оидиБарномаToolStripMenuItem
            // 
            this.оидиБарномаToolStripMenuItem.Name = "оидиБарномаToolStripMenuItem";
            this.оидиБарномаToolStripMenuItem.Size = new System.Drawing.Size(113, 22);
            this.оидиБарномаToolStripMenuItem.Text = "Оиди барнома";
            this.оидиБарномаToolStripMenuItem.Click += new System.EventHandler(this.оидиБарномаToolStripMenuItem_Click);
            // 
            // splitter1
            // 
            this.splitter1.BackColor = System.Drawing.Color.AliceBlue;
            this.splitter1.Dock = System.Windows.Forms.DockStyle.Top;
            this.splitter1.Location = new System.Drawing.Point(0, 28);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(995, 3);
            this.splitter1.TabIndex = 3;
            this.splitter1.TabStop = false;
            // 
            // toolStrip1
            // 
            this.toolStrip1.BackColor = System.Drawing.Color.Lavender;
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.toolStrip1.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.status_date,
            this.status_rows});
            this.toolStrip1.Location = new System.Drawing.Point(0, 637);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(995, 25);
            this.toolStrip1.TabIndex = 5;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // status_date
            // 
            this.status_date.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.status_date.Name = "status_date";
            this.status_date.Size = new System.Drawing.Size(33, 22);
            this.status_date.Text = "data";
            // 
            // status_rows
            // 
            this.status_rows.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.status_rows.Name = "status_rows";
            this.status_rows.Size = new System.Drawing.Size(41, 22);
            this.status_rows.Text = "0 сатр";
            // 
            // Form_Jurnal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.MistyRose;
            this.ClientSize = new System.Drawing.Size(995, 662);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.splitter1);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "Form_Jurnal";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Шартнома";
            this.Load += new System.EventHandler(this.Form_Jurnal_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripLabel status_date;
        private System.Windows.Forms.ToolStripLabel status_rows;
        public System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem воридотToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem таъғиротToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem тозакунӣToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem шартномаToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem оидиБарномаToolStripMenuItem;
    }
}

