namespace Informator
{
    partial class AddFileForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddFileForm));
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.maskedTextBox2 = new System.Windows.Forms.MaskedTextBox();
            this.maskedTextBox1 = new System.Windows.Forms.MaskedTextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label18 = new System.Windows.Forms.Label();
            this.richTextBox15 = new System.Windows.Forms.RichTextBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.label16 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.richTextBox12 = new System.Windows.Forms.RichTextBox();
            this.richTextBox11 = new System.Windows.Forms.RichTextBox();
            this.richTextBox10 = new System.Windows.Forms.RichTextBox();
            this.richTextBox9 = new System.Windows.Forms.RichTextBox();
            this.richTextBox8 = new System.Windows.Forms.RichTextBox();
            this.richTextBox7 = new System.Windows.Forms.RichTextBox();
            this.richTextBox6 = new System.Windows.Forms.RichTextBox();
            this.richTextBox3 = new System.Windows.Forms.RichTextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "(*.doc)|*.doc|(*.docx)|*.docx";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.maskedTextBox2);
            this.panel1.Controls.Add(this.maskedTextBox1);
            this.panel1.Controls.Add(this.label12);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.richTextBox15);
            this.panel1.Controls.Add(this.richTextBox1);
            this.panel1.Controls.Add(this.label13);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.label16);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.richTextBox12);
            this.panel1.Controls.Add(this.richTextBox11);
            this.panel1.Controls.Add(this.richTextBox10);
            this.panel1.Controls.Add(this.richTextBox9);
            this.panel1.Controls.Add(this.richTextBox8);
            this.panel1.Controls.Add(this.richTextBox7);
            this.panel1.Controls.Add(this.richTextBox6);
            this.panel1.Controls.Add(this.richTextBox3);
            this.panel1.Controls.Add(this.label15);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Location = new System.Drawing.Point(12, 13);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(854, 468);
            this.panel1.TabIndex = 0;
            // 
            // maskedTextBox2
            // 
            this.maskedTextBox2.BeepOnError = true;
            this.maskedTextBox2.Location = new System.Drawing.Point(709, 55);
            this.maskedTextBox2.Mask = "00/00/0000";
            this.maskedTextBox2.Name = "maskedTextBox2";
            this.maskedTextBox2.Size = new System.Drawing.Size(110, 20);
            this.maskedTextBox2.TabIndex = 3;
            this.maskedTextBox2.ValidatingType = typeof(System.DateTime);
            this.maskedTextBox2.TypeValidationCompleted += new System.Windows.Forms.TypeValidationEventHandler(this.maskedTextBox2_TypeValidationCompleted);
            // 
            // maskedTextBox1
            // 
            this.maskedTextBox1.BeepOnError = true;
            this.maskedTextBox1.Location = new System.Drawing.Point(443, 55);
            this.maskedTextBox1.Mask = "00/00/0000";
            this.maskedTextBox1.Name = "maskedTextBox1";
            this.maskedTextBox1.Size = new System.Drawing.Size(111, 20);
            this.maskedTextBox1.TabIndex = 2;
            this.maskedTextBox1.ValidatingType = typeof(System.DateTime);
            this.maskedTextBox1.TypeValidationCompleted += new System.Windows.Forms.TypeValidationEventHandler(this.maskedTextBox1_TypeValidationCompleted);
            // 
            // label12
            // 
            this.label12.Location = new System.Drawing.Point(-11, 75);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(10, 17);
            this.label12.TabIndex = 146;
            this.label12.Text = "Выбор файла:";
            this.label12.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Palatino Linotype", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button3.ImageIndex = 9;
            this.button3.ImageList = this.imageList1;
            this.button3.Location = new System.Drawing.Point(441, 418);
            this.button3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(110, 30);
            this.button3.TabIndex = 15;
            this.button3.Text = "Отмена";
            this.button3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "2.ico");
            this.imageList1.Images.SetKeyName(1, "Folder.ico");
            this.imageList1.Images.SetKeyName(2, "navigator.ico");
            this.imageList1.Images.SetKeyName(3, "text.ico");
            this.imageList1.Images.SetKeyName(4, "openfol.png");
            this.imageList1.Images.SetKeyName(5, "fol.png");
            this.imageList1.Images.SetKeyName(6, "log.bmp");
            this.imageList1.Images.SetKeyName(7, "hgj.ico");
            this.imageList1.Images.SetKeyName(8, "info.png");
            this.imageList1.Images.SetKeyName(9, "otm.png");
            this.imageList1.Images.SetKeyName(10, "neotm copy.png");
            this.imageList1.Images.SetKeyName(11, "poisk.ico");
            // 
            // comboBox1
            // 
            this.comboBox1.DisplayMember = "1";
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "действует",
            "отменен"});
            this.comboBox1.Location = new System.Drawing.Point(111, 93);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(166, 21);
            this.comboBox1.TabIndex = 4;
            this.comboBox1.ValueMember = "1";
            // 
            // label18
            // 
            this.label18.Location = new System.Drawing.Point(47, 58);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(58, 17);
            this.label18.TabIndex = 25;
            this.label18.Text = "Номер:";
            this.label18.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // richTextBox15
            // 
            this.richTextBox15.Location = new System.Drawing.Point(111, 55);
            this.richTextBox15.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox15.Name = "richTextBox15";
            this.richTextBox15.Size = new System.Drawing.Size(166, 21);
            this.richTextBox15.TabIndex = 1;
            this.richTextBox15.Tag = "2";
            this.richTextBox15.Text = "";
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(400, 92);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.richTextBox1.Size = new System.Drawing.Size(154, 21);
            this.richTextBox1.TabIndex = 5;
            this.richTextBox1.Tag = "3";
            this.richTextBox1.Text = "";
            // 
            // label13
            // 
            this.label13.Location = new System.Drawing.Point(334, 95);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(48, 17);
            this.label13.TabIndex = 22;
            this.label13.Text = "Тип:";
            this.label13.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(-16, 96);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 17);
            this.label1.TabIndex = 24;
            this.label1.Text = "Статус документа:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(783, 21);
            this.button2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(38, 21);
            this.button2.TabIndex = 1;
            this.button2.Text = "...";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.ForeColor = System.Drawing.Color.Red;
            this.label16.Location = new System.Drawing.Point(100, 22);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(11, 13);
            this.label16.TabIndex = 19;
            this.label16.Text = "*";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(111, 22);
            this.textBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(673, 20);
            this.textBox1.TabIndex = 0;
            this.textBox1.Tag = "4";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(-9, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(113, 17);
            this.label3.TabIndex = 4;
            this.label3.Text = "Выбор файла:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Palatino Linotype", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.ImageIndex = 10;
            this.button1.ImageList = this.imageList1;
            this.button1.Location = new System.Drawing.Point(300, 418);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(110, 30);
            this.button1.TabIndex = 14;
            this.button1.Text = "OK";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // richTextBox12
            // 
            this.richTextBox12.Location = new System.Drawing.Point(110, 373);
            this.richTextBox12.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox12.Name = "richTextBox12";
            this.richTextBox12.Size = new System.Drawing.Size(711, 25);
            this.richTextBox12.TabIndex = 13;
            this.richTextBox12.Tag = "15";
            this.richTextBox12.Text = "";
            // 
            // richTextBox11
            // 
            this.richTextBox11.Location = new System.Drawing.Point(110, 339);
            this.richTextBox11.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox11.Name = "richTextBox11";
            this.richTextBox11.Size = new System.Drawing.Size(711, 25);
            this.richTextBox11.TabIndex = 12;
            this.richTextBox11.Tag = "14";
            this.richTextBox11.Text = "";
            // 
            // richTextBox10
            // 
            this.richTextBox10.Location = new System.Drawing.Point(111, 305);
            this.richTextBox10.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox10.Name = "richTextBox10";
            this.richTextBox10.Size = new System.Drawing.Size(710, 25);
            this.richTextBox10.TabIndex = 11;
            this.richTextBox10.Tag = "13";
            this.richTextBox10.Text = "";
            // 
            // richTextBox9
            // 
            this.richTextBox9.Location = new System.Drawing.Point(110, 271);
            this.richTextBox9.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox9.Name = "richTextBox9";
            this.richTextBox9.Size = new System.Drawing.Size(711, 25);
            this.richTextBox9.TabIndex = 10;
            this.richTextBox9.Tag = "12";
            this.richTextBox9.Text = "";
            // 
            // richTextBox8
            // 
            this.richTextBox8.Location = new System.Drawing.Point(110, 237);
            this.richTextBox8.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox8.Name = "richTextBox8";
            this.richTextBox8.Size = new System.Drawing.Size(711, 25);
            this.richTextBox8.TabIndex = 9;
            this.richTextBox8.Tag = "11";
            this.richTextBox8.Text = "";
            // 
            // richTextBox7
            // 
            this.richTextBox7.Location = new System.Drawing.Point(110, 203);
            this.richTextBox7.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox7.Name = "richTextBox7";
            this.richTextBox7.Size = new System.Drawing.Size(711, 25);
            this.richTextBox7.TabIndex = 8;
            this.richTextBox7.Tag = "10";
            this.richTextBox7.Text = "";
            // 
            // richTextBox6
            // 
            this.richTextBox6.Location = new System.Drawing.Point(110, 169);
            this.richTextBox6.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox6.Name = "richTextBox6";
            this.richTextBox6.Size = new System.Drawing.Size(711, 25);
            this.richTextBox6.TabIndex = 7;
            this.richTextBox6.Tag = "9";
            this.richTextBox6.Text = "";
            // 
            // richTextBox3
            // 
            this.richTextBox3.Location = new System.Drawing.Point(111, 133);
            this.richTextBox3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.richTextBox3.Name = "richTextBox3";
            this.richTextBox3.Size = new System.Drawing.Size(710, 25);
            this.richTextBox3.TabIndex = 6;
            this.richTextBox3.Tag = "8";
            this.richTextBox3.Text = "";
            // 
            // label15
            // 
            this.label15.Location = new System.Drawing.Point(-39, 138);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(143, 17);
            this.label15.TabIndex = 29;
            this.label15.Text = "Утвержден:";
            this.label15.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label14
            // 
            this.label14.Location = new System.Drawing.Point(581, 58);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(122, 17);
            this.label14.TabIndex = 23;
            this.label14.Text = "Дата утверждения:";
            this.label14.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(-43, 377);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(146, 17);
            this.label10.TabIndex = 36;
            this.label10.Text = "Содержание:";
            this.label10.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(-46, 343);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(149, 17);
            this.label9.TabIndex = 35;
            this.label9.Text = "Комментарий:";
            this.label9.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(-46, 309);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(150, 17);
            this.label8.TabIndex = 34;
            this.label8.Text = "Область действия:";
            this.label8.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(-39, 275);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(142, 17);
            this.label7.TabIndex = 33;
            this.label7.Text = "Изменения:";
            this.label7.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(-39, 241);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(142, 17);
            this.label6.TabIndex = 32;
            this.label6.Text = "Взамен:";
            this.label6.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(-46, 207);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(149, 17);
            this.label5.TabIndex = 31;
            this.label5.Text = "Опубликован:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(-36, 173);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(139, 17);
            this.label4.TabIndex = 30;
            this.label4.Text = "Разработчики:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(313, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(122, 17);
            this.label2.TabIndex = 28;
            this.label2.Text = "Начало действия:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // AddFileForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkTurquoise;
            this.ClientSize = new System.Drawing.Size(878, 502);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(886, 536);
            this.MinimumSize = new System.Drawing.Size(886, 536);
            this.Name = "AddFileForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = " Добавление";
            this.Load += new System.EventHandler(this.AddFileForm_Load);
            this.Shown += new System.EventHandler(this.AddFileForm_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.AddFileForm_KeyDown);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.RichTextBox richTextBox15;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.RichTextBox richTextBox12;
        private System.Windows.Forms.RichTextBox richTextBox11;
        private System.Windows.Forms.RichTextBox richTextBox10;
        private System.Windows.Forms.RichTextBox richTextBox9;
        private System.Windows.Forms.RichTextBox richTextBox8;
        private System.Windows.Forms.RichTextBox richTextBox7;
        private System.Windows.Forms.RichTextBox richTextBox6;
        private System.Windows.Forms.RichTextBox richTextBox3;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.MaskedTextBox maskedTextBox1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.MaskedTextBox maskedTextBox2;
    }
}