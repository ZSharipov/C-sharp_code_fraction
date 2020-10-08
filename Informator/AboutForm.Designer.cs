namespace Informator
{
    partial class AboutForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutForm));
            this.verProg = new System.Windows.Forms.Label();
            this.lverProg = new System.Windows.Forms.Label();
            this.lprouctName = new System.Windows.Forms.Label();
            this.name = new System.Windows.Forms.Label();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lnote = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // verProg
            // 
            this.verProg.AutoSize = true;
            this.verProg.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.verProg.ForeColor = System.Drawing.SystemColors.ControlText;
            this.verProg.Location = new System.Drawing.Point(300, 57);
            this.verProg.Name = "verProg";
            this.verProg.Size = new System.Drawing.Size(26, 17);
            this.verProg.TabIndex = 10;
            this.verProg.Text = "sss";
            // 
            // lverProg
            // 
            this.lverProg.AutoSize = true;
            this.lverProg.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lverProg.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lverProg.Location = new System.Drawing.Point(185, 57);
            this.lverProg.Name = "lverProg";
            this.lverProg.Size = new System.Drawing.Size(110, 17);
            this.lverProg.TabIndex = 9;
            this.lverProg.Text = "Версияи барнома:";
            // 
            // lprouctName
            // 
            this.lprouctName.AutoSize = true;
            this.lprouctName.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lprouctName.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lprouctName.Location = new System.Drawing.Point(185, 23);
            this.lprouctName.Name = "lprouctName";
            this.lprouctName.Size = new System.Drawing.Size(95, 17);
            this.lprouctName.TabIndex = 8;
            this.lprouctName.Text = "Номи барнома:";
            // 
            // name
            // 
            this.name.AutoSize = true;
            this.name.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.name.ForeColor = System.Drawing.SystemColors.ControlText;
            this.name.Location = new System.Drawing.Point(299, 23);
            this.name.Name = "name";
            this.name.Size = new System.Drawing.Size(138, 17);
            this.name.TabIndex = 7;
            this.name.Text = "\"Иттилооти Шаҳрсозӣ\"";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "info.png");
            this.imageList1.Images.SetKeyName(1, "neotm copy.png");
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.ImageList = this.imageList1;
            this.button1.Location = new System.Drawing.Point(260, 185);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(87, 29);
            this.button1.TabIndex = 14;
            this.button1.Text = "OK";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lnote);
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.verProg);
            this.groupBox1.Controls.Add(this.lverProg);
            this.groupBox1.Controls.Add(this.lprouctName);
            this.groupBox1.Controls.Add(this.name);
            this.groupBox1.Location = new System.Drawing.Point(16, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(574, 171);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            // 
            // lnote
            // 
            this.lnote.AutoSize = true;
            this.lnote.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lnote.Location = new System.Drawing.Point(185, 100);
            this.lnote.Name = "lnote";
            this.lnote.Size = new System.Drawing.Size(340, 51);
            this.lnote.TabIndex = 3;
            this.lnote.Text = "Барномаи ахборотие, ки  ГОСТ, МҚС, СНиП, СП, РДС, СН,\r\nРСН, дастурҳо ,тавсияҳои м" +
                "етодӣ, фармоишҳо,қарорҳо,\r\nмактубҳо ва дигар ҳуҷҷатҳои меъёриро дар бар мегирад." +
                "";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Control;
            this.panel1.BackgroundImage = global::Informator.Properties.Resources.icon128;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Location = new System.Drawing.Point(15, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(133, 133);
            this.panel1.TabIndex = 5;
            // 
            // AboutForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(606, 228);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(614, 262);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(614, 262);
            this.Name = "AboutForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.AboutForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label verProg;
        private System.Windows.Forms.Label lverProg;
        private System.Windows.Forms.Label lprouctName;
        private System.Windows.Forms.Label name;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lnote;

    }
}