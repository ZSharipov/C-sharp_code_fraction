namespace Журнал
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
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.imageList2 = new System.Windows.Forms.ImageList(this.components);
            this.verProg = new System.Windows.Forms.Label();
            this.lverProg = new System.Windows.Forms.Label();
            this.lprouctName = new System.Windows.Forms.Label();
            this.name = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
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
            this.button1.ImageIndex = 1;
            this.button1.ImageList = this.imageList1;
            this.button1.Location = new System.Drawing.Point(217, 182);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(101, 31);
            this.button1.TabIndex = 6;
            this.button1.Text = "OK";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // imageList2
            // 
            this.imageList2.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList2.ImageStream")));
            this.imageList2.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList2.Images.SetKeyName(0, "info.png");
            this.imageList2.Images.SetKeyName(1, "neotm copy.png");
            // 
            // verProg
            // 
            this.verProg.AutoSize = true;
            this.verProg.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.verProg.ForeColor = System.Drawing.SystemColors.ControlText;
            this.verProg.Location = new System.Drawing.Point(350, 56);
            this.verProg.Name = "verProg";
            this.verProg.Size = new System.Drawing.Size(41, 17);
            this.verProg.TabIndex = 10;
            this.verProg.Text = "1.0.0.1";
            // 
            // lverProg
            // 
            this.lverProg.AutoSize = true;
            this.lverProg.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lverProg.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lverProg.Location = new System.Drawing.Point(216, 56);
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
            this.lprouctName.Location = new System.Drawing.Point(216, 30);
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
            this.name.Location = new System.Drawing.Point(349, 30);
            this.name.Name = "name";
            this.name.Size = new System.Drawing.Size(63, 17);
            this.name.TabIndex = 7;
            this.name.Text = "\"Журнал\"";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.verProg);
            this.groupBox1.Controls.Add(this.lverProg);
            this.groupBox1.Controls.Add(this.lprouctName);
            this.groupBox1.Controls.Add(this.name);
            this.groupBox1.Location = new System.Drawing.Point(4, 2);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox1.Size = new System.Drawing.Size(526, 172);
            this.groupBox1.TabIndex = 17;
            this.groupBox1.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(216, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(297, 51);
            this.label2.TabIndex = 11;
            this.label2.Text = "Барнома барои додани калид ба истифодабарони\r\nбарномаи \"Иттилооти Шаҳрсозӣ\"  ва б" +
                "а қайдгирии \r\nистифодабарон.";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Control;
            this.panel1.BackgroundImage = global::Журнал.Properties.Resources.icon128;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Location = new System.Drawing.Point(34, 24);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(133, 133);
            this.panel1.TabIndex = 12;
            // 
            // AboutForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(534, 226);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button1);
            this.Font = new System.Drawing.Font("Palatino Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(542, 260);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(542, 260);
            this.Name = "AboutForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ImageList imageList2;
        private System.Windows.Forms.Label verProg;
        private System.Windows.Forms.Label lverProg;
        private System.Windows.Forms.Label lprouctName;
        private System.Windows.Forms.Label name;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
    }
}