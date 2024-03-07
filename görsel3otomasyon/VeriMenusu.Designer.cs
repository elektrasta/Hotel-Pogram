
namespace görsel3otomasyon
{
    partial class VeriMenusu
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
            this.txtKontrol = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.inputDizini = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonImp = new System.Windows.Forms.Button();
            this.btngozat = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnExpGozat = new System.Windows.Forms.Button();
            this.tbExportPath = new System.Windows.Forms.TextBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnVTYDGozat = new System.Windows.Forms.Button();
            this.tbVTE = new System.Windows.Forms.TextBox();
            this.btnVTYedektenDon = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.tbVTI = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnVTYedekle = new System.Windows.Forms.Button();
            this.btnVTYGozat = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtKontrol
            // 
            this.txtKontrol.AutoSize = true;
            this.txtKontrol.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtKontrol.Location = new System.Drawing.Point(511, 282);
            this.txtKontrol.Name = "txtKontrol";
            this.txtKontrol.Size = new System.Drawing.Size(263, 32);
            this.txtKontrol.TabIndex = 2;
            this.txtKontrol.Text = "işlem bekleniyor...";
            // 
            // groupBox1
            // 
            this.groupBox1.AutoSize = true;
            this.groupBox1.BackColor = System.Drawing.Color.Azure;
            this.groupBox1.Controls.Add(this.inputDizini);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.buttonImp);
            this.groupBox1.Controls.Add(this.btngozat);
            this.groupBox1.Location = new System.Drawing.Point(12, 133);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(536, 113);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            // 
            // inputDizini
            // 
            this.inputDizini.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.inputDizini.Location = new System.Drawing.Point(200, 18);
            this.inputDizini.Name = "inputDizini";
            this.inputDizini.ReadOnly = true;
            this.inputDizini.Size = new System.Drawing.Size(233, 30);
            this.inputDizini.TabIndex = 21;
            this.inputDizini.TextChanged += new System.EventHandler(this.inputDizini_TextChanged_1);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label2.Location = new System.Drawing.Point(6, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(154, 25);
            this.label2.TabIndex = 18;
            this.label2.Text = "Öğrenci Import";
            // 
            // buttonImp
            // 
            this.buttonImp.BackColor = System.Drawing.Color.WhiteSmoke;
            this.buttonImp.Enabled = false;
            this.buttonImp.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.buttonImp.Location = new System.Drawing.Point(450, 49);
            this.buttonImp.Name = "buttonImp";
            this.buttonImp.Size = new System.Drawing.Size(80, 30);
            this.buttonImp.TabIndex = 20;
            this.buttonImp.Text = "IMPORT";
            this.buttonImp.UseVisualStyleBackColor = false;
            this.buttonImp.Click += new System.EventHandler(this.buttonImp_Click);
            // 
            // btngozat
            // 
            this.btngozat.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btngozat.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btngozat.Location = new System.Drawing.Point(460, 13);
            this.btngozat.Name = "btngozat";
            this.btngozat.Size = new System.Drawing.Size(70, 30);
            this.btngozat.TabIndex = 19;
            this.btngozat.Text = "Gözat...";
            this.btngozat.UseVisualStyleBackColor = false;
            this.btngozat.Click += new System.EventHandler(this.btngozat_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label4.Location = new System.Drawing.Point(298, 282);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(207, 32);
            this.label4.TabIndex = 8;
            this.label4.Text = "İşlem Sonucu:";
            // 
            // groupBox2
            // 
            this.groupBox2.AutoSize = true;
            this.groupBox2.BackColor = System.Drawing.Color.Azure;
            this.groupBox2.Controls.Add(this.btnExpGozat);
            this.groupBox2.Controls.Add(this.tbExportPath);
            this.groupBox2.Controls.Add(this.btnExport);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(12, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(536, 102);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            // 
            // btnExpGozat
            // 
            this.btnExpGozat.AutoSize = true;
            this.btnExpGozat.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnExpGozat.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnExpGozat.Location = new System.Drawing.Point(460, 13);
            this.btnExpGozat.Name = "btnExpGozat";
            this.btnExpGozat.Size = new System.Drawing.Size(70, 30);
            this.btnExpGozat.TabIndex = 9;
            this.btnExpGozat.Text = "Gözat...";
            this.btnExpGozat.UseVisualStyleBackColor = false;
            this.btnExpGozat.Click += new System.EventHandler(this.btnExpGozat_Click);
            // 
            // tbExportPath
            // 
            this.tbExportPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.tbExportPath.Location = new System.Drawing.Point(200, 15);
            this.tbExportPath.Name = "tbExportPath";
            this.tbExportPath.ReadOnly = true;
            this.tbExportPath.Size = new System.Drawing.Size(231, 30);
            this.tbExportPath.TabIndex = 11;
            this.tbExportPath.TextChanged += new System.EventHandler(this.tbExportPath_TextChanged);
            // 
            // btnExport
            // 
            this.btnExport.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnExport.Enabled = false;
            this.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnExport.Location = new System.Drawing.Point(450, 49);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(80, 30);
            this.btnExport.TabIndex = 9;
            this.btnExport.Text = "EXPORT";
            this.btnExport.UseVisualStyleBackColor = false;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(6, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Öğrenci Export";
            // 
            // groupBox3
            // 
            this.groupBox3.AutoSize = true;
            this.groupBox3.BackColor = System.Drawing.Color.Azure;
            this.groupBox3.Controls.Add(this.btnVTYDGozat);
            this.groupBox3.Controls.Add(this.tbVTE);
            this.groupBox3.Controls.Add(this.btnVTYedektenDon);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Location = new System.Drawing.Point(610, 120);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(444, 126);
            this.groupBox3.TabIndex = 13;
            this.groupBox3.TabStop = false;
            // 
            // btnVTYDGozat
            // 
            this.btnVTYDGozat.AutoSize = true;
            this.btnVTYDGozat.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnVTYDGozat.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnVTYDGozat.Location = new System.Drawing.Point(368, 13);
            this.btnVTYDGozat.Name = "btnVTYDGozat";
            this.btnVTYDGozat.Size = new System.Drawing.Size(70, 30);
            this.btnVTYDGozat.TabIndex = 9;
            this.btnVTYDGozat.Text = "Gözat...";
            this.btnVTYDGozat.UseVisualStyleBackColor = false;
            this.btnVTYDGozat.Click += new System.EventHandler(this.btnVTYDGozat_Click);
            // 
            // tbVTE
            // 
            this.tbVTE.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.tbVTE.Location = new System.Drawing.Point(11, 49);
            this.tbVTE.Name = "tbVTE";
            this.tbVTE.ReadOnly = true;
            this.tbVTE.Size = new System.Drawing.Size(231, 30);
            this.tbVTE.TabIndex = 11;
            // 
            // btnVTYedektenDon
            // 
            this.btnVTYedektenDon.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnVTYedektenDon.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnVTYedektenDon.Enabled = false;
            this.btnVTYedektenDon.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnVTYedektenDon.Location = new System.Drawing.Point(316, 49);
            this.btnVTYedektenDon.Name = "btnVTYedektenDon";
            this.btnVTYedektenDon.Size = new System.Drawing.Size(122, 30);
            this.btnVTYedektenDon.TabIndex = 9;
            this.btnVTYedektenDon.Text = "Yedekten Dön";
            this.btnVTYedektenDon.UseVisualStyleBackColor = false;
            this.btnVTYedektenDon.Click += new System.EventHandler(this.btnVTYedektenDon_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label3.Location = new System.Drawing.Point(6, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(183, 25);
            this.label3.TabIndex = 0;
            this.label3.Text = "VT Yedekten Dön";
            // 
            // groupBox4
            // 
            this.groupBox4.AutoSize = true;
            this.groupBox4.BackColor = System.Drawing.Color.Azure;
            this.groupBox4.Controls.Add(this.tbVTI);
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Controls.Add(this.btnVTYedekle);
            this.groupBox4.Controls.Add(this.btnVTYGozat);
            this.groupBox4.Location = new System.Drawing.Point(610, 12);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(444, 102);
            this.groupBox4.TabIndex = 12;
            this.groupBox4.TabStop = false;
            // 
            // tbVTI
            // 
            this.tbVTI.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.tbVTI.Location = new System.Drawing.Point(11, 51);
            this.tbVTI.Name = "tbVTI";
            this.tbVTI.ReadOnly = true;
            this.tbVTI.Size = new System.Drawing.Size(233, 30);
            this.tbVTI.TabIndex = 21;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label5.Location = new System.Drawing.Point(6, 18);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(125, 25);
            this.label5.TabIndex = 18;
            this.label5.Text = "VT Yedekle";
            // 
            // btnVTYedekle
            // 
            this.btnVTYedekle.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnVTYedekle.Enabled = false;
            this.btnVTYedekle.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnVTYedekle.Location = new System.Drawing.Point(358, 51);
            this.btnVTYedekle.Name = "btnVTYedekle";
            this.btnVTYedekle.Size = new System.Drawing.Size(80, 30);
            this.btnVTYedekle.TabIndex = 20;
            this.btnVTYedekle.Text = "Yedekle";
            this.btnVTYedekle.UseVisualStyleBackColor = false;
            this.btnVTYedekle.Click += new System.EventHandler(this.btnVTYedekle_Click);
            // 
            // btnVTYGozat
            // 
            this.btnVTYGozat.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnVTYGozat.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnVTYGozat.Location = new System.Drawing.Point(368, 15);
            this.btnVTYGozat.Name = "btnVTYGozat";
            this.btnVTYGozat.Size = new System.Drawing.Size(70, 30);
            this.btnVTYGozat.TabIndex = 19;
            this.btnVTYGozat.Text = "Gözat...";
            this.btnVTYGozat.UseVisualStyleBackColor = false;
            this.btnVTYGozat.Click += new System.EventHandler(this.btnVTYGozat_Click);
            // 
            // VeriMenusu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.Azure;
            this.ClientSize = new System.Drawing.Size(1076, 506);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.txtKontrol);
            this.Name = "VeriMenusu";
            this.Text = "VeriMenusu";
            this.Load += new System.EventHandler(this.VeriMenusu_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label txtKontrol;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox inputDizini;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonImp;
        private System.Windows.Forms.Button btngozat;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnExpGozat;
        private System.Windows.Forms.TextBox tbExportPath;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnVTYDGozat;
        private System.Windows.Forms.TextBox tbVTE;
        private System.Windows.Forms.Button btnVTYedektenDon;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox tbVTI;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnVTYedekle;
        private System.Windows.Forms.Button btnVTYGozat;
    }
}