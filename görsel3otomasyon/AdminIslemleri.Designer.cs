
namespace görsel3otomasyon
{
    partial class AdminIslemleri
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
            this.dgvCalisanlar = new System.Windows.Forms.DataGridView();
            this.btnCalDuzenle = new System.Windows.Forms.Button();
            this.btnCalCikar = new System.Windows.Forms.Button();
            this.btnCalEkle = new System.Windows.Forms.Button();
            this.btnListe = new System.Windows.Forms.Button();
            this.txtAd = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtSoyad = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.maskedTCKimlik = new System.Windows.Forms.MaskedTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.comboYetki = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtPass = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCalisanlar)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvCalisanlar
            // 
            this.dgvCalisanlar.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvCalisanlar.BackgroundColor = System.Drawing.Color.Azure;
            this.dgvCalisanlar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCalisanlar.Location = new System.Drawing.Point(401, 132);
            this.dgvCalisanlar.Name = "dgvCalisanlar";
            this.dgvCalisanlar.RowHeadersWidth = 51;
            this.dgvCalisanlar.RowTemplate.Height = 24;
            this.dgvCalisanlar.Size = new System.Drawing.Size(663, 362);
            this.dgvCalisanlar.TabIndex = 1;
            this.dgvCalisanlar.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvCalisanlar_CellClick);
            // 
            // btnCalDuzenle
            // 
            this.btnCalDuzenle.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnCalDuzenle.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCalDuzenle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnCalDuzenle.Location = new System.Drawing.Point(542, 12);
            this.btnCalDuzenle.Name = "btnCalDuzenle";
            this.btnCalDuzenle.Size = new System.Drawing.Size(201, 60);
            this.btnCalDuzenle.TabIndex = 34;
            this.btnCalDuzenle.Text = "Çalışan Düzenle";
            this.btnCalDuzenle.UseVisualStyleBackColor = false;
            this.btnCalDuzenle.Click += new System.EventHandler(this.btnCalDuzenle_Click);
            // 
            // btnCalCikar
            // 
            this.btnCalCikar.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnCalCikar.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCalCikar.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnCalCikar.Location = new System.Drawing.Point(305, 12);
            this.btnCalCikar.Name = "btnCalCikar";
            this.btnCalCikar.Size = new System.Drawing.Size(201, 60);
            this.btnCalCikar.TabIndex = 33;
            this.btnCalCikar.Text = "Çalışan Sil";
            this.btnCalCikar.UseVisualStyleBackColor = false;
            this.btnCalCikar.Click += new System.EventHandler(this.btnCalCikar_Click);
            // 
            // btnCalEkle
            // 
            this.btnCalEkle.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnCalEkle.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCalEkle.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnCalEkle.Location = new System.Drawing.Point(68, 12);
            this.btnCalEkle.Name = "btnCalEkle";
            this.btnCalEkle.Size = new System.Drawing.Size(201, 60);
            this.btnCalEkle.TabIndex = 32;
            this.btnCalEkle.Text = "Çalışan Ekle";
            this.btnCalEkle.UseVisualStyleBackColor = false;
            this.btnCalEkle.Click += new System.EventHandler(this.btnCalEkle_Click);
            // 
            // btnListe
            // 
            this.btnListe.AutoSize = true;
            this.btnListe.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnListe.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnListe.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnListe.Location = new System.Drawing.Point(779, 12);
            this.btnListe.Name = "btnListe";
            this.btnListe.Size = new System.Drawing.Size(201, 60);
            this.btnListe.TabIndex = 35;
            this.btnListe.Text = "Listeyi Çıktı Al";
            this.btnListe.UseVisualStyleBackColor = false;
            this.btnListe.Click += new System.EventHandler(this.btnListe_Click);
            // 
            // txtAd
            // 
            this.txtAd.Location = new System.Drawing.Point(95, 148);
            this.txtAd.Name = "txtAd";
            this.txtAd.Size = new System.Drawing.Size(196, 22);
            this.txtAd.TabIndex = 36;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label2.Location = new System.Drawing.Point(14, 145);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 25);
            this.label2.TabIndex = 37;
            this.label2.Text = "Ad:";
            // 
            // txtSoyad
            // 
            this.txtSoyad.Location = new System.Drawing.Point(95, 185);
            this.txtSoyad.Name = "txtSoyad";
            this.txtSoyad.Size = new System.Drawing.Size(196, 22);
            this.txtSoyad.TabIndex = 38;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label3.Location = new System.Drawing.Point(14, 182);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 25);
            this.label3.TabIndex = 39;
            this.label3.Text = "Soyad:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label4.Location = new System.Drawing.Point(14, 220);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(81, 25);
            this.label4.TabIndex = 40;
            this.label4.Text = "TC No.:";
            // 
            // maskedTCKimlik
            // 
            this.maskedTCKimlik.Location = new System.Drawing.Point(95, 224);
            this.maskedTCKimlik.Mask = "00000000000";
            this.maskedTCKimlik.Name = "maskedTCKimlik";
            this.maskedTCKimlik.Size = new System.Drawing.Size(196, 22);
            this.maskedTCKimlik.TabIndex = 41;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(72, 99);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 36);
            this.label1.TabIndex = 42;
            this.label1.Text = "Çalışanın:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label7.Location = new System.Drawing.Point(14, 262);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(61, 25);
            this.label7.TabIndex = 43;
            this.label7.Text = "Yetki:";
            this.label7.Click += new System.EventHandler(this.label7_Click);
            // 
            // comboYetki
            // 
            this.comboYetki.FormattingEnabled = true;
            this.comboYetki.Items.AddRange(new object[] {
            "admin",
            "guvenlik",
            "muhasebe",
            "sekreter"});
            this.comboYetki.Location = new System.Drawing.Point(95, 263);
            this.comboYetki.Name = "comboYetki";
            this.comboYetki.Size = new System.Drawing.Size(196, 24);
            this.comboYetki.TabIndex = 44;
            this.comboYetki.SelectedIndexChanged += new System.EventHandler(this.comboYetki_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label5.Location = new System.Drawing.Point(45, 311);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(224, 36);
            this.label5.TabIndex = 45;
            this.label5.Text = "Sistem Erişimi:";
            // 
            // txtUsername
            // 
            this.txtUsername.Location = new System.Drawing.Point(145, 361);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(196, 22);
            this.txtUsername.TabIndex = 46;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label6.Location = new System.Drawing.Point(14, 361);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(125, 25);
            this.label6.TabIndex = 47;
            this.label6.Text = "Kullanıcı Adı:";
            // 
            // txtPass
            // 
            this.txtPass.Location = new System.Drawing.Point(145, 398);
            this.txtPass.Name = "txtPass";
            this.txtPass.Size = new System.Drawing.Size(196, 22);
            this.txtPass.TabIndex = 48;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label8.Location = new System.Drawing.Point(14, 398);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(58, 25);
            this.label8.TabIndex = 49;
            this.label8.Text = "Şifre:";
            // 
            // AdminIslemleri
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.Azure;
            this.ClientSize = new System.Drawing.Size(1076, 506);
            this.Controls.Add(this.txtUsername);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtPass);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.comboYetki);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtAd);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtSoyad);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.maskedTCKimlik);
            this.Controls.Add(this.btnListe);
            this.Controls.Add(this.btnCalDuzenle);
            this.Controls.Add(this.btnCalCikar);
            this.Controls.Add(this.btnCalEkle);
            this.Controls.Add(this.dgvCalisanlar);
            this.Name = "AdminIslemleri";
            this.Text = "AdminIslemleri";
            this.Load += new System.EventHandler(this.AdminIslemleri_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvCalisanlar)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dgvCalisanlar;
        private System.Windows.Forms.Button btnCalDuzenle;
        private System.Windows.Forms.Button btnCalCikar;
        private System.Windows.Forms.Button btnCalEkle;
        private System.Windows.Forms.Button btnListe;
        private System.Windows.Forms.TextBox txtAd;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtSoyad;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.MaskedTextBox maskedTCKimlik;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboYetki;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtUsername;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtPass;
        private System.Windows.Forms.Label label8;
    }
}