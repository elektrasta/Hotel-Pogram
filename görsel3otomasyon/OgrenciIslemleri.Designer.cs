
namespace görsel3otomasyon
{
    partial class OgrenciIslemleri
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtOgrAd = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtOgrSoyad = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.maskedTCKimlik = new System.Windows.Forms.MaskedTextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBolum = new System.Windows.Forms.ComboBox();
            this.maskedOgrTel = new System.Windows.Forms.MaskedTextBox();
            this.maskedIban = new System.Windows.Forms.MaskedTextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboOda = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.txtVeliSoyad = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtVeliAd = new System.Windows.Forms.TextBox();
            this.maskedVeliTelefon = new System.Windows.Forms.MaskedTextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.richAdres = new System.Windows.Forms.RichTextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.OgrenciBilgi = new System.Windows.Forms.GroupBox();
            this.dataOgrenciler = new System.Windows.Forms.DataGridView();
            this.button4 = new System.Windows.Forms.Button();
            this.txtAra = new System.Windows.Forms.TextBox();
            this.comboArama = new System.Windows.Forms.ComboBox();
            this.button5 = new System.Windows.Forms.Button();
            this.ogrenciBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.ogrenciBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.btnXlsx = new System.Windows.Forms.Button();
            this.OgrenciBilgi.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataOgrenciler)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ogrenciBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ogrenciBindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(15, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(137, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "Öğrencinin:";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // txtOgrAd
            // 
            this.txtOgrAd.Location = new System.Drawing.Point(96, 72);
            this.txtOgrAd.Name = "txtOgrAd";
            this.txtOgrAd.Size = new System.Drawing.Size(196, 22);
            this.txtOgrAd.TabIndex = 1;
            this.txtOgrAd.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label2.Location = new System.Drawing.Point(15, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 25);
            this.label2.TabIndex = 2;
            this.label2.Text = "Ad:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label3.Location = new System.Drawing.Point(15, 106);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 25);
            this.label3.TabIndex = 4;
            this.label3.Text = "Soyad:";
            // 
            // txtOgrSoyad
            // 
            this.txtOgrSoyad.Location = new System.Drawing.Point(96, 109);
            this.txtOgrSoyad.Name = "txtOgrSoyad";
            this.txtOgrSoyad.Size = new System.Drawing.Size(196, 22);
            this.txtOgrSoyad.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label4.Location = new System.Drawing.Point(15, 144);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(81, 25);
            this.label4.TabIndex = 6;
            this.label4.Text = "TC No.:";
            // 
            // maskedTCKimlik
            // 
            this.maskedTCKimlik.Location = new System.Drawing.Point(96, 148);
            this.maskedTCKimlik.Mask = "00000000000";
            this.maskedTCKimlik.Name = "maskedTCKimlik";
            this.maskedTCKimlik.Size = new System.Drawing.Size(196, 22);
            this.maskedTCKimlik.TabIndex = 7;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label6.Location = new System.Drawing.Point(6, 181);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(84, 25);
            this.label6.TabIndex = 11;
            this.label6.Text = "Telefon:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label7.Location = new System.Drawing.Point(15, 217);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(73, 25);
            this.label7.TabIndex = 9;
            this.label7.Text = "Bölüm:";
            this.label7.Click += new System.EventHandler(this.label7_Click);
            // 
            // comboBolum
            // 
            this.comboBolum.FormattingEnabled = true;
            this.comboBolum.Location = new System.Drawing.Point(96, 218);
            this.comboBolum.Name = "comboBolum";
            this.comboBolum.Size = new System.Drawing.Size(196, 24);
            this.comboBolum.TabIndex = 14;
            this.comboBolum.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // maskedOgrTel
            // 
            this.maskedOgrTel.Location = new System.Drawing.Point(96, 185);
            this.maskedOgrTel.Mask = "(999) 000-0000";
            this.maskedOgrTel.Name = "maskedOgrTel";
            this.maskedOgrTel.Size = new System.Drawing.Size(196, 22);
            this.maskedOgrTel.TabIndex = 15;
            // 
            // maskedIban
            // 
            this.maskedIban.Location = new System.Drawing.Point(96, 257);
            this.maskedIban.Mask = "TR000000000000000000000000";
            this.maskedIban.Name = "maskedIban";
            this.maskedIban.Size = new System.Drawing.Size(196, 22);
            this.maskedIban.TabIndex = 17;
            this.maskedIban.MaskInputRejected += new System.Windows.Forms.MaskInputRejectedEventHandler(this.maskedTextBox3_MaskInputRejected);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label5.Location = new System.Drawing.Point(15, 253);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(64, 25);
            this.label5.TabIndex = 16;
            this.label5.Text = "IBAN:";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // comboOda
            // 
            this.comboOda.FormattingEnabled = true;
            this.comboOda.Location = new System.Drawing.Point(96, 296);
            this.comboOda.Name = "comboOda";
            this.comboOda.Size = new System.Drawing.Size(196, 24);
            this.comboOda.TabIndex = 19;
            this.comboOda.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label8.Location = new System.Drawing.Point(23, 295);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(56, 25);
            this.label8.TabIndex = 18;
            this.label8.Text = "Oda:";
            this.label8.Click += new System.EventHandler(this.label8_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label9.Location = new System.Drawing.Point(322, 32);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(110, 29);
            this.label9.TabIndex = 20;
            this.label9.Text = "Velisinin:";
            this.label9.Click += new System.EventHandler(this.label9_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label10.Location = new System.Drawing.Point(322, 109);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(75, 25);
            this.label10.TabIndex = 24;
            this.label10.Text = "Soyad:";
            // 
            // txtVeliSoyad
            // 
            this.txtVeliSoyad.Location = new System.Drawing.Point(403, 112);
            this.txtVeliSoyad.Name = "txtVeliSoyad";
            this.txtVeliSoyad.Size = new System.Drawing.Size(196, 22);
            this.txtVeliSoyad.TabIndex = 23;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label11.Location = new System.Drawing.Point(322, 72);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(43, 25);
            this.label11.TabIndex = 22;
            this.label11.Text = "Ad:";
            // 
            // txtVeliAd
            // 
            this.txtVeliAd.Location = new System.Drawing.Point(403, 75);
            this.txtVeliAd.Name = "txtVeliAd";
            this.txtVeliAd.Size = new System.Drawing.Size(196, 22);
            this.txtVeliAd.TabIndex = 21;
            // 
            // maskedVeliTelefon
            // 
            this.maskedVeliTelefon.Location = new System.Drawing.Point(403, 149);
            this.maskedVeliTelefon.Mask = "(999) 000-0000";
            this.maskedVeliTelefon.Name = "maskedVeliTelefon";
            this.maskedVeliTelefon.Size = new System.Drawing.Size(196, 22);
            this.maskedVeliTelefon.TabIndex = 26;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label12.Location = new System.Drawing.Point(313, 145);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(84, 25);
            this.label12.TabIndex = 25;
            this.label12.Text = "Telefon:";
            // 
            // richAdres
            // 
            this.richAdres.Location = new System.Drawing.Point(96, 335);
            this.richAdres.Name = "richAdres";
            this.richAdres.Size = new System.Drawing.Size(196, 96);
            this.richAdres.TabIndex = 27;
            this.richAdres.Text = "";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label13.Location = new System.Drawing.Point(18, 335);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(70, 25);
            this.label13.TabIndex = 28;
            this.label13.Text = "Adres:";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.button1.Location = new System.Drawing.Point(108, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(201, 60);
            this.button1.TabIndex = 29;
            this.button1.Text = "Öğrenci Ekle";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.WhiteSmoke;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.button2.Location = new System.Drawing.Point(345, 12);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(201, 60);
            this.button2.TabIndex = 30;
            this.button2.Text = "Öğrenci Çıkar";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.WhiteSmoke;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.button3.Location = new System.Drawing.Point(582, 12);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(201, 60);
            this.button3.TabIndex = 31;
            this.button3.Text = "Öğrenci Düzenle";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // OgrenciBilgi
            // 
            this.OgrenciBilgi.BackColor = System.Drawing.Color.Azure;
            this.OgrenciBilgi.Controls.Add(this.label1);
            this.OgrenciBilgi.Controls.Add(this.txtOgrAd);
            this.OgrenciBilgi.Controls.Add(this.label2);
            this.OgrenciBilgi.Controls.Add(this.label13);
            this.OgrenciBilgi.Controls.Add(this.txtOgrSoyad);
            this.OgrenciBilgi.Controls.Add(this.richAdres);
            this.OgrenciBilgi.Controls.Add(this.label3);
            this.OgrenciBilgi.Controls.Add(this.maskedVeliTelefon);
            this.OgrenciBilgi.Controls.Add(this.label4);
            this.OgrenciBilgi.Controls.Add(this.label12);
            this.OgrenciBilgi.Controls.Add(this.maskedTCKimlik);
            this.OgrenciBilgi.Controls.Add(this.label10);
            this.OgrenciBilgi.Controls.Add(this.label7);
            this.OgrenciBilgi.Controls.Add(this.txtVeliSoyad);
            this.OgrenciBilgi.Controls.Add(this.label6);
            this.OgrenciBilgi.Controls.Add(this.label11);
            this.OgrenciBilgi.Controls.Add(this.comboBolum);
            this.OgrenciBilgi.Controls.Add(this.txtVeliAd);
            this.OgrenciBilgi.Controls.Add(this.maskedOgrTel);
            this.OgrenciBilgi.Controls.Add(this.label9);
            this.OgrenciBilgi.Controls.Add(this.label5);
            this.OgrenciBilgi.Controls.Add(this.comboOda);
            this.OgrenciBilgi.Controls.Add(this.maskedIban);
            this.OgrenciBilgi.Controls.Add(this.label8);
            this.OgrenciBilgi.Location = new System.Drawing.Point(12, 88);
            this.OgrenciBilgi.Name = "OgrenciBilgi";
            this.OgrenciBilgi.Size = new System.Drawing.Size(609, 447);
            this.OgrenciBilgi.TabIndex = 32;
            this.OgrenciBilgi.TabStop = false;
            // 
            // dataOgrenciler
            // 
            this.dataOgrenciler.BackgroundColor = System.Drawing.Color.Azure;
            this.dataOgrenciler.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataOgrenciler.Location = new System.Drawing.Point(645, 172);
            this.dataOgrenciler.Name = "dataOgrenciler";
            this.dataOgrenciler.RowHeadersWidth = 51;
            this.dataOgrenciler.RowTemplate.Height = 24;
            this.dataOgrenciler.Size = new System.Drawing.Size(466, 363);
            this.dataOgrenciler.TabIndex = 33;
            this.dataOgrenciler.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataOgrenciler_CellClick);
            this.dataOgrenciler.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataOgrenciler_CellContentClick);
            // 
            // button4
            // 
            this.button4.AutoSize = true;
            this.button4.BackColor = System.Drawing.Color.WhiteSmoke;
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.button4.Location = new System.Drawing.Point(805, 12);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(201, 60);
            this.button4.TabIndex = 34;
            this.button4.Text = "Öğrenci Listesi";
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // txtAra
            // 
            this.txtAra.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtAra.Location = new System.Drawing.Point(645, 136);
            this.txtAra.Name = "txtAra";
            this.txtAra.Size = new System.Drawing.Size(223, 27);
            this.txtAra.TabIndex = 29;
            this.txtAra.TextChanged += new System.EventHandler(this.txtAra_TextChanged);
            // 
            // comboArama
            // 
            this.comboArama.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.comboArama.FormattingEnabled = true;
            this.comboArama.Items.AddRange(new object[] {
            "TC Kimlik",
            "Ad",
            "Soyad",
            "Telefon",
            "Bölüm",
            "Oda",
            "Ödeme Durumu"});
            this.comboArama.Location = new System.Drawing.Point(874, 136);
            this.comboArama.Name = "comboArama";
            this.comboArama.Size = new System.Drawing.Size(83, 28);
            this.comboArama.TabIndex = 29;
            this.comboArama.SelectedIndexChanged += new System.EventHandler(this.comboArama_SelectedIndexChanged);
            // 
            // button5
            // 
            this.button5.BackColor = System.Drawing.Color.WhiteSmoke;
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.button5.Location = new System.Drawing.Point(963, 136);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(83, 30);
            this.button5.TabIndex = 35;
            this.button5.Text = "ARA";
            this.button5.UseVisualStyleBackColor = false;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // ogrenciBindingSource
            // 
            this.ogrenciBindingSource.DataMember = "Ogrenci";
            // 
            // ogrenciBindingSource1
            // 
            this.ogrenciBindingSource1.DataMember = "Ogrenci";
            // 
            // btnXlsx
            // 
            this.btnXlsx.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnXlsx.Enabled = false;
            this.btnXlsx.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnXlsx.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnXlsx.Location = new System.Drawing.Point(1053, 136);
            this.btnXlsx.Name = "btnXlsx";
            this.btnXlsx.Size = new System.Drawing.Size(58, 30);
            this.btnXlsx.TabIndex = 36;
            this.btnXlsx.Text = "XLSX";
            this.btnXlsx.UseVisualStyleBackColor = false;
            this.btnXlsx.Click += new System.EventHandler(this.btnXlsx_Click);
            // 
            // OgrenciIslemleri
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.Azure;
            this.ClientSize = new System.Drawing.Size(1123, 543);
            this.Controls.Add(this.btnXlsx);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.comboArama);
            this.Controls.Add(this.txtAra);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.dataOgrenciler);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.OgrenciBilgi);
            this.Name = "OgrenciIslemleri";
            this.Text = "OgrenciIslemleri";
            this.Load += new System.EventHandler(this.OgrenciIslemleri_Load);
            this.OgrenciBilgi.ResumeLayout(false);
            this.OgrenciBilgi.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataOgrenciler)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ogrenciBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ogrenciBindingSource1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtOgrAd;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtOgrSoyad;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.MaskedTextBox maskedTCKimlik;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBolum;
        private System.Windows.Forms.MaskedTextBox maskedOgrTel;
        private System.Windows.Forms.MaskedTextBox maskedIban;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboOda;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtVeliSoyad;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtVeliAd;
        private System.Windows.Forms.MaskedTextBox maskedVeliTelefon;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.RichTextBox richAdres;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.GroupBox OgrenciBilgi; 
        private System.Windows.Forms.BindingSource ogrenciBindingSource; 
        private System.Windows.Forms.DataGridView dataOgrenciler;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox txtAra;
        private System.Windows.Forms.ComboBox comboArama;
        private System.Windows.Forms.Button button5; 
        private System.Windows.Forms.BindingSource ogrenciBindingSource1;
        private System.Windows.Forms.Button btnXlsx;
    }
}