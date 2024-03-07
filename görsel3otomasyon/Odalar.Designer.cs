
namespace görsel3otomasyon
{
    partial class Odalar
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
            this.tableOdalar = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dataOdadakiler = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.btnHata = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.tableOdalar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataOdadakiler)).BeginInit();
            this.SuspendLayout();
            // 
            // tableOdalar
            // 
            this.tableOdalar.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tableOdalar.BackgroundColor = System.Drawing.Color.Azure;
            this.tableOdalar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tableOdalar.Location = new System.Drawing.Point(12, 81);
            this.tableOdalar.Name = "tableOdalar";
            this.tableOdalar.RowHeadersWidth = 51;
            this.tableOdalar.RowTemplate.Height = 24;
            this.tableOdalar.Size = new System.Drawing.Size(515, 290);
            this.tableOdalar.TabIndex = 0;
            this.tableOdalar.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tableOdalar_CellClick);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(178, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(179, 32);
            this.label1.TabIndex = 1;
            this.label1.Text = "Yurt Odaları";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label2.Location = new System.Drawing.Point(670, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(279, 32);
            this.label2.TabIndex = 2;
            this.label2.Text = "Odadaki Öğrenciler";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // dataOdadakiler
            // 
            this.dataOdadakiler.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataOdadakiler.BackgroundColor = System.Drawing.Color.Azure;
            this.dataOdadakiler.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataOdadakiler.Location = new System.Drawing.Point(544, 81);
            this.dataOdadakiler.Name = "dataOdadakiler";
            this.dataOdadakiler.RowHeadersWidth = 51;
            this.dataOdadakiler.RowTemplate.Height = 24;
            this.dataOdadakiler.Size = new System.Drawing.Size(520, 290);
            this.dataOdadakiler.TabIndex = 3;
            this.dataOdadakiler.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tableOdadakiler_CellClick);
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label3.Location = new System.Drawing.Point(68, 438);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(944, 29);
            this.label3.TabIndex = 4;
            this.label3.Text = "Not: Odadaki öğrencileri görmek için görmek istediğiniz odaya tablodan basınız.";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // btnHata
            // 
            this.btnHata.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnHata.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnHata.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnHata.Location = new System.Drawing.Point(12, 377);
            this.btnHata.Name = "btnHata";
            this.btnHata.Size = new System.Drawing.Size(515, 42);
            this.btnHata.TabIndex = 5;
            this.btnHata.Text = "Odalarda Hata Olması Durumunda Basınız";
            this.btnHata.UseVisualStyleBackColor = false;
            this.btnHata.Click += new System.EventHandler(this.btnHata_Click);
            // 
            // Odalar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.Azure;
            this.ClientSize = new System.Drawing.Size(1076, 506);
            this.Controls.Add(this.btnHata);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dataOdadakiler);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tableOdalar);
            this.Name = "Odalar";
            this.Text = "Odalar";
            this.Load += new System.EventHandler(this.Odalar_Load);
            ((System.ComponentModel.ISupportInitialize)(this.tableOdalar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataOdadakiler)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView tableOdalar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView dataOdadakiler;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnHata;
    }
}