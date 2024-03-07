using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace görsel3otomasyon
/****************************************************************************
** Düzce Üniversitesi
Akçakoca MYO
Bilgisayar Teknolojileri Bölümü
** 
** 
Video Linki :
** ÖDEV NUMARASI
** ÖĞRENCİ ADI ömer faruk çetinkaya
** ÖĞRENCİ NUMARASI.:211501047
** ÖĞRENİM TÜRÜ normal öğretim
****************************************************************************/
{
    public partial class GirisCikis : Form
    {
        public GirisCikis()
        {
            InitializeComponent();
        }

        private void GirisCikis_Load(object sender, EventArgs e)
        {
            KisiGetir();
            GCGetir();
            dataOgrencilerTablosu.SelectionMode = DataGridViewSelectionMode.CellSelect;
        } 
        SqlConnection baglanti = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true"); 
        SqlDataAdapter adapter;
        DataTable tablo;
        SqlCommand komut;

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void dataOgrencilerTablosu_CellClick(object sender, DataGridViewCellEventArgs e)
        {//bastığımızda öğrencinin yurta olup olmadığını görürüz
            string durum = dataOgrencilerTablosu.CurrentRow.Cells[5].Value.ToString();
            if (durum == "EVET")
            { 
                txtDurum.Text = "Yurtta";
            }
            else
            {
                txtDurum.Text = "Yurtta Değil";
            }
        }

        private void BaglantiAc()
        {//bağlantıyı açarız
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
        }

        private void KisiGetir()
        {
            adapter = new SqlDataAdapter("SELECT OgrID,Ad,Soyad,OdaNo as 'Oda Numarası',TCKimlik, yurttaMi as 'Yurtta Mı' FROM Ogrenci", baglanti);//öğre4ncileri getirir
            tablo = new DataTable();
            BaglantiAc();//bağlantı açılıyor
            adapter.Fill(tablo);//veritabanından gelen verilerle dolduruluyor
            dataOgrencilerTablosu.DataSource = tablo;
            dataOgrencilerTablosu.ReadOnly = true;//sadece okunabilir yapar
            baglanti.Close();//bağlantı kapatılıyor
        } 
        private void GCGetir()
        {
            adapter = new SqlDataAdapter("select top (10) Ogrenci.ogrID,TcKimlik,Ad,Soyad, GirisCikisState, IslemTarihi from Ogrenci right outer join GirisCikis on Ogrenci.ogrID=GirisCikis.ogrID order by IslemTarihi desc", baglanti);//giriş çıkış verilerini getirir
            tablo = new DataTable();
            BaglantiAc();//bağlantı açılıyor
            adapter.Fill(tablo);//veritabanından gelen verilerle dolduruluyor
            dataVeriler.DataSource = tablo;//veriler tabloda gösterilir
            dataVeriler.ReadOnly = true;//sadece okunabilir yapar
            baglanti.Close();//bağlantı kapatılıyor
        } 
        private void GCGetir(string ogrID)
        {
            adapter = new SqlDataAdapter("select top (100) Ogrenci.ogrID,TcKimlik,Ad,Soyad, GirisCikisState, IslemTarihi from Ogrenci right outer join GirisCikis on Ogrenci.ogrID=GirisCikis.ogrID order by IslemTarihi desc", baglanti);//giriş cikis tabosundan belirli sütünleri seçer
            tablo = new DataTable();
            BaglantiAc();//bağlantı açılıyor
            adapter.Fill(tablo); //veritabanından gelen verilerle dolduruluyor
            dataVeriler.ReadOnly = true;//sadece okunabilir yapar
            baglanti.Close();//bağlantı kapatılıyor
            DataView dv = tablo.DefaultView;//filtre uygulanması sağlar
            dv.RowFilter = "TCKimlik Like '" + maskedTCKimlik.Text + "%'";//tc bulunuyor
            dataVeriler.DataSource = dv;//filtrelenmiş verileri gösterir
        } 
        private void button1_Click(object sender, EventArgs e)
        {
            adapter = new SqlDataAdapter("SELECT OgrID,Ad,Soyad,OdaNo as 'Oda Numarası',TCKimlik, yurttaMi as 'Yurtta Mı' FROM Ogrenci", baglanti);//öğrenci tabosundan belirli sütünleri seçer
            tablo = new DataTable();
            BaglantiAc();//bağlantı açılıyor
            adapter.Fill(tablo);//veritabanından gelen verilerle dolduruluyor
            baglanti.Close();//bağlantı kapatılıyor
            DataView dv = tablo.DefaultView;//filtre uygulanması sağlar
            dv.RowFilter = "TCKimlik Like '" + maskedTCKimlik.Text + "%'";//tc bulunuyor
            dataOgrencilerTablosu.DataSource = dv;//filtrelenmiş verileri gösterir
            GCGetir(maskedTCKimlik.Text);//işlemler getiriliyor
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void maskedTCKimlik_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void dataGCikis_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnGiris_Click(object sender, EventArgs e)
        {
            if (dataOgrencilerTablosu.CurrentRow.Cells[5].Value.ToString() == "HAYIR")
            {
                var sorgu = "Insert into GirisCikis(ogrID,GirisCikisState) values (@id,'GİRİŞ')";//öğrencinin giriş yaptığını belirtiyoruz
                string updSorgu = "Update Ogrenci set yurttaMi='EVET' where OgrID=@id";//evet olarak güncellenir
                komut = new SqlCommand(sorgu, baglanti);
                var komutIki = new SqlCommand(updSorgu, baglanti);
                komut.Parameters.AddWithValue("@id", dataOgrencilerTablosu.CurrentRow.Cells[0].Value);
                komutIki.Parameters.AddWithValue("@id", dataOgrencilerTablosu.CurrentRow.Cells[0].Value);//id değeri seçilir ve tc ile doldurulur
                BaglantiAc();//bağlantı açar
                komut.ExecuteNonQuery();//sorgu çalıştırılır
                komutIki.ExecuteNonQuery();//sorgu çalıştırılır
                baglanti.Close();//bağlantı kapatılır
                KisiGetir();
                GCGetir();
            }
            else
            {
                MessageBox.Show("Bu öğrenci zaten yurt içerisinde!");//hata mesajı
            }
            
        }

        private void btnCikis_Click(object sender, EventArgs e)
        {
            if (dataOgrencilerTablosu.CurrentRow.Cells[5].Value.ToString() == "EVET")
            {
                var sorgu = "Insert into GirisCikis(ogrID,GirisCikisState) values (@id,'ÇIKIŞ')";//öğrencinin çıkış yaptığını belirtiyoruz
                string updSorgu = "Update Ogrenci set yurttaMi='HAYIR' where OgrID=@id";//hayır olarak güncellenir
                komut = new SqlCommand(sorgu, baglanti);
                var komutIki = new SqlCommand(updSorgu, baglanti);
                komut.Parameters.AddWithValue("@id", dataOgrencilerTablosu.CurrentRow.Cells[0].Value);
                komutIki.Parameters.AddWithValue("@id", dataOgrencilerTablosu.CurrentRow.Cells[0].Value);//id değeri seçilir ve tc ile doldurulur
                BaglantiAc();//bağlantı açar
                komut.ExecuteNonQuery();//sorgu çalıştırılır
                komutIki.ExecuteNonQuery();//sorgu çalıştırılır
                baglanti.Close();//bağlantı kapatılır
                KisiGetir();
                GCGetir();
            }
            else
            {
                MessageBox.Show("Bu öğrenci zaten yurt dışında!");
            }

        }

        private void dataVeriler_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataOgrencilerTablosu_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnPDF_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Sayfası| *.xlsx" })//excel dosyası uzantısını yapar
            {
                if (sfd.ShowDialog() == DialogResult.OK)//dosya kaydetme şeysini açar
                {
                    try
                    {
                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            workbook.Worksheets.Add(tablo, "Giriş-Çıkış");//giriş çıkış verisinden alır dosya oluşturmak
                            workbook.SaveAs(sfd.FileName);//excel dosyasını kaydeder

                        }
                    }
                    catch (Exception ex)//hata oluşursa hata mesajı verir
                    {
                        MessageBox.Show("HATA! \n :" + ex.Message);
                    }
                }
            }
        }

        private void btnIzinMenu_Click(object sender, EventArgs e)
        {//izin menüsünü açar
            if ((Application.OpenForms["IzinMenu"] as IzinMenu) != null)
            { 
            }//hata oluşursa hata verir
            else
            { 
                IzinMenu izin = new IzinMenu();
                izin.Show();
            }
        }
    }
}
