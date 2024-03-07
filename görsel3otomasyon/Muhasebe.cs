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
    public partial class Muhasebe : Form
    {
        SqlConnection baglanti = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true"); 
        SqlDataAdapter adapter;
        DataTable tablo;

        public Muhasebe()
        {
            InitializeComponent();
        }

        private void Muhasebe_Load(object sender, EventArgs e)
        {//form başla başlamaz getirilir
            KisiGetir();
            DekontGetir();
        } 

        private void dataOgrenci_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (maskedTCKimlik.Text.ToString() != "")//boş olup olmadığını kontrol eder
            {
                KisiGetir(maskedTCKimlik.Text.ToString());//tc ye göre kişiyi getirir
                DekontGetir(maskedTCKimlik.Text.ToString()); //tc ye göre dekontu getirir
            }
            else
            {
                KisiGetir();//hepsini getirir
                DekontGetir();//hepsini getirir
            }
        }
        private void KisiGetir(string TC)
        {
            //öğrenci ve izin tablolarındaki tc ye göre muhasebe bilgilerine göre ilk 50 kaydı sıralar
            adapter = new SqlDataAdapter("SELECT ogrID,Ad,Soyad,OdendiMi FROM Ogrenci where TcKimlik = '"+TC+"'", baglanti);
            tablo = new DataTable();
            BaglantiAc();//bağlantıyı açarız
            adapter.Fill(tablo);//tablo nesnesi doldurulur
            dataOgrenci.DataSource = tablo;
            dataOgrenci.ReadOnly = true;//sadece salt okunur sadece okunabilir yani
            baglanti.Close();//bağlantıdan çıkıldı
        }
        private void KisiGetir()
        {
            //öğrenci ve izin tablolarındaki tc ye göre muhasebe bilgileri listelenir
            adapter = new SqlDataAdapter("SELECT ogrID,Ad,Soyad,OdendiMi FROM Ogrenci", baglanti);
            tablo = new DataTable();
            BaglantiAc();//bağlantıyı açarız
            adapter.Fill(tablo);//tablo nesnesi doldurulur
            dataOgrenci.DataSource = tablo;
            dataOgrenci.ReadOnly = true;//sadece salt okunur sadece okunabilir yani
            baglanti.Close();//bağlantıdan çıkıldı
        }

        private void DekontGetir()
        {
            //öğrenci ve izin tablolarındaki tc ye göre dekont bilgileri listelenir
            adapter = new SqlDataAdapter("select top (100) Ad,Soyad,odemeTarihi as 'Ödeme Tarihi',tutar as 'Tutar' from Ogrenci full outer join Muhasebe on Ogrenci.ogrID = Muhasebe.ogrID", baglanti);
            tablo = new DataTable();
            BaglantiAc();//bağlantıyı açarız
            adapter.Fill(tablo);//tablo nesnesi doldurulur
            dataDekontlar.DataSource = tablo;
            dataDekontlar.ReadOnly = true;//sadece salt okunur sadece okunabilir yani
            baglanti.Close();//bağlantıdan çıkıldı
        }
        private void DekontGetir(string TC)
        {
            //öğrenci ve izin tablolarındaki tc ye göre dekont bilgilerine göre ilk 50 kaydı sıralar
            adapter = new SqlDataAdapter("select top (100) Ad,Soyad,odemeTarihi as 'Ödeme Tarihi',tutar as 'Tutar' from Ogrenci full outer join Muhasebe on Ogrenci.ogrID = Muhasebe.ogrID where Ogrenci.TcKimlik ='" + TC + "'", baglanti);
            tablo = new DataTable();
            BaglantiAc();//bağlantıyı açarız
            adapter.Fill(tablo);//tablo nesnesi doldurulur
            dataDekontlar.DataSource = tablo;
            dataDekontlar.ReadOnly = true;//sadece salt okunur sadece okunabilir yani
            baglanti.Close();//bağlantıdan çıkıldı
        }
        private void BaglantiAc()
        {//ConnectionState.Closed ise (kapalı durumdaysa), bağlantıyı açmak için bir işlem yapılır.
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sorgu = "Insert into Muhasebe(ogrID,tutar) values("+ this.dataOgrenci.CurrentRow.Cells[0].Value + " , "+ textboxTutar.Text.ToString() +")";//girilen değeri veritabanına göre ekler
            var komut = new SqlCommand(sorgu, baglanti);  
            BaglantiAc();//bağlantıyı açtık
            komut.ExecuteNonQuery();//sorgunun çalıştırılmasını sağlar
            baglanti.Close();//bağlantıyı kapattık
            KisiGetir();
            DekontGetir();
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnPDF_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Sayfası| *.xlsx" })
            {//dosya kaydetme kutusu açar ve xlsx olarak kaydetme şeyini yapar
                if (sfd.ShowDialog() == DialogResult.OK)//kullanıcı dosya seçer ve tamam a basmasıyla olur
                {
                    try
                    {
                        using (XLWorkbook workbook = new XLWorkbook())//excel çalışma kitabı oluşturulur
                        {
                            workbook.Worksheets.Add(tablo, "Muhasebe");//tablo ekler ve sayfanın adını muhasebe olur
                            workbook.SaveAs(sfd.FileName);//seçilen dosya yoluna kaydeder

                        }
                    }
                    catch (Exception ex)
                    {//herhangi bir hata durumunda olur burası
                        MessageBox.Show("HATA! \n :" + ex.Message);
                    }
                }
            }
        }
    }
}
