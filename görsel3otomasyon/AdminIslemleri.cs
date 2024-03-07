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

namespace görsel3otomasyon
{
    public partial class AdminIslemleri : Form
    {
        SqlConnection baglanti = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true");
        DataTable tablo;
        public AdminIslemleri()
        {
            InitializeComponent();
            KisiGetir(); 
        }
        private void BaglantiAc()
        {
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
        } 
        private void KisiGetir()
        {
            var adapter = new SqlDataAdapter("SELECT * FROM Calisan", baglanti);//çalışanları getirdik
            tablo = new DataTable();//yeni tablo olusturdum
            BaglantiAc();//çalışanların bağlantısını açtık sql de
            adapter.Fill(tablo);//dolduruyoruz
            dgvCalisanlar.DataSource = tablo;//veriler görüntelenecek
            dgvCalisanlar.ReadOnly = true;//hücrelerin düzenlenmesi engelleniyor
            baglanti.Close();//bağlantıyı kapattık
        }
        private void AdminIslemleri_Load(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void comboYetki_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnCalEkle_Click(object sender, EventArgs e)
        {
            string sorgu = "Insert into Calisan (ad,soyad,tckimlik,vasif,username,pass) values (@ad,@soyad,@tckimlik,@vasif,@username,@pass)";//veritabanına eklemek için
            var komut = new SqlCommand(sorgu, baglanti);
            //verileri alıyoz
            komut.Parameters.AddWithValue("@tckimlik", maskedTCKimlik.Text); 
            komut.Parameters.AddWithValue("@ad", txtAd.Text); 
            komut.Parameters.AddWithValue("@soyad", txtSoyad.Text); 
            komut.Parameters.AddWithValue("@vasif", comboYetki.Text); 
            komut.Parameters.AddWithValue("@username", txtUsername.Text); 
            komut.Parameters.AddWithValue("@pass", txtPass.Text); 
            BaglantiAc();//bağlantıyı açtık
            try
            {
                komut.ExecuteNonQuery();//sqle ekliyoruz
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show("HATA! \n" + sqlEx.Message);//hata mesajı alıyoruz
            }
            catch (Exception)
            {
                MessageBox.Show("HATA!");//genel bir hata mesajı veriyor
            }
            baglanti.Close();//bağlantıyı kapattık
            KisiGetir();//kişi getirildi
        }

        private void btnCalCikar_Click(object sender, EventArgs e)
        {
            string sorgu = "DELETE Calisan Where id=@id";//silme sorgusunu yazdık
            var komut = new SqlCommand(sorgu, baglanti);//
            komut.Parameters.AddWithValue("@id", dgvCalisanlar.CurrentRow.Cells[0].Value);//ilk id değeri belirleniyor
            BaglantiAc();//bağlantıyı açtık
            komut.ExecuteNonQuery();//siliniyor
            baglanti.Close();//bağlantıyı kapattık
            KisiGetir();//lişi getir
        }

        private void btnCalDuzenle_Click(object sender, EventArgs e)
        {
            string sorgu = "Update Calisan set ad=@ad,soyad=@soyad,tckimlik=@tckimlik,vasif=@vasif,username=@username,pass=@pass where id=@id";//veritabanına eklemek için
            //verileri alıyoz
            var komut = new SqlCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@id", dgvCalisanlar.CurrentRow.Cells[0].Value);
            komut.Parameters.AddWithValue("@tckimlik", maskedTCKimlik.Text);
            komut.Parameters.AddWithValue("@ad", txtAd.Text);
            komut.Parameters.AddWithValue("@soyad", txtSoyad.Text);
            komut.Parameters.AddWithValue("@vasif", comboYetki.Text);
            komut.Parameters.AddWithValue("@username", txtUsername.Text);
            komut.Parameters.AddWithValue("@pass", txtPass.Text);
            BaglantiAc();//bağlantıyı açtık
            try
            {
                komut.ExecuteNonQuery();//sql e ekliyoz
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show("HATA! \n" + sqlEx.Message);//hata mesajı
            }
            catch (Exception)
            {
                MessageBox.Show("HATA!");//genel hata mesajı
            }
            baglanti.Close();//bağlantıyı kapattık
            KisiGetir();//kişileri getirdik
        }

        private void dgvCalisanlar_CellClick(object sender, DataGridViewCellEventArgs e)
        {//ilgili form kontrollerine değerler atanıyor ve stringe çevriliyor
            txtUsername.Text = Convert.ToString(this.dgvCalisanlar.CurrentRow.Cells[1].Value); 
            txtPass.Text = Convert.ToString(this.dgvCalisanlar.CurrentRow.Cells[2].Value); 
            comboYetki.Text = Convert.ToString(this.dgvCalisanlar.CurrentRow.Cells[3].Value); 
            txtAd.Text = Convert.ToString(this.dgvCalisanlar.CurrentRow.Cells[4].Value); 
            txtSoyad.Text = Convert.ToString(this.dgvCalisanlar.CurrentRow.Cells[5].Value);
            maskedTCKimlik.Text = Convert.ToString(this.dgvCalisanlar.CurrentRow.Cells[6].Value);  
        }

        private void btnListe_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Sayfası| *.xlsx" })//kaydetme iletişim kutusunu açar
            {
                if (sfd.ShowDialog() == DialogResult.OK)//kullanıcı dosya seçer
                {
                    try
                    {
                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            workbook.Worksheets.Add(tablo, "Çalışanlar");
                            workbook.SaveAs(sfd.FileName);//excel dosyasını kullanıcının seçtiği ad ile kaydediyor

                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("HATA! \n :" + ex.Message);//hata mesajını gösterir
                    }
                }
            }
        }
    }
}
