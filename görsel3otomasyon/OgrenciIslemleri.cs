using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using ClosedXML.Excel;

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
    public partial class OgrenciIslemleri : Form
    {
        public OgrenciIslemleri()
        {
            InitializeComponent();
        }

        SqlConnection baglanti = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true");
        SqlCommand komut; 
        SqlDataReader oku; 
        SqlDataAdapter adapter; 
        DataTable tablo;
        

        private void OgrenciIslemleri_Load(object sender, EventArgs e)
        { //"OgrenciIslemleri" formunun yüklendiği anında çalışacak
            ComboBoxDoldurucu();
            KisiGetir();
            dataOgrenciler.SelectionMode = DataGridViewSelectionMode.CellSelect;
        }

        private void KisiGetir()
        { 
            adapter = new SqlDataAdapter("SELECT * FROM Ogrenci", baglanti);//sorguyu veritabanında çalıştırılır
            tablo = new DataTable();//tablo olusturulur
            BaglantiAc();//bağlantıyı açar
            adapter.Fill(tablo);//dolduruyoruz
            dataOgrenciler.DataSource = tablo;
            dataOgrenciler.ReadOnly = true;//hücrelerin düzenlenmesi engelleniyor
            baglanti.Close();//bağlantıyı kapattık
        }
        

        private void ComboBoxDoldurucu()
        {
            // Birinci combobox
            BaglantiAc();//bağlantı açılır
            komut = new SqlCommand("select distinct bolumAd from Bolum", baglanti);//sorgu çalıştırılır
            oku = komut.ExecuteReader();//oku değişkenine atanır
            while (oku.Read())//döngü çalıştırılır
            {
                comboBolum.Items.Add(oku[0].ToString());
            }
            baglanti.Close();//bağlantı kapatılır
            // İkinci combobox
            BaglantiAc();//bağlantı açılır
            komut = new SqlCommand("select OdaNo from Oda where kalanKisi != kapasite and OdaNo!='000'", baglanti);//sorgu çalıştırılır
            oku = komut.ExecuteReader();//oku değişkenine atanır
            while (oku.Read())//döngü çalıştırılır
            {
                comboOda.Items.Add(oku[0].ToString());
            }
            baglanti.Close();//bağlantı kapatılır
        }
        private void BaglantiAc()
        {//bağlantıyı açmak için yazdım
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string sorgu = "Insert into Ogrenci (TcKimlik,Ad,Soyad,Telefon,BolumAd,Iban,OdaNo,Adres,VeliAd,VeliSoyad,VeliTelefon) values (@TcKimlik,@Ad,@Soyad,@Telefon,@Bolum,@Iban,@Oda,@Adres,@VeliAd,@VeliSoyad,@VeliTelefon)";//veritabanına eklemek için
            //verileri alıyoz
            komut = new SqlCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@TcKimlik", maskedTCKimlik.Text); 
            komut.Parameters.AddWithValue("@Ad", txtOgrAd.Text); 
            komut.Parameters.AddWithValue("@Soyad", txtOgrSoyad.Text);
            komut.Parameters.AddWithValue("@Telefon", maskedOgrTel.Text); 
            komut.Parameters.AddWithValue("@Bolum", comboBolum.Text);
            komut.Parameters.AddWithValue("@Iban", maskedIban.Text); 
            komut.Parameters.AddWithValue("@Oda", comboOda.Text); 
            komut.Parameters.AddWithValue("@Adres", richAdres.Text);
            komut.Parameters.AddWithValue("@VeliAd", txtVeliAd.Text);
            komut.Parameters.AddWithValue("@VeliSoyad", txtVeliSoyad.Text);
            komut.Parameters.AddWithValue("@VeliTelefon", maskedVeliTelefon.Text); 
            BaglantiAc();//bağlantıyı açtık
            try
            {
                komut.ExecuteNonQuery();//sqle ekliyoruz
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show("HATA! \n" + sqlEx.Message); //hata mesajı alıyoruz
            }
            catch (Exception)
            {
                MessageBox.Show("HATA!");//genel bir hata mesajı veriyor
            }
            baglanti.Close();//bağlantıyı kapattık
            KisiGetir();//kişi getirildi
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string sorgu = "DELETE Ogrenci Where OgrId=@id";//silme sorgusunu yazdık
            komut = new SqlCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@id", dataOgrenciler.CurrentRow.Cells[0].Value);//ilk id değeri belirleniyor
            BaglantiAc();//bağlantıyı açtık
            komut.ExecuteNonQuery();//siliniyor
            baglanti.Close();//bağlantıyı kapattık
            KisiGetir();//kişi getir
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string sorgu = "Update Ogrenci Set TcKimlik=@TcKimlik,Ad=@Ad,Soyad=@Soyad,Telefon=@Telefon,BolumAd=@Bolum,Iban=@Iban,OdaNo=@Oda,Adres=@Adres,VeliAd=@VeliAd,VeliSoyad=@VeliSoyad,VeliTelefon=@VeliTelefon Where OgrId=@id";//veritabanına eklemek için
            //verileri alıyoz
            komut = new SqlCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@TcKimlik", maskedTCKimlik.Text);
            komut.Parameters.AddWithValue("@Ad", txtOgrAd.Text);
            komut.Parameters.AddWithValue("@Soyad", txtOgrSoyad.Text);
            komut.Parameters.AddWithValue("@Telefon", maskedOgrTel.Text);
            komut.Parameters.AddWithValue("@Bolum", comboBolum.Text);
            komut.Parameters.AddWithValue("@Iban", maskedIban.Text);
            komut.Parameters.AddWithValue("@Oda", comboOda.Text);
            komut.Parameters.AddWithValue("@Adres", richAdres.Text);
            komut.Parameters.AddWithValue("@VeliAd", txtVeliAd.Text);
            komut.Parameters.AddWithValue("@VeliSoyad", txtVeliSoyad.Text);
            komut.Parameters.AddWithValue("@VeliTelefon", maskedVeliTelefon.Text);
            komut.Parameters.AddWithValue("@id", dataOgrenciler.CurrentRow.Cells[0].Value);
            BaglantiAc();//bağlantıyı açtık
            komut.ExecuteNonQuery();//sql e ekliyoz güncelliyoruz
            baglanti.Close();//bağlantıyı kapattık
            KisiGetir();//kişileri getirdik
        }

        private void dataOgrenciler_CellClick(object sender, DataGridViewCellEventArgs e)
        {//veriler sql deki yerine göre şey yapılır.
            maskedTCKimlik.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[1].Value);
            txtOgrAd.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[2].Value);
            txtOgrSoyad.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[3].Value);
            maskedOgrTel.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[4].Value);
            richAdres.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[5].Value);
            maskedIban.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[6].Value);
            comboOda.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[9].Value);
            comboBolum.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[10].Value);
            txtVeliAd.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[11].Value);
            txtVeliSoyad.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[12].Value);
            maskedVeliTelefon.Text = Convert.ToString(this.dataOgrenciler.CurrentRow.Cells[13].Value);
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (txtAra.Text.ToString() != "")//boş olup olmadığını kontrol eder
            { //eğer boş değilse seçilen şeye göre farlı işlemler yapar
                if (comboArama.SelectedIndex == 0)
                {//tc ye göre
                    TcAra(txtAra.Text.ToString());
                }
                else if (comboArama.SelectedIndex == 1)
                {//isme göre
                    AdAra(txtAra.Text.ToString());
                }
                else//diğer seçilen şeylere göre
                {
                    string sorgu = "select * from Ogrenci where " + SecilenArama(comboArama.SelectedIndex) + " = '" + txtAra.Text.ToString() + "'";
                    adapter = new SqlDataAdapter(sorgu, baglanti);
                    tablo = new DataTable();
                    adapter.Fill(tablo);
                    dataOgrenciler.DataSource = tablo;
                    dataOgrenciler.ReadOnly = true;
                    baglanti.Close(); 
                }
                btnXlsx.Enabled = true;
            }
            else//kutucuğu seçmediğimiz içğin hata verir
            {
                MessageBox.Show("Lütfen kutucuğu doldurun.");
            }
        }

        private void TcAra(string dis)
        {
            string sorgu = "exec spTCArama "+ dis;
            adapter = new SqlDataAdapter(sorgu, baglanti);
            tablo = new DataTable();
            BaglantiAc();//öğrencilerin bağlantısını açtık sql de
            adapter.Fill(tablo);//dolduruyoruz
            dataOgrenciler.DataSource = tablo;
            dataOgrenciler.ReadOnly = true;//hücrelerin düzenlenmesi engelleniyor
            baglanti.Close();//bağlantıyı kapattık//bağlantıyı kapattık
        }
        private void AdAra(string dis)
        {
            string sorgu = "exec spAdlaArama "+ dis;
            adapter = new SqlDataAdapter(sorgu, baglanti);
            tablo = new DataTable();
            BaglantiAc();//öğrencilerin bağlantısını açtık sql de
            adapter.Fill(tablo);//dolduruyoruz
            dataOgrenciler.DataSource = tablo;
            dataOgrenciler.ReadOnly = true;//hücrelerin düzenlenmesi engelleniyor
            baglanti.Close();//bağlantıyı kapattık
        } 
        private string SecilenArama(int secilenIndex)
        {//seçiyoruz yani bu işe yarar
            switch (secilenIndex)
            {
                case 0:
                    return "TcKimlik";
                case 1:
                    return "Ad";
                case 2:
                    return "Soyad";
                case 3:
                    return "Telefon";
                case 4:
                    return "BolumAd"; 
                case 6:
                    return "OdendiMi"; 
                default:
                    return "OdaNo";
            }
        }

        private void dataOgrenciler_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {// "OgrenciListesi" adında bir form açık ise burada yapılacak işlemler olacak
            if ((Application.OpenForms["OgrenciListesi"] as OgrenciListesi) != null)
            { 
            }
            else// "OgrenciListesi" adında bir form açık değil ise burada yapılacak işlemler olacak
                // Yeni bir OgrenciListesi formu oluşturulacak ve gösterilecek.
            {
                OgrenciListesi ogr = new OgrenciListesi();
                ogr.Show();
            }
        }

        private void comboArama_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtAra_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnXlsx_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Sayfası| *.xlsx" })//kaydetme iletişim kutusunu açar
            {
                if (sfd.ShowDialog() == DialogResult.OK)//kullanıcı dosya seçer
                {
                    try
                    {
                        using (XLWorkbook workbook = new XLWorkbook())
                        {
                            workbook.Worksheets.Add(tablo, "Öğrenciler");
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
