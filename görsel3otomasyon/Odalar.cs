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
    public partial class Odalar : Form
    {
        SqlConnection baglanti = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true"); 
        SqlDataAdapter adapter;
        DataTable tablo;
        string sorgu;
        List<string> odalar = new List<string>();
        private string OdaNoTutucu;

        public Odalar()
        {
            InitializeComponent();
        }

        OgrenciIslemleri ogrenciIslemleri = new OgrenciIslemleri();

        private void Odalar_Load(object sender, EventArgs e)
        {
            KisiGetir();
        }
        private void KisiGetir()
        {
            //oda tablosuna tablolarındaki tc ye göre kalan kisi bilgileri listelenir
            sorgu = "SELECT OdaNo as 'Oda No.', kapasite as 'Kapasite', KalanKisi as 'Kalan Kişi' FROM Oda where OdaNo!='000'";
            adapter = new SqlDataAdapter(sorgu, baglanti);
            tablo = new DataTable();
            BaglantiAc();//bağlantı açıolır
            adapter.Fill(tablo);//tablo nesnesi doldurulur
            tableOdalar.DataSource = tablo;
            tableOdalar.ReadOnly = true;//salt okunur olur
            baglanti.Close();//bağlantıyı kapattık
        }
        private void BaglantiAc()
        {
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
        } 

        private void tableOdalar_CellClick(object sender, DataGridViewCellEventArgs e)
        { //ad soyad tc oda tutucu değişkenine atar
            OdaNoTutucu = Convert.ToString(this.tableOdalar.CurrentRow.Cells[0].Value);
            sorgu = "Select Ad,Soyad,TcKimlik from Ogrenci where OdaNo='"+ OdaNoTutucu +"'"; 
            adapter = new SqlDataAdapter(sorgu, baglanti);
            tablo = new DataTable();
            BaglantiAc();//bağlantı açıolır
            adapter.Fill(tablo);//tablo nesnesi doldurulur
            dataOdadakiler.DataSource = tablo;
            dataOdadakiler.ReadOnly = true;//salt okunur olur
            baglanti.Close();//bağlantıyı kapattık
        }

        private void tableOdadakiler_CellClick(object sender, DataGridViewCellEventArgs e)
        { 

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void btnHata_Click(object sender, EventArgs e)
        {
            SqlCommand komut = new SqlCommand("select OdaNo from Oda", baglanti);//bağlantıyı çalıştırır sorguyu çaıştırır yani
            BaglantiAc();//bağlantı açılır
            var oku= komut.ExecuteReader();//sorguyu çalıştırır ve oku değişkenine atar
            while (oku.Read())//döngüyü başlattık
            {
                odalar.Add(oku.GetString(0)); //oda no yu alır oda nesnesinde saklkanır
            }
            oku.Close(); //çıkılır
            for (int i = 0; i < odalar.Count(); i++)
            {
                //MessageBox.Show(odalar[i]);
                //kayıtlar güncellemek için yazdık
                SqlCommand komutIki = new SqlCommand("update Oda set kalanKisi=(select Count(Ad) from Ogrenci where odaNo='" + odalar[i] + "') where odaNo='"+odalar[i]+"'", baglanti);
                komutIki.ExecuteNonQuery();
            }
        }
    }
}
