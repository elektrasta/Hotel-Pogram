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
    public partial class Giris : Form
    {
        SqlConnection baglanti = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true");   
        private string yetki;
        AnaMenu anaMenu = new AnaMenu();
        SqlDataReader reader;

        public Giris()
        {
            InitializeComponent();
        }

        private void BaglantiAc()
        {
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
        }

        private void girisButonu_Click(object sender, EventArgs e)
        {
            // Stored Procedure'suz
            //kontrol eder kullanıcı adı ve parola doğru mu diye
            SqlCommand sqlCommand = new SqlCommand("SELECT * from Calisan where username='" + textBox1.Text + "' and pass='" + textBox2.Text + "'", baglanti);
            BaglantiAc();//bağlantıyı açarız
            reader = sqlCommand.ExecuteReader();
            if (reader.Read())
            {
                reader.Close();
                AnaMenu anamenu = new AnaMenu();
                SqlCommand cmd = new SqlCommand("select vasif from Calisan where username = '" + textBox1.Text.ToString() + "'", baglanti);
                SqlDataReader areader = cmd.ExecuteReader();
                while (areader.Read())
                {
                    yetki = areader[0].ToString();
                }
                if (yetki == "admin")
                {//admin ile girersek hepsi gözükür
                    this.Hide();
                    anamenu.ShowDialog();
                    this.Close();

                }
                else if (yetki == "muhasebe")
                {//muhasebe ile girersek belli başlı şeyler gözükür
                    anamenu.button1.Enabled = false;
                    anamenu.button1.Visible = false;

                    anamenu.button2.Enabled = false;
                    anamenu.button2.Visible = false;

                    anamenu.button4.Enabled = false;
                    anamenu.button4.Visible = false;

                    anamenu.button5.Enabled = false;
                    anamenu.button5.Visible = false;

                    anamenu.button6.Enabled = false;
                    anamenu.button6.Visible = false;

                    this.Hide();
                    anamenu.ShowDialog();
                    this.Close();
                }
                else if (yetki == "guvenlik")
                {//güvenlik ile girersek belli başlı şeyler gözükür
                    anamenu.button1.Enabled = false;
                    anamenu.button1.Visible = false;

                    anamenu.button3.Enabled = false;
                    anamenu.button3.Visible = false;

                    anamenu.button4.Enabled = false;
                    anamenu.button4.Visible = false;

                    anamenu.button5.Enabled = false;
                    anamenu.button5.Visible = false;

                    anamenu.button6.Enabled = false;
                    anamenu.button6.Visible = false;

                    this.Hide();
                    anamenu.ShowDialog();
                    this.Close();
                }
                else if (yetki == "sekreter")
                {//sekreter ile girersek belli başlı şeyler gözükür
                    anamenu.button2.Enabled = false;
                    anamenu.button2.Visible = false;

                    anamenu.button3.Enabled = false;
                    anamenu.button3.Visible = false;

                    anamenu.button5.Enabled = false;
                    anamenu.button5.Visible = false;

                    anamenu.button6.Enabled = false;
                    anamenu.button6.Visible = false;

                    this.Hide();
                    anamenu.ShowDialog();
                    this.Close();
                }
            } 
            else 
            {
                MessageBox.Show("Hatalı giriş yaptınız");//hata mesajı verir
            }
            baglanti.Close();  //bağlantıdan çıkılır
        } 

        private void Giris_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
