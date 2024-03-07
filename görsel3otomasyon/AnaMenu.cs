using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
    public partial class AnaMenu : Form
    {

        public AnaMenu()
        {
            InitializeComponent(); 
        }


        private void AnaMenu_Load(object sender, EventArgs e)
        {
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DegisenPanel.Controls.Clear();//kaldırır ve yeni bir panelin eklenmesine hazırlık yapar
            OgrenciIslemleri ogrenciIslemleri = new OgrenciIslemleri();//öğrenci işlemleri formu olsturulur
            ogrenciIslemleri.MdiParent = this;//anaformun altında çocuk form açar
            ogrenciIslemleri.FormBorderStyle = FormBorderStyle.None;//kenarlık belirtir
            DegisenPanel.Controls.Add(ogrenciIslemleri);//öğrenci işlemleri formu açılır
            ogrenciIslemleri.Show();//öğrenci işlemleri formu gösterilir
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DegisenPanel.Controls.Clear();//kaldırır ve yeni bir panelin eklenmesine hazırlık yapar
            GirisCikis girisCikis = new GirisCikis();//giris cikis işlemleri formu olsturulur
            girisCikis.MdiParent = this;//anaformun altında çocuk form açar
            girisCikis.FormBorderStyle = FormBorderStyle.None;//kenarlık belirtir
            DegisenPanel.Controls.Add(girisCikis);//giris cikis işlemleri formu açılır
            girisCikis.Show();//giris cikis işlemleri formu gösterilir
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DegisenPanel.Controls.Clear();//kaldırır ve yeni bir panelin eklenmesine hazırlık yapar
            Odalar odalar = new Odalar();//odalar işlemleri formu olsturulur
            odalar.MdiParent = this;//anaformun altında çocuk form açar
            odalar.FormBorderStyle = FormBorderStyle.None;//kenarlık belirtir
            DegisenPanel.Controls.Add(odalar);//odalar işlemleri formu açılır
            odalar.Show();//odalar işlemleri formu gösterilir
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DegisenPanel.Controls.Clear();//kaldırır ve yeni bir panelin eklenmesine hazırlık yapar
            Muhasebe muh = new Muhasebe();//muhasebe işlemleri formu olsturulur
            muh.MdiParent = this;//anaformun altında çocuk form açar
            muh.FormBorderStyle = FormBorderStyle.None;//kenarlık belirtir
            DegisenPanel.Controls.Add(muh);//muhasebe işlemleri formu açılır
            muh.Show();//muhasebe işlemleri formu gösterilir

        }

        private void btnCikisYap_Click(object sender, EventArgs e)
        {//çıkış yapıyor
            Giris cikis = new Giris();
            this.Hide();
            cikis.ShowDialog();
            this.Close();


        }

        private void button6_Click(object sender, EventArgs e)
        {

            AdminIslemleri adm = new AdminIslemleri();
            DegisenPanel.Controls.Clear();//admin işlemleri formu olsturulur
            adm.MdiParent = this;//anaformun altında çocuk form açar
            adm.FormBorderStyle = FormBorderStyle.None;//kenarlık belirtir
            DegisenPanel.Controls.Add(adm);//admin işlemleri formu açılır
            adm.Show();//admin işlemleri formu gösterilir
        }

        private void button5_Click(object sender, EventArgs e)
        {
            VeriMenusu veri = new VeriMenusu();
            DegisenPanel.Controls.Clear(); //veri işlemleri formu olsturulur
            veri.MdiParent = this;//veri altında çocuk form açar
            veri.FormBorderStyle = FormBorderStyle.None;//kenarlık belirtir
            DegisenPanel.Controls.Add(veri);//veri işlemleri formu açılır
            veri.Show();//veri işlemleri formu gösterilir
        }
    }
}
