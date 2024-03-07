using iText.Layout.Element;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.Windows.Forms;
using System.Diagnostics;
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
    public partial class OgrenciListesi : Form
    {
        SqlConnection baglanti = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true"); 
        SqlDataAdapter adapter;
        DataTable tablo;
        DataSet ds = new DataSet();

        public OgrenciListesi()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using(SaveFileDialog sfd = new SaveFileDialog() { Filter="Excel Sayfası| *.xlsx" })//kaydetme iletişim kutusunu açar
            {
                if (sfd.ShowDialog()==DialogResult.OK)//kullanıcı dosya seçer
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

        private void OgrenciListesi_Load(object sender, EventArgs e)
        {
            KisiGetir();
        }

        private void KisiGetir()
        {
            adapter = new SqlDataAdapter("SELECT * FROM Ogrenci", baglanti);
            tablo = new DataTable();
            BaglantiAc();//öğrencilerin bağlantısını açtık sql de
            adapter.Fill(tablo);//dolduruyoruz
            dataOgrenciler.DataSource = tablo;
            dataOgrenciler.ReadOnly = true;//hücrelerin düzenlenmesi engelleniyor
            baglanti.Close();//bağlantıyı kapattık
        }

        private void BaglantiAc()
        {//bağlantıyı açmak için yazdım
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
        } 
        private void dataOgrenciler_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
