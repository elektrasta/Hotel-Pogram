using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
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
    public partial class VeriMenusu : Form
    {
        string yol,expYol;
        string yolbackup, yolrestore;
        SqlCommand komut;
        SqlConnection dbConnection = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true");

        private void BaglantiAc()
        {//bağlantıyı açmak için yazdım
            if (dbConnection.State == ConnectionState.Closed)
            {
                dbConnection.Open();
            }
        }

        public VeriMenusu()
        {
            InitializeComponent();
        }

        public void btngozat_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";//ilk c de gözat açılır default olarak
            openFileDialog1.Filter = "CSV Dosyaları|*.csv";//csv dosyaları seçilmek için yazdım
            openFileDialog1.FilterIndex = 0;//default olarak 0 berlirttim
            openFileDialog1.RestoreDirectory = true;//kutu kapatıldığında orjinale döner

            if (openFileDialog1.ShowDialog() == DialogResult.OK)//kullanıcı dosya seçtiyse
            {
                yol = openFileDialog1.FileName;//yol değişkenliğine atanır
                inputDizini.Text = yol;//textbox a yazılır
                buttonImp.Enabled = true;//buton etkinleştirilir
            }
        }
         

        public void buttonImp_Click(object sender, EventArgs e)
        {
            TabledanSQLe(CSVdenTable(yol));
        }


        private void btnVTIGozat_Click(object sender, EventArgs e)
        {
           
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                getCSV();
                txtKontrol.Text = "Export tamamlandı.";//Text yerine export tamamlandı yazaılır
            }
            catch (Exception exp)
            {//hata olursa hata mesajı verir
                MessageBox.Show("Hata: " + exp.Message);
            }
        }

        private void btnVTIImport_Click(object sender, EventArgs e)
        {
            
        }

        private void btnExpGozat_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())//bellekten serbest bırakılır
            {
                if (fbd.ShowDialog() == DialogResult.OK)//kullanıcı klasör seçtiyse eğer
                {
                    tbExportPath.Text = fbd.SelectedPath + @"\exportOgrenci.csv";//csv uzantılı yapılır
                    expYol = tbExportPath.Text;//ylu  belirler
                    btnExport.Enabled = true;//verileri aktarır
                }
            }
        }

        private static DataTable CSVdenTable(string csv_file_path)
        {//csv dosya yolunu bulur ve geri döndürür
            DataTable csvData = new DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { ";" });//noktalı virgülle ayrılır
                    csvReader.HasFieldsEnclosedInQuotes = true;//kabul edilir
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datecolumn = new DataColumn(column);
                        datecolumn.AllowDBNull = true;
                        csvData.Columns.Add(datecolumn);
                    }//her satır dizi olarak elde alınır ve boş olup olmadığı kontrol ed,ilir ve boşsa null değeri atanır
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields(); 
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "")
                            {
                                fieldData[i] = null;
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                }
                
            }//hata oluşursa alt yer çalışır
            catch (Exception ex)
            {
            }
            return csvData;
        }
         
        public void TabledanSQLe(DataTable csvFileData)
        {//import edince sql e aktarır
            using (SqlConnection dbConnection = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true"))
            {//veri tabanı bağlantısı sağlanır
                dbConnection.Open();//bağlantı açılır
                using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))//hızlı şekilde aktarılır sql e
                {
                    s.DestinationTableName = "Ogrenci";//öğrenci tablosunu seçtik

                    foreach (var column in csvFileData.Columns)//sütun eşleşmesi yapılır
                        s.ColumnMappings.Add(column.ToString(), column.ToString()); //sql e aktarılır
                    s.WriteToServer(csvFileData);
                }
            } 
            txtKontrol.Text = "Import tamamlandı.";//texte yazılır
        }

        private string CSVOlustur(IDataReader reader)
        { 
            StreamWriter sw = new StreamWriter(expYol);//csv yolu oluşturulur
            object[] output = new object[reader.FieldCount]; //diziler oluşturulur
            try
            { 
                for (int i = 0; i < reader.FieldCount; i++)
                    output[i] = reader.GetName(i);//başlıklar csv dosya yoluna yazılır

                sw.WriteLine(string.Join(";", output));

                while (reader.Read())
                {//veriler csv dosyasına yazılır
                    reader.GetValues(output);
                    sw.WriteLine(string.Join(";", output));
                }
            }
            catch (Exception e)//hata durumunda alt yer çalışır
            {
                MessageBox.Show("Hata: " + e.Message);
            } 
            sw.Close();
            reader.Close();
            dbConnection.Close();
            return expYol;
        }
         
        private string getCSV()
        {
            using (dbConnection)//csv dosyasına dönüştürülür
            {
                BaglantiAc();//bağlantı açılır
                komut = new SqlCommand("select * from Ogrenci", dbConnection);//sql komutu sql e çalıştırılır
                return CSVOlustur(komut.ExecuteReader()); //ve yazar ve döndürülür return ile
            }
        }

        private void VeriMenusu_Load(object sender, EventArgs e){} 
        private void label5_Click(object sender, EventArgs e)
        { 
        }

        private void tbExportPath_TextChanged(object sender, EventArgs e)
        { 
        }

        private void inputDizini_TextChanged(object sender, EventArgs e)
        { 
        }

        private void label3_Click(object sender, EventArgs e)
        { 
        } 

        private void btnVTYGozat_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderbrowserdialog = new FolderBrowserDialog();

            if (folderbrowserdialog.ShowDialog() == DialogResult.OK)//kutu açılır ve seçim yapılırsa aşağısı çalışır
            {
                yolbackup = folderbrowserdialog.SelectedPath;//yolbackup değişkenine atanır
                tbVTI.Text = yolbackup;
                btnVTYedekle.Enabled = true;//görünür olur
            }
        }

        private void btnVTYedekle_Click(object sender, EventArgs e)
        {
            try
            { 
                if (tbVTI.Text != "")//boş olup lmadığı kontrol edilir
                {
                    DateTime now = DateTime.Now;  //şuanki saat ve tarih bilgisi alınır
                    //bağlantımızı yaptık
                    SqlConnection baglanti = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true");
                    baglanti.Open();//bağlantı açılır
                    //sql sorgusu çalıştırılır
                    string backupyol = @"BACKUP DATABASE [görsel3otomasyon] TO  DISK ='" + yolbackup + @"\görsel3otomasyon " + now.ToString("dd-MM-yyyy HH;mm;ss") + ".bak'";
                    komut = new SqlCommand(backupyol, baglanti);//sql komutu oluşturulur
                    komut.Connection = baglanti;
                    komut.ExecuteNonQuery();//yedekleme işlemi gerçekleştirilir
                    baglanti.Close();//bağlantı kapatılır
                    txtKontrol.Text = "Yedekleme Tamamlandı!";//rext e bu mesajı yazar
                }
                else
                {
                    MessageBox.Show("Yedekle alınacak klasörü seçin!");//hata mesajını gösterir
                }
            }
            catch (Exception ex)//extra bir hatada bunu verir
            {
                MessageBox.Show(ex.Message); 
            }
        }

        private void btnVTYDGozat_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();//iletişim kutusunu açar

            openFileDialog1.InitialDirectory = "c:\\";//default olarak c dizisini seçer
            openFileDialog1.FilterIndex = 0;//ilk değer olarak 0 ı default olarak alır
            openFileDialog1.RestoreDirectory = true;//seçilen yolun hatırlanması için

            if (openFileDialog1.ShowDialog() == DialogResult.OK)//eğer bir klsör seçildiyse
            {
                yolrestore = openFileDialog1.FileName;//değişken atanır
                tbVTE.Text = yolrestore;//seçilen dosya yolunu görüntüler
                btnVTYedektenDon.Enabled = true;//butonu görünür yapar

            }
        }

        private void inputDizini_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void btnVTYedektenDon_Click(object sender, EventArgs e)
        {
            try
            {
                if (tbVTE.Text != "")//boş olup olmadığı kontrol edilir
                {//sql bağlantısını yazdık
                    SqlConnection baglantirestore = new SqlConnection("Data Source=DESKTOP-KMRLDBJ\\SQLEXPRESS;initial catalog=görsel3otomasyon;integrated security=true");
                    //sorguyu yazdım
                    string restoreyol = @"alter database görsel3otomasyon set offline with rollback immediate " +" \n "+ @"RESTORE DATABASE [görsel3otomasyon] FROM  DISK  ='" + yolrestore + "' with replace" + "\n" + @"alter database görsel3otomasyon set online";
                    komut = new SqlCommand(restoreyol, baglantirestore);
                    baglantirestore.Open();//bağlantı açılır
                    komut.ExecuteNonQuery();//yedekten dönme işlemi gerçekleştirilir
                    baglantirestore.Close();//bağlantıdan çıkılır
                    txtKontrol.Text = "Yedekten dönme başarılı!";//texte  bu mesajı yazar başarılı oldu diye
                }
                else//boşsa aşağıdaki hatayı verir
                {
                    MessageBox.Show("Yedekten dönmek için dosya seçiniz!");
                }
            }
            catch (Exception ex)//genel bir hata varsa bunu verir
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtExportDizin_Click(object sender, EventArgs e){}
    } 
}
