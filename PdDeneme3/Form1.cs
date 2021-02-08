using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace PdDeneme3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string text;
        object font = new object();
        private void btn_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog()
            { ValidateNames = true, Multiselect = false, Filter = "Word 97-2003|*.doc|Word Document|*.docx" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    object readOnly = false;
                    object visible = true;
                    object save = false;
                    object fileName = ofd.FileName;
                    object newTemplate = false;
                    object docType = 0;
                    object missing = Type.Missing;
                    Microsoft.Office.Interop.Word._Document document;
                    Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application() { Visible = false };
                    document = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing
                        , ref missing, ref missing, ref missing, ref missing, ref missing, ref missing
                        , ref visible, ref missing, ref missing, ref missing, ref missing);
                    document.ActiveWindow.Selection.WholeStory();
                    document.ActiveWindow.Selection.Copy();
                    IDataObject dataObject = Clipboard.GetDataObject();
                    rtb.Rtf = dataObject.GetData(DataFormats.Rtf).ToString();
                    Microsoft.Office.Interop.Word.Range rng = document.Content;
                    text = rng.Text;
                    application.Quit(ref missing, ref missing, ref missing);
                }
            }
        }
        private void btnCheck_Click(object sender, EventArgs e)
        {
            string[] words = text.Split(' '); // Kelimeler boşluklarla birbirinden ayrılıp words dizisine aktarılmıştır.
            Boolean ciftTirnak()
            {
                for (int i = 0; i < words.Length; i++)
                {
                    int sayac = 0; //çift tirnak arası keliemelerin sayacı her tırnak başında sıfırlanır.
                    if (words[i].Contains("“")) // 
                    {
                        sayac++;
                        while (!words[i].Contains("”"))//kapanış çift tırnağına kadar aradaki keliemeleri sayar ve sayacı bir arttırır.
                        {
                            sayac++;
                            i++;
                        }
                        if (sayac > 50)  //tırnaklar arası kelimelerin sayısı 50'den büyükse false döndürür ve döngüden çıkar.
                        {
                            return false;
                            break;
                        }
                    }
                }
                return true; // eğer false döndürülmemişse true döndürülür
            }

            int kaynakBaslangic = 0; // kaynaklar listesinin başlangıç indisi için değişken oluşturulur.

            string[] kaynakListe()
            {
                for (int i = 0; i < words.Length; i++)
                {
                    if (words[i].Contains("KAYNAKLAR") && !words[i].Contains("SONUÇLAR")) // içindekiler kısmındaki KAYNAKLAR başlığını almaması için filtreleme yapılır.
                    {
                        kaynakBaslangic = i; // kaynaklar listesinin başlangıç indisi bulunduğunda değişkenimize atanır ve döngüden çıkılır.
                        break;
                    }
                }
                string wholeNumber = "";  // kaynakları depolamak için string değişken oluşturulr. **
                string[] kaynakNumberList; // kaynakların son kısımda birer birer alınması için string dizi oluşturulur.
                for (int i = kaynakBaslangic; i < words.Length; i++)
                {
                    char[] c = words[i].ToCharArray(); //her kelime char dizisine dönüştürülür.
                    for (int j = 0; j < c.Length; j++)
                    {
                        string number = ""; //sayılar char olarak alınacağı için birden fazla basamaklı olanları yan yana eklemek üzere string değişken oluşturulur ve her döngüde sıfırlanır.
                        if (c[j] == '[')   // sayılar köşeli parantezler içinde bulunmaktadır.
                        {
                            j++;
                            while (c[j] != ']')
                            {
                                number += c[j];  //rakamlar yanyana eklenerek sayı oluşturulur.
                                j++;
                            }
                            wholeNumber += number + " "; //döngüden çıkıldığında elimizdeki sayı depolayacağımız stringe eklenir ve boşluk bırakılır.
                            break;
                        }
                    }
                }
                kaynakNumberList = wholeNumber.Split(' '); // boşluklarla ayırdığımız sayılarımız split ile dizimize aktarılır.
                kaynakNumberList = kaynakNumberList.Take(kaynakNumberList.Length - 1).ToArray(); //son boşluk diziden çıkartılır
                return kaynakNumberList;
            }
            int girisBaslangic = 0; // girişin başlangıç indisi için değişken oluşturulur.
            string[] atifListe()
            {
                for (int i = 0; i < words.Length; i++)
                {
                    if (words[i].Contains("GİRİŞ") && !words[i].Contains("x")) // içindekiler kısmındaki GİRİŞ başlığını almaması için filtreleme yapılır.
                    {
                        girisBaslangic = i;
                        break;
                    }
                }
                string wholeNumber = "";
                string[] atifNumberList;
                for (int i = girisBaslangic; i < kaynakBaslangic; i++)
                {

                    char[] c = words[i].ToCharArray();
                    for (int j = 0; j < c.Length; j++)
                    {
                        string number = "";
                        if (c[j] == '[')
                        {
                            j++;
                            while (c[j] != ']')
                            {
                                if (c[j] == ',' || c[j] == '–') //burda kaynaklistesinde yaptığımız işlemden farklı olarak bu durum sağlandığında sayı bitmiş olur ve direkt eklenir.
                                {
                                    wholeNumber += number + " ";
                                    number = "";
                                }
                                else
                                {
                                    number += c[j];
                                }
                                j++;
                            }
                            wholeNumber += number + " ";
                            break;
                        }
                    }
                }
                atifNumberList = wholeNumber.Split(' ');
                atifNumberList = atifNumberList.Take(atifNumberList.Length - 1).ToArray();//son boşluk diziden çıkartılır
                atifNumberList = atifNumberList.Distinct().ToArray(); //tekrar eden öğeler silinir
                return atifNumberList;
            }
            string[] InsertionSort(string[] arr)
            {
                int[] intArr = new int[arr.Length];
                for (int i = 0; i < arr.Length; i++)
                {
                    intArr[i] = Int32.Parse(arr[i]);
                }
                for (int i = 0; i < intArr.Length - 1; i++)
                {
                    for (int j = i + 1; j > 0; j--)
                    {
                        if (intArr[j - 1] > intArr[j])
                        {
                            int temp = intArr[j - 1];
                            intArr[j - 1] = intArr[j];
                            intArr[j] = temp;
                        }
                    }
                }
                for (byte i = 0; i < arr.Length; i++)
                {
                    arr[i] = intArr[i].ToString();
                }
                return arr;
            }
            string[] kaynakNumaralar = kaynakListe(); 
            string[] atifNumaralar = atifListe();
            string[] atifSirali = InsertionSort(atifNumaralar); // daha kolay eşleşebilmesi için atıf yapılan numaralar sıralanır
            string kaynakKontrol(string[] atif, string[] kaynak)
            {
                bool atifTamMi = false; // buradaki değişken tüm kaynaklara atıf yapılıp yapılmadığını gösterir.
                for (int i = 0; i < kaynak.Length; i++)
                {
                    for (int j = 0; j < atif.Length; j++)
                    {
                        if (atif[j] == kaynak[i])
                        {
                            atifTamMi = true;
                            break;
                        }
                        else
                        {
                            atifTamMi = false;
                        }
                    }
                    if (!atifTamMi)
                    {
                        break;
                    }
                }
                bool kaynakTamMi = false; // buradaki değişken ise tanımlanmamış bir kaynağa atıf yapılıp yapılmadığını gösterir.
                for (int i = 0; i < atif.Length; i++)
                {
                    for (int j = 0; j < kaynak.Length; j++)
                    {
                        if (kaynak[j] == atif[i])
                        {
                            kaynakTamMi = true;
                            break;
                        }
                        else
                        {
                            kaynakTamMi = false;
                        }
                    }
                    if (!kaynakTamMi)
                    {
                        break;
                    }
                }
                if (kaynakTamMi && atifTamMi)
                {
                    return "BAŞARILI ! Tüm kaynaklara başarılı bir şekilde atıf yapılmıştır.";
                }
                else if (kaynakTamMi && !atifTamMi)
                {
                    return "BAŞARISIZ ! Kaynaklar kısmında, dökümanda kullanılmayan bir kaynak tanımlanmış.";
                }
                else if (!kaynakTamMi && atifTamMi)
                {
                    return "BAŞARISIZ ! Dökümanda, kaynaklar kısmında tanımlanmayan bir kaynak numarası kullanılmış.";
                }
                else
                {
                    return "BAŞARISIZ ! Kaynaklar kısmında tanımlanan kaynakların tamamına atıf yaptığınızdan emin olunuz.";
                }
            }
            int sekillerBaslangic = 0;
            void sekillerListe()
            {
                for (int i = 0; i < words.Length; i++)
                {
                    if (words[i].Contains("ŞEKİLLER"))
                    {
                        sekillerBaslangic = i;
                    }
                }
            }
            int tablolarBaslangic = 0;
            void tablolarListe()
            {
                for (int i = 0; i < words.Length; i++)
                {
                    if (words[i].Contains("TABLOLAR"))
                    {
                        tablolarBaslangic = i;
                    }
                }
            }
            string[] sekilListe()
            {
                int noktaSayisi = 0; // şekiller listesi tanımlanırken en fazla 2 nokta kullanılır.Onu saymak için değişken oluşturuldu.
                string tumSekiller = "";
                string sekilNo = "";
                for (int i = sekillerBaslangic; i < tablolarBaslangic; i++)
                {
                    if (words[i].Contains("Şekil"))
                    {
                        if (words[i + 1].Contains("."))
                        {
                            char[] c = words[i + 1].ToCharArray();
                            sekilNo = "";
                            noktaSayisi = 0;
                            for (int j = 0; j < c.Length; j++)
                            {
                                while (noktaSayisi < 2)  //eğer nokta sayısı 2 den küçükse sayımız henüz bitmemiştir.
                                {
                                    if (c[j] == '.')
                                    {
                                        noktaSayisi++; //eğer noktaya denk gelirsek nokta sayısını 1 arttırırız
                                    }
                                    sekilNo += c[j];
                                    j++;
                                }
                                tumSekiller += sekilNo + " ";
                                break;
                            }
                        }
                    }
                }
                string[] sekilNumberList = tumSekiller.Split(' ');
                sekilNumberList = sekilNumberList.Take(sekilNumberList.Length - 1).ToArray();
                return sekilNumberList;
            }
            string[] tabloListe()
            {
                int noktaSayisi = 0;
                string tumTablolar = "";
                string tabloNo = "";
                for (int i = tablolarBaslangic; i < girisBaslangic; i++)
                {
                    if (words[i].Contains("Tablo"))
                    {
                        if (words[i + 1].Contains("."))
                        {
                            char[] c = words[i + 1].ToCharArray();
                            tabloNo = "";
                            noktaSayisi = 0;
                            for (int j = 0; j < c.Length; j++)
                            {
                                while (noktaSayisi < 2)
                                {
                                    if (c[j] == '.')
                                    {
                                        noktaSayisi++;
                                    }
                                    tabloNo += c[j];
                                    j++;
                                }
                                tumTablolar += tabloNo + " ";
                                break;
                            }
                        }
                    }
                }
                string[] tabloNumberList = tumTablolar.Split(' ');
                tabloNumberList = tabloNumberList.Take(tabloNumberList.Length - 1).ToArray(); //son boşluk diziden çıkartılır
                return tabloNumberList;
            }
            string sekilAtifDene(string[] sekiller) //tüm şekillere atıf yapılmış mı
            {
                int hepsiMi = 0;
                for (int i = girisBaslangic; i < kaynakBaslangic; i++)
                {
                    for(int j = 0; j < sekiller.Length; j++) {
                        if (words[i].Contains("Şekil") && words[i + 1].Contains(sekiller[j]))
                        {
                            hepsiMi++;
                        }
                    }
                }
                if (hepsiMi >= sekiller.Length)
                {
                    return "BAŞARILI ! Şekiller listesinde tanımlanan tüm şekillere atıf yapılmıştır";
                }
                else
                {
                    return "BAŞARISIZ ! Şekiller listesinde tanımlanan tüm şekillere atıf yapılmamış";
                }
            }
            string tabloAtifDene(string[] tablolar) //tüm tablolara atıf yapılmış mı
            {
                int hepsiMi = 0;
                for (int i = girisBaslangic; i < kaynakBaslangic; i++)
                {
                    for (int j = 0; j < tablolar.Length; j++)
                    {
                        if (words[i].Contains("Tablo") && words[i + 1].Contains(tablolar[j]))
                        {
                            hepsiMi++;
                        }
                    }
                }
                if (hepsiMi >= tablolar.Length)
                {
                    return "BAŞARILI ! Tablolar listesinde tanımlanan tüm tablolara atıf yapılmıştır";
                }
                else
                {
                    return "BAŞARISIZ ! Tablolar listesinde tanımlanan tüm tablolara atıf yapılmamış";
                }
            }
            sekillerListe();
            tablolarListe();
            string[] sekilNumaralar = sekilListe();
            string[] tabloNumaralar = tabloListe();
            MessageBox.Show(kaynakKontrol(atifSirali, kaynakNumaralar));
            MessageBox.Show(sekilAtifDene(sekilNumaralar));
            MessageBox.Show(tabloAtifDene(tabloNumaralar));
            if (!ciftTirnak())
            {
                MessageBox.Show("BAŞARISIZ! Dökümanınızda iki tırnak arasında 50'den fazla kelime kullanmışsınız.");
            }
        }

    }
}
