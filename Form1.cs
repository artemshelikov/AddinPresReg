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
// Вставка библиотек
using Inventor;
using System.Diagnostics;
using System.Reflection;

namespace InvAddIn
{
    public partial class Form1 : Form
    {
        public string Версия_Inventor = "2022";
        /// <summary>
        /// ThisApplication - Объект для определения активного состояния Инвентора
        /// </summary>
        private Inventor.Application ThisApplication = null;
        /// <summary>
        /// Словарь для хранения ссылок на документы деталей
        /// </summary>
        private Dictionary<string, PartDocument> oPartDoc = new Dictionary<string,
        PartDocument>();
        /// <summary>
        /// Словарь для хранения ссылок на определения деталей
        /// </summary>
        private Dictionary<string, PartComponentDefinition> oCompDef = new Dictionary<string,
        PartComponentDefinition>();
        /// <summary>
        /// Словарь для хранения ссылок на инструменты создания деталей
        /// </summary>
        private Dictionary<string, TransientGeometry> oTransGeom = new Dictionary<string,
        TransientGeometry>();
        /// <summary>
        /// Словарь для хранения ссылок на транзакции редактирования
        /// </summary>
        private Dictionary<string, Transaction> oTrans = new Dictionary<string, Transaction>();
        /// <summary>
        /// Словарь для хранения имен сохраненных документов деталей
        /// </summary>
        private Dictionary<string, string> oFileName = new Dictionary<string, string>();
        private AssemblyDocument oAssemblyDocName;
        private AssemblyComponentDefinition oAssCompDef;
        // Необходим для поиска индивидуального кода плоскости детали
        static UserInputEvents UIevents;
        public Form1()
        {
            InitializeComponent();

            //Инициализация списков элементов
            comboBox1.Text = Convert.ToString(DGasPr);
            this.comboBox1.Items.AddRange(new object[] { "20,0", "18,0", "22,0", "24,0" });
            comboBox2.Text = Convert.ToString(DKlap);
            this.comboBox2.Items.AddRange(new object[] { "24,0", "25,0", "26,0", "27,5", "28,0" });
            comboBox3.Text = Convert.ToString(DKolca);
            this.comboBox3.Items.AddRange(new object[] { "76,0", "80,0", "84,0", "88,0" });
            comboBox4.Text = Convert.ToString(DFlancaVnesh);
            this.comboBox4.Items.AddRange(new object[] { "115,0", "116,0", "117,0", "118,0", "119,0", "120,0", "121,0", "122,0" });
            comboBox32.Text = Convert.ToString(DFlancaVnutr);
            this.comboBox32.Items.AddRange(new object[] { "90,0", "91,0", "92,0", "93,0", "94,0", "95,0", "96,0", "97,0" });
            comboBox5.Text = Convert.ToString(HKorp);
            this.comboBox5.Items.AddRange(new object[] { "75,0" });
            comboBox6.Text = Convert.ToString(DOsn);
            this.comboBox6.Items.AddRange(new object[] { "63,0", "64,0", "65,0", "66,0", "67,0", "68,0", "69,0", "70,0" });
            comboBox7.Text = Convert.ToString(RVtullka);
            this.comboBox7.Items.AddRange(new object[] { "7,5", "8,0", "8,5", "9,0", "9,5", "10,0" });
            comboBox8.Text = Convert.ToString(RVtullkaSmall);
            this.comboBox8.Items.AddRange(new object[] { "5,0", "5,5", "6,0", "6,5", "7,0", "7,5" });
            comboBox9.Text = Convert.ToString(HFlanca);
            this.comboBox9.Items.AddRange(new object[] { "12,0", "15,0" });
            comboBox10.Text = Convert.ToString(DOtvTr);
            this.comboBox10.Items.AddRange(new object[] { "33,23", "25,57", "21,224" });
            comboBox11.Text = Convert.ToString(DOtvBolt);
            this.comboBox11.Items.AddRange(new object[] { "8,0",  "10,0",  "12,0", "14,0" });
            comboBox12.Text = Convert.ToString(HFlancaKr);
            this.comboBox12.Items.AddRange(new object[] { "10,0", "11,0" });
            comboBox13.Text = Convert.ToString(DVint);
            this.comboBox13.Items.AddRange(new object[] { "16,0", "14,0", "18,0", "20,0"});
            comboBox14.Text = Convert.ToString(ShirStKr);
            this.comboBox14.Items.AddRange(new object[] { "5,0", "4,0", "6,0", "7,0" });
            comboBox15.Text = Convert.ToString(HKr);
            this.comboBox15.Items.AddRange(new object[] { "75,0" });
            comboBox16.Text = Convert.ToString(HGorlKr);
            this.comboBox16.Items.AddRange(new object[] { "11,0", "10,0", "12,0", "13,0" });
            comboBox17.Text = Convert.ToString(HVint);
            this.comboBox17.Items.AddRange(new object[] { "63,0", "57,0", "58,0", "59,0", "60,0", "61,0", "62,0", "64,0", "65,0" });
            comboBox18.Text = Convert.ToString(HRezbVint);
            this.comboBox18.Items.AddRange(new object[] { "37,0", "31,0", "32,0", "33,0", "34,0", "35,0", "36,0", "38,0", "39,0" });
            comboBox19.Text = Convert.ToString(HShtift);
            this.comboBox19.Items.AddRange(new object[] { "75,0", "60,0", "62,0", "65,0", "68,0", "70,0", "73,0", "78,0", "80,0" });
            comboBox20.Text = Convert.ToString(DShtift);
            this.comboBox20.Items.AddRange(new object[] { "8,0", "5,0", "6,0", "7,0", "9,0", "10,0", "11,0" });
            comboBox21.Text = Convert.ToString(HVtulka);
            this.comboBox21.Items.AddRange(new object[] { "10,0", "8,0", "9,0", "11,0", "12,0", "13,0" });
            comboBox23.Text = Convert.ToString(HYgol);
            this.comboBox23.Items.AddRange(new object[] { "22,0", "18,0", "20,0", "24,0" });
            comboBox22.Text = Convert.ToString(ShirYgol);
            this.comboBox22.Items.AddRange(new object[] { "18,0", "14,0", "16,0", "20,0", "22,0" });
            comboBox24.Text = Convert.ToString(LYgol);
            this.comboBox24.Items.AddRange(new object[] { "70,0", "60,0", "62,0", "63,0", "64,0", "65,0", "66,0", "67,0", "68,0", "69,0", "71,0", "72,0", "73,0" });
            comboBox25.Text = Convert.ToString(LTrYgol);
            this.comboBox25.Items.AddRange(new object[] { "52,0", "46,0", "47,0", "48,0", "49,0", "50,0", "51,0", "53,0", "54,0", "55,0", "56,0", "57,0", "58,0" });
            comboBox26.Text = Convert.ToString(DKolco);
            this.comboBox26.Items.AddRange(new object[] { "5,8", "6,0", "5,2", "5,4", "5,6", "6,2", "6,4", "6,6", "6,8" });
            comboBox27.Text = Convert.ToString(HProbka);
            this.comboBox27.Items.AddRange(new object[] { "20,0", "21,0", "22,0", "23,0", "24,0", "25,0", "26,0", "27,0" });
            comboBox28.Text = Convert.ToString(HRezbProbka);
            this.comboBox28.Items.AddRange(new object[] { "14,0", "15,0", "16,0", "17,0", "18,0", "19,0", "20,0", "21,0" });
            comboBox29.Text = Convert.ToString(LShtokRuchka);
            this.comboBox29.Items.AddRange(new object[] { "15,0", "14,0", "16,0", "17,0" });
            comboBox30.Text = Convert.ToString(LShtok);
            this.comboBox30.Items.AddRange(new object[] { "65,0", "60,0", "61,0", "62,0", "63,0", "64,0", "66,0", "67,0" });
            comboBox31.Text = Convert.ToString(DShtok);
            this.comboBox31.Items.AddRange(new object[] { "5,0", "4,0", "6,0", "7,0", "8,0" });
            comboBox38.Text = Convert.ToString(HPr6);
            this.comboBox38.Items.AddRange(new object[] { "45,0"});
            comboBox39.Text = Convert.ToString(DPr6);
            this.comboBox39.Items.AddRange(new object[] { "43,5", "40,5", "41,5", "42,5", "44,5", "45,5", "46,5" });
            comboBox40.Text = Convert.ToString(dPr6);
            this.comboBox40.Items.AddRange(new object[] { "7", "8", "9", "10", "11", "12" });
            comboBox41.Text = Convert.ToString(nPr6);
            this.comboBox41.Items.AddRange(new object[] { "4", "4,5", "5", "5,5"});
            comboBox45.Text = Convert.ToString(HPr7);
            this.comboBox45.Items.AddRange(new object[] { "42,0" });
            comboBox44.Text = Convert.ToString(DPr7);
            this.comboBox44.Items.AddRange(new object[] { "21,5", "18,5", "19,5", "20,5", "22,5", "23,5", "24,5" });
            comboBox43.Text = Convert.ToString(dPr7);
            this.comboBox43.Items.AddRange(new object[] { "3", "2", "4", "5"});
            comboBox42.Text = Convert.ToString(nPr7);
            this.comboBox42.Items.AddRange(new object[] { "7,5", "5,5", "6,0", "6,5", "7,0", "8,0" });
            comboBox49.Text = Convert.ToString(HPr18);
            this.comboBox49.Items.AddRange(new object[] { "19,0" });
            comboBox48.Text = Convert.ToString(DPr18);
            this.comboBox48.Items.AddRange(new object[] { "19,0", "16,0", "17,0", "18,0", "20,0"});
            comboBox47.Text = Convert.ToString(dPr18);
            this.comboBox47.Items.AddRange(new object[] { "2", "3", "4", "5" });
            comboBox46.Text = Convert.ToString(nPr18);
            this.comboBox46.Items.AddRange(new object[] { "6,0", "4,0", "4,5", "5,0", "5,5", "6,5" });
            comboBox34.Text = Convert.ToString(DFlancaVnesh);
            this.comboBox34.Items.AddRange(new object[] { "115,0", "116,0", "117,0", "118,0", "119,0", "120,0", "121,0", "122,0" });
            comboBox33.Text = Convert.ToString(DFlancaVnutr);
            this.comboBox33.Items.AddRange(new object[] { "90,0", "91,0", "92,0", "93,0", "94,0", "95,0", "96,0", "97,0" });
            comboBox35.Text = Convert.ToString(DOtvBolt);
            this.comboBox35.Items.AddRange(new object[] { "8,0", "9,0", "10,0", "11,0", "12,0", "14,0" });
            comboBox36.Text = Convert.ToString(DShtift);
            this.comboBox36.Items.AddRange(new object[] { "8,0", "5,0", "6,0", "7,0", "9,0", "10,0", "11,0" });
            comboBox37.Text = Convert.ToString(DShtok);
            this.comboBox37.Items.AddRange(new object[] { "5,0", "4,0", "6,0", "7,0", "8,0" });
            comboBox50.Text = Convert.ToString(DTar);
            this.comboBox50.Items.AddRange(new object[] { "58,0", "59,0", "60,0", "61,0", "62,0", "63,0"});
            comboBox51.Text = Convert.ToString(HTar);
            this.comboBox51.Items.AddRange(new object[] { "8,0", "7,0", "9,0", "10,0", "11,0", "12,0" });
            comboBox52.Text = Convert.ToString(hTar);
            this.comboBox52.Items.AddRange(new object[] { "5,0", "6,0", "7,0", "8,0", "9,0" });
            comboBox53.Text = Convert.ToString(d1Tar);
            this.comboBox53.Items.AddRange(new object[] { "50,0", "49,0", "48,0", "47,0", "46,0" });
            comboBox54.Text = Convert.ToString(d2Tar);
            this.comboBox54.Items.AddRange(new object[] { "22,0", "19,0", "20,0", "21,0", "23,0" });
            comboBox56.Text = Convert.ToString(DKolca17);
            this.comboBox56.Items.AddRange(new object[] { "50,0", "45,0", "47,0", "53,0", "55,0" });
            comboBox55.Text = Convert.ToString(dKolca17);
            this.comboBox55.Items.AddRange(new object[] { "4,1", "3,6", "3,8", "4,3" });
            comboBox57.Text = Convert.ToString(DYpor);
            this.comboBox57.Items.AddRange(new object[] { "40,0", "37,0", "38,0", "39,0", "41,0", "42,0", "43,0" });
            comboBox58.Text = Convert.ToString(HYpor);
            this.comboBox58.Items.AddRange(new object[] { "12,0", "9,0", "10,0", "11,0", "13,0", "14,0" });
            comboBox59.Text = Convert.ToString(d1Ypor);
            this.comboBox59.Items.AddRange(new object[] { "26,0", "22,0", "23,0", "24,0", "25,0", "27,0", "28,0" });
            comboBox60.Text = Convert.ToString(d2Ypor);
            this.comboBox60.Items.AddRange(new object[] { "22,0", "18,0", "19,0", "20,0", "21,0", "23,0", "24,0" });
            comboBox61.Text = Convert.ToString(h1Ypor);
            this.comboBox61.Items.AddRange(new object[] { "8,0", "9,0", "10,0", "11,0", "12,0" });
            comboBox62.Text = Convert.ToString(h2Ypor);
            this.comboBox62.Items.AddRange(new object[] { "5,0", "6,0", "7,0", "8,0", "9,0" });
            comboBox63.Text = Convert.ToString(DOpor);
            this.comboBox63.Items.AddRange(new object[] { "25,0", "20,0", "21,0", "22,0", "23,0", "24,0", "26,0", "27,0", "28,0", "29,0", "30,0" });
            comboBox64.Text = Convert.ToString(HOpor);
            this.comboBox64.Items.AddRange(new object[] { "8,0", "7,0", "9,0", "10,0", "11,0", "12,0", "13,0" });
            comboBox65.Text = Convert.ToString(d1Opor);
            this.comboBox65.Items.AddRange(new object[] { "12,0", "10,0", "11,0", "13,0", "14,0", "15,0", "16,0" });
            comboBox66.Text = Convert.ToString(h1Opor);
            this.comboBox66.Items.AddRange(new object[] { "2,0", "1,0", "3,0", "4,0", "5,0" });
            comboBox67.Text = Convert.ToString(h2Opor);
            this.comboBox67.Items.AddRange(new object[] { "3,5", "3,0", "4,0", "4,5", "5,0" });

            ThisApplication = (Inventor.Application)System.Runtime.InteropServices.
                Marshal.GetActiveObject("Inventor.Application");
            UIevents = ThisApplication.CommandManager.UserInputEvents;
            UIevents.OnSelect += click;
        }




        /// <summary>
        /// Функция создания имен новых элементов (документ, определение, инструменты,транзакции, имена файлов), необходимых для работы с деталями
        /// </summary>
        /// <param name="Name">Имя для ассоциативного обращения к элементам словарей</param>
        private void Имя_нового_документа(string Name)
        {
            // Новый документ детали
            oPartDoc[Name] = (PartDocument)ThisApplication.Documents.Add(
                DocumentTypeEnum.kPartDocumentObject, ThisApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
            // Новое определение
            oCompDef[Name] = oPartDoc[Name].ComponentDefinition;
            // Выбор инструментов
            oTransGeom[Name] = ThisApplication.TransientGeometry;
            // Создание транзакции
            oTrans[Name] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            // Имя файла
            oFileName[Name] = null;

        }

        private static double DGasPr = 20, DOsn = 63, DKlap = 48, RVtullka = 7.5,
            RVtullkaSmall = 5, DKolca = 80, DFlancaVnesh = 115, DFlancaVnutr = 90,
            HFlanca = 12, HKorp = 75, DOtvTr = 33.23, DOtvBolt = 8, HFlancaKr = 10,
            DVint = 16, ShirStKr = 5, HKr = 75, HGorlKr = 11, HVint = 63,
            HRezbVint = 37, DShtift = 8, HShtift = 75, HVtulka = 10, HYgol = 22,
            ShirYgol = 18, LYgol = 70, LTrYgol = 52, DKolco = 5.8, HProbka = 20,
            HRezbProbka = 14, DShtok = 5, LShtok = 65, LShtokRuchka = 15,
            HPr6 = 45, HPr7 = 42, HPr18 = 19, DPr6 = 43.5, DPr7 = 21.5, DPr18 = 19,
            dPr6 = 7, dPr7 = 3, dPr18 = 2, nPr6 = 4, nPr7 = 7.5, nPr18 = 6,
            DTar = 58, HTar = 8, hTar = 5, d1Tar = 50, d2Tar = 22, DKolca17 = 50,
            dKolca17 = 4.1, DYpor = 40, HYpor = 12, d1Ypor = 26, d2Ypor = 22, h1Ypor = 8,
            h2Ypor = 5, DOpor = 25 , HOpor = 8, d1Opor = 12, h1Opor= 2, h2Opor = 3.5;

        private void comboBox44_TextChanged(object sender, EventArgs e)
        {
            DPr7 = Convert.ToDouble(comboBox44.Text);
        }

        private void comboBox43_TextChanged(object sender, EventArgs e)
        {
            dPr7 = Convert.ToDouble(comboBox43.Text);
        }

        private void comboBox42_TextChanged(object sender, EventArgs e)
        {
            nPr7 = Convert.ToDouble(comboBox42.Text);
        }

        private void comboBox49_TextChanged(object sender, EventArgs e)
        {
            HPr18 = Convert.ToDouble(comboBox49.Text);
        }

        private void comboBox48_TextChanged(object sender, EventArgs e)
        {
            DPr18 = Convert.ToDouble(comboBox48.Text);
        }

        private void comboBox47_TextChanged(object sender, EventArgs e)
        {
            dPr18 = Convert.ToDouble(comboBox47.Text);
        }

        private void comboBox57_TextChanged(object sender, EventArgs e)
        {
            DYpor = Convert.ToDouble(comboBox57.Text);
            if (DYpor > DOsn*0.88-ShirStKr)
            {
                MessageBox.Show("Выберите значение меньше!");
            }
        }

        private void comboBox63_TextChanged(object sender, EventArgs e)
        {
            DOpor = Convert.ToDouble(comboBox63.Text);
        }

        private void comboBox64_TextChanged(object sender, EventArgs e)
        {
            HOpor = Convert.ToDouble(comboBox64.Text);
        }

        private void comboBox65_TextChanged(object sender, EventArgs e)
        {
            d1Opor = Convert.ToDouble(comboBox65.Text);
        }

        private void comboBox66_TextChanged(object sender, EventArgs e)
        {
            h1Opor = Convert.ToDouble(comboBox66.Text);
        }

        private void comboBox67_TextChanged(object sender, EventArgs e)
        {
            h2Opor = Convert.ToDouble(comboBox67.Text);
        }

        private void label70_Click(object sender, EventArgs e)
        {

        }

        private void comboBox58_TextChanged(object sender, EventArgs e)
        {
            HYpor = Convert.ToDouble(comboBox58.Text);
        }

        private void comboBox59_TextChanged(object sender, EventArgs e)
        {
            d1Ypor = Convert.ToDouble(comboBox59.Text);
            if (d1Ypor > DPr6-dPr6)
            {
                MessageBox.Show("Некорректное значение! Выберите значение меньше или измените параметры пружины 6!");
            }
        }

        private void comboBox60_TextChanged(object sender, EventArgs e)
        {
            d2Ypor = Convert.ToDouble(comboBox60.Text);
            if (d2Ypor < DPr7)
            {
                MessageBox.Show("Некорректное значение! Выберите значение ,больше или измените параметры пружины 7!");
            }
            if (d2Ypor > d1Ypor-4)
            {
                MessageBox.Show("Диаметр под пружину 7 не может быть больше диаметра под пружину 6!");
            }
        }

        private void comboBox61_TextChanged(object sender, EventArgs e)
        {
            h1Ypor = Convert.ToDouble(comboBox61.Text);
            if (h1Ypor > HYpor)
            {
                MessageBox.Show("Высота под пружину 6 не может быть больше высоты детали!");
            }
        }

        private void comboBox62_TextChanged(object sender, EventArgs e)
        {
            h2Ypor = Convert.ToDouble(comboBox62.Text);
            if (h2Ypor > HYpor)
            {
                MessageBox.Show("Высота под пружину 7 не может быть больше высоты детали!");
            }
            if (h2Ypor > h1Ypor)
            {
                MessageBox.Show("Высота под пружину 7 не может быть больше высоты под пружину 6!");
            }
        }

        private void comboBox46_TextChanged(object sender, EventArgs e)
        {
            nPr18 = Convert.ToDouble(comboBox46.Text);
        }

        private void comboBox50_TextChanged(object sender, EventArgs e)
        {
            DTar = Convert.ToDouble(comboBox50.Text);
        }

        private void comboBox51_TextChanged(object sender, EventArgs e)
        {
            HTar = Convert.ToDouble(comboBox51.Text);
        }

        private void comboBox52_TextChanged(object sender, EventArgs e)
        {
            hTar = Convert.ToDouble(comboBox52.Text);
            if (hTar==HTar||hTar>HTar) 
            { 
                MessageBox.Show("Это значение не может равняться или быть больше высоты"); 
            }
        }

        private void comboBox56_TextChanged(object sender, EventArgs e)
        {
            DKolca17 = Convert.ToDouble(comboBox56.Text);
        }

        private void comboBox55_TextChanged(object sender, EventArgs e)
        {
            dKolca17 = Convert.ToDouble(comboBox55.Text);
        }

        private void comboBox53_TextChanged(object sender, EventArgs e)
        {
            d1Tar = Convert.ToDouble(comboBox53.Text);
            if (d1Tar==DPr6||d1Tar<DPr6)
            {
                MessageBox.Show("Это значение не может равняться или быть меньше диаметра пружины 6");
            }
        }

        private void comboBox54_TextChanged(object sender, EventArgs e)
        {
            d2Tar = Convert.ToDouble(comboBox54.Text);
            if (d2Tar == DPr7 || d2Tar < DPr7)
            {
                MessageBox.Show("Это значение не может равняться или быть меньше диаметра пружины 7");
            }
            if (d2Tar/2*1.2 == DPr6/2-dPr6 || d1Tar/2*1.2 < DPr6/2-dPr6)
            {
                MessageBox.Show("Уменьшите это значение или выберите другие параметры пружины 6");
            }
            
        }

        private void comboBox34_TextChanged(object sender, EventArgs e)
        {
            if (comboBox34.Text != comboBox4.Text)
            {
                MessageBox.Show("Поменяйте соответствующее значение у детали Корпус");
            }
            DFlancaVnesh = Convert.ToDouble(comboBox34.Text);
            
        }

        private void comboBox33_TextChanged(object sender, EventArgs e)
        {
            DFlancaVnutr = Convert.ToDouble(comboBox33.Text);
            if (comboBox33.Text != comboBox32.Text)
            {
                MessageBox.Show("Поменяйте соответствующее значение у детали Корпус");
            }
            
        }

        private void comboBox35_TextChanged(object sender, EventArgs e)
        {
            DOtvBolt = Convert.ToDouble(comboBox35.Text);
            if (comboBox35.Text != comboBox11.Text)
            {
                MessageBox.Show("Поменяйте соответствующее значение у детали Корпус");
            }
        }

        private void comboBox36_TextChanged(object sender, EventArgs e)
        {
            DShtift = Convert.ToDouble(comboBox36.Text);
            if (comboBox36.Text != comboBox20.Text)
            {
                MessageBox.Show("Поменяйте соответствующее значение у детали Штифт");
            }
        }

        private void comboBox37_TextChanged(object sender, EventArgs e)
        {
            DShtok = Convert.ToDouble(comboBox37.Text);
            if (comboBox37.Text != comboBox31.Text)
            {
                MessageBox.Show("Поменяйте соответствующее значение у детали Шток");
            }
        }

        private void comboBox45_TextChanged(object sender, EventArgs e)
        {
            HPr7 = Convert.ToDouble(comboBox45.Text);
        }

        private void comboBox41_TextChanged(object sender, EventArgs e)
        {
            nPr6 = Convert.ToDouble(comboBox41.Text);
        }

        private void comboBox40_TextChanged(object sender, EventArgs e)
        {
            dPr6 =  Convert.ToDouble(comboBox40.Text);
        }

        private void comboBox39_TextChanged(object sender, EventArgs e)
        {
            DPr6 = Convert.ToDouble(comboBox39.Text);
        }

        private void comboBox38_TextChanged(object sender, EventArgs e)
        {
            HPr6 = Convert.ToDouble(comboBox38.Text);
        }

        private void comboBox15_TextChanged(object sender, EventArgs e)
        {
            HKr = Convert.ToDouble(comboBox15.Text);
        }

        private void comboBox16_TextChanged(object sender, EventArgs e)
        {
            HGorlKr = Convert.ToDouble(comboBox16.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show(FBD.SelectedPath);
            }
            AssemblyDocument oAssDoc = (AssemblyDocument)ThisApplication.Documents.Add(
                DocumentTypeEnum.kAssemblyDocumentObject, ThisApplication.FileManager.GetTemplateFile(
                    DocumentTypeEnum.kAssemblyDocumentObject));
            AssemblyComponentDefinition oAssCompDef = oAssDoc.ComponentDefinition;
            oAssDoc.DisplayName = "Регулятор давления";
            TransientGeometry aTransGeom = ThisApplication.TransientGeometry;
            Matrix oPositionMatrix = aTransGeom.CreateMatrix();
            //Переменные для выбранных поверхностей
            Face oFace1, oFace2, oFace3, oFace4, oFace5, oFace6, oFace7, oFace8, oFace9, oFace10, oFace11,
                oFace12, oFace13, oFace14, oFace15, oFace16, oFace17, oFace18, oFace19, oFace20, oFace21, 
                oFace22, oFace23, oFace24, oFace25, oFace26, oFace27, oFace28, oFace29, oFace30, oFace31,
                oFace32, oFace33, oFace34, oFace35, oFace36, oFace37, oFace38, oFace39, oFace40, oFace41, 
                oFace42, oFace43, oFace44, oFace45, oFace46, oFace47, oFace48, oFace49, oFace50, oFace51, 
                oFace52, oFace53, oFace54, oFace55, oFace56, oFace57, oFace58, oFace59, oFace60, oFace61,
                oFace62, oFace63, oFace64;
            //Переменные для сопряжений
            MateConstraint Поверхность1, Поверхность2, Поверхность3, Поверхность5, Поверхность6, 
                Поверхность7, Поверхность8, Поверхность9,Поверхность10, Поверхность11, Поверхность12, 
                Поверхность13, Поверхность14, Поверхность15, Поверхность16, Поверхность17, Поверхность18,
                Поверхность19, Поверхность20, Поверхность21, Поверхность22, Поверхность23, Поверхность24, 
                Поверхность25, Поверхность26, Поверхность28,Поверхность29,  Поверхность32, 
                Поверхность33, Поверхность34, Поверхность35, Поверхность36, Поверхность37, Поверхность38,
                Поверхность39;
            InsertConstraint Поверхность27, Поверхность30, Поверхность4, Поверхность31;

            //Вставка в сборку корпуса
            ComponentOccurrence Korpus_Model = oAssDoc.ComponentDefinition.
                Occurrences.Add(FBD.SelectedPath + "\\13. Корпус.ipt", oPositionMatrix);
            //Вставка в сборку пробки
            ComponentOccurrence Probka_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\20. Пробка.ipt", oPositionMatrix);
            //Сопряжение корпуса и пробки
            oFace1 = Korpus_Model.SurfaceBodies[1].Faces[74];
            oFace2 = Probka_Model.SurfaceBodies[1].Faces[27];
            oFace3 = Korpus_Model.SurfaceBodies[1].Faces[73];
            oFace4 = Probka_Model.SurfaceBodies[1].Faces[7];

            Поверхность1 = oAssCompDef.Constraints.AddMateConstraint(oFace1, oFace2, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            Поверхность2 = oAssCompDef.Constraints.AddMateConstraint(oFace3, oFace4, 0.4, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку кольца
            ComponentOccurrence Kolco17_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\17. Кольцо 55х48.ipt", oPositionMatrix);
            //Вставка в сборку клапана
            ComponentOccurrence Klapan_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\19. Клапан.ipt", oPositionMatrix);
            //Вставка в сборку Крышки
            ComponentOccurrence Krishka_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\8. Крышка.ipt", oPositionMatrix);
            //Сопряжение корпуса и крышки
            oFace27 = Korpus_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Korpus_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Korpus_Model.SurfaceBodies[1].Faces[i].
                    InternalName == "{AB5AE357-3018-9014-F686-2AB10D12C087}")
                {
                    oFace27 = Korpus_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace28 = Krishka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Krishka_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Krishka_Model.SurfaceBodies[1].Faces[i].
                    InternalName == "{25F2A212-7BDB-168B-0BA7-4139FA20B9C4}")
                {
                    oFace28 = Krishka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность15 = oAssCompDef.Constraints.AddMateConstraint(oFace27, oFace28, 0, 
                InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace11 = Korpus_Model.SurfaceBodies[1].Faces[1];
            oFace12 = Krishka_Model.SurfaceBodies[1].Faces[27];
            Поверхность6 = oAssCompDef.Constraints.AddMateConstraint(oFace11, oFace12, HFlanca / 10, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Сопряжение корпуса и кольца 17
            oFace34 = Kolco17_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Kolco17_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Kolco17_Model.SurfaceBodies[1].Faces[i].InternalName == "{381EFEE0-3770-60E0-8E21-BBA3E30E5064}")
                {
                    oFace34 = Kolco17_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность21 = oAssCompDef.Constraints.AddMateConstraint(oFace27, oFace34, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace5 = Korpus_Model.SurfaceBodies[1].Faces[73];
            oFace6 = Kolco17_Model.SurfaceBodies[1].Faces[1];
            Поверхность3 = oAssCompDef.Constraints.AddMateConstraint(oFace5, oFace6, 0.205, 
                InferredTypeEnum.kNoInference, InferredTypeEnum.kInferredPoint);
            //Сопряжение корпуса и клапана
            oFace7 = Klapan_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Klapan_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Klapan_Model.SurfaceBodies[1].Faces[i].InternalName == "{9DEEDF33-408A-612F-BBB5-1BF39E2C23D3}")
                {
                    oFace7 = Klapan_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность22 = oAssCompDef.Constraints.AddMateConstraint(oFace27, oFace7, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace9 = Korpus_Model.SurfaceBodies[1].Faces[50];
            oFace10 = Klapan_Model.SurfaceBodies[1].Faces[8];
            Поверхность5 = oAssCompDef.Constraints.AddMateConstraint(oFace9, oFace10, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку диафрагмы
            ComponentOccurrence Diafragma_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\11. Диафрагма.ipt", oPositionMatrix);
            //Сопряжение крышки и диафрагмы
            oFace29 = Diafragma_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Diafragma_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Diafragma_Model.SurfaceBodies[1].Faces[i].InternalName == "{A0662467-0A84-47FC-0240-B05415B700AF}")
                {
                    oFace29 = Diafragma_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность16 = oAssCompDef.Constraints.AddMateConstraint(oFace28, oFace29, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace14 = Diafragma_Model.SurfaceBodies[1].Faces[3];
            Поверхность7 = oAssCompDef.Constraints.AddMateConstraint(oFace12, oFace14, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку тарелки
            ComponentOccurrence Tarelka_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\10. Тарелка.ipt", oPositionMatrix);
            //Сопряжение диафрагмы и тарелки
            oFace30 = Tarelka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Tarelka_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Tarelka_Model.SurfaceBodies[1].Faces[i].InternalName == "{C414203D-913C-B3B7-4EB5-5AD7DE07B8D6}")
                {
                    oFace30 = Tarelka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность17 = oAssCompDef.Constraints.AddMateConstraint(oFace29, oFace30, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace13 = Tarelka_Model.SurfaceBodies[1].Faces[13];
            Поверхность8 = oAssCompDef.Constraints.AddMateConstraint(oFace14, oFace13, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку опоры
            ComponentOccurrence Opora_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\1. Опора.ipt", oPositionMatrix);
            //Сопряжение диафрагмы и опоры
            oFace31 = Opora_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Opora_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Opora_Model.SurfaceBodies[1].Faces[i].InternalName == "{72F96A85-9247-18C0-FAE5-A2A009A1AEFA}")
                {
                    oFace31 = Opora_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность18 = oAssCompDef.Constraints.AddMateConstraint(oFace29, oFace31, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace15 = Diafragma_Model.SurfaceBodies[1].Faces[2];
            oFace16 = Opora_Model.SurfaceBodies[1].Faces[9];
            Поверхность9 = oAssCompDef.Constraints.AddMateConstraint(oFace15, oFace16, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //вставка в сборку упора
            ComponentOccurrence Ypor_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\5. Упор.ipt", oPositionMatrix);
            //Сопряжение крышки и упора
            oFace32 = Ypor_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Ypor_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Ypor_Model.SurfaceBodies[1].Faces[i].InternalName == "{FD282878-B904-8B6B-92A6-06ED9B28D097}")
                {
                    oFace32 = Ypor_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность19 = oAssCompDef.Constraints.AddMateConstraint(oFace29, oFace32, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace17 = Krishka_Model.SurfaceBodies[1].Faces[35];
            oFace18 = Ypor_Model.SurfaceBodies[1].Faces[9];
            Поверхность10 = oAssCompDef.Constraints.AddMateConstraint(oFace17, oFace18, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку втулки
            ComponentOccurrence Vtylka_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\14. Втулка.ipt", oPositionMatrix);
            //Сопряжение корпуса и втулки
            oFace33 = Vtylka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Vtylka_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Vtylka_Model.SurfaceBodies[1].Faces[i].InternalName == "{F8D804EE-D843-9D14-F9E4-F2E26F897B19}")
                {
                    oFace33 = Vtylka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность20 = oAssCompDef.Constraints.AddMateConstraint(oFace27, oFace33, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace19 = Korpus_Model.SurfaceBodies[1].Faces[56];
            oFace20 = Vtylka_Model.SurfaceBodies[1].Faces[11];
            Поверхность11 = oAssCompDef.Constraints.AddMateConstraint(oFace19, oFace20, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку прокладки
            ComponentOccurrence Prokladka_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\16. Прокладка.ipt", oPositionMatrix);
            //Сопряжение клапана и прокладки
            oFace8 = Prokladka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Kolco17_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Prokladka_Model.SurfaceBodies[1].Faces[i].InternalName == "{A48345A5-C8EE-3683-99B2-5471FD271833}")
                {
                    oFace8 = Prokladka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность23 = oAssCompDef.Constraints.AddMateConstraint(oFace27, oFace8, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace21 = Klapan_Model.SurfaceBodies[1].Faces[11];
            oFace22 = Prokladka_Model.SurfaceBodies[1].Faces[2];
            Поверхность12 = oAssCompDef.Constraints.AddMateConstraint(oFace21, oFace22, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку Шайбы
            ComponentOccurrence Shaiba_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\21. Шайба.ipt", oPositionMatrix);
            //Сопряжение прокладки и шайбы
            oFace35 = Shaiba_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Kolco17_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Shaiba_Model.SurfaceBodies[1].Faces[i].InternalName == "{0D5600F2-033F-3817-FECE-24A8B9801802}")
                {
                    oFace35 = Shaiba_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность24 = oAssCompDef.Constraints.AddMateConstraint(oFace8, oFace35, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace23 = Prokladka_Model.SurfaceBodies[1].Faces[4];
            oFace24 = Shaiba_Model.SurfaceBodies[1].Faces[2];
            Поверхность13 = oAssCompDef.Constraints.AddMateConstraint(oFace23, oFace24, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку Штока
            ComponentOccurrence Shtok_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\15. Шток.ipt", oPositionMatrix);
            //Сопряжение прокладки и штока
            oFace36 = Shtok_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Shtok_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Shtok_Model.SurfaceBodies[1].Faces[i].InternalName == "{494601EB-D8A3-6367-F92E-21C9B03B2E86}")
                {
                    oFace36 = Shtok_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность25 = oAssCompDef.Constraints.AddMateConstraint(oFace8, oFace36, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace25 = Shaiba_Model.SurfaceBodies[1].Faces[4];
            oFace26 = Shtok_Model.SurfaceBodies[1].Faces[6];
            Поверхность14 = oAssCompDef.Constraints.AddMateConstraint(oFace25, oFace26, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку Винта
            ComponentOccurrence Vint_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\3. Винт.ipt", oPositionMatrix);
            //Сопряжение корпуса и винта

            oFace37 = Vint_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Vint_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Vint_Model.SurfaceBodies[1].Faces[i].InternalName == "{AAE9C4AF-0487-F0BA-E4C4-371F535A3B4E}")
                {
                    oFace37 = Vint_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность26 = oAssCompDef.Constraints.AddMateConstraint(oFace27, oFace37, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace38 = Krishka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Krishka_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Krishka_Model.SurfaceBodies[1].Faces[i].InternalName == "{2E716E67-D9C4-A22F-9519-083F5E7AA48B}")
                {
                    oFace38 = Krishka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace39 = Vint_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Vint_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Vint_Model.SurfaceBodies[1].Faces[i].InternalName == "{99F2EE8C-853D-EA0D-A959-7565D1DAAAA9}")
                {
                    oFace39 = Vint_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность27 = oAssCompDef.Constraints.AddInsertConstraint(oFace38, oFace39, false, -HGorlKr / 10);
            //Вставка в сборку Штифта
            ComponentOccurrence Shtift_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\4. Штифт.ipt", oPositionMatrix);
            //Сопряжение винта и штифта
            oFace40 = Vint_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Vint_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Vint_Model.SurfaceBodies[1].Faces[i].InternalName == "{9950862B-8E71-823C-C713-4225E1778E4E}")
                {
                    oFace40 = Vint_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace41 = Shtift_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Shtift_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Shtift_Model.SurfaceBodies[1].Faces[i].InternalName == "{41423855-E87E-EAB9-47A1-0DD7853E1DBA}")
                {
                    oFace41 = Shtift_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace42 = Shtift_Model.SurfaceBodies[1].Faces[3];
            Поверхность28 = oAssCompDef.Constraints.AddMateConstraint(oFace40, oFace41, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            Поверхность29 = oAssCompDef.Constraints.AddMateConstraint(oFace39, oFace42, -HShtift / 2 / 10 + DShtift / 10, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredPoint);
            //Вставка в сборку угольника
            ComponentOccurrence Ygolnik_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\22. Угольник.ipt", oPositionMatrix);
            //Сопряжение Корпуса и угольника
            oFace42 = Korpus_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Korpus_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Korpus_Model.SurfaceBodies[1].Faces[i].InternalName == "{9B5A1372-8578-7139-E2AE-9F8AEB37FDFC}")
                {
                    oFace42 = Korpus_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace43 = Ygolnik_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Ygolnik_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Ygolnik_Model.SurfaceBodies[1].Faces[i].InternalName == "{0C997AA2-F05D-B378-19B3-BA2FF624F5F1}")
                {
                    oFace43 = Ygolnik_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность30 = oAssCompDef.Constraints.AddInsertConstraint(oFace42, oFace43, true, 1.6);
            //Вставка в сборку кольца
            ComponentOccurrence Kolco12_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\12. Кольцо 80х70.ipt", oPositionMatrix);
            //Сопряжение корпуса и кольца 12
            oFace44 = Kolco12_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Kolco12_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Kolco12_Model.SurfaceBodies[1].Faces[i].InternalName == "{381EFEE0-3770-60E0-8E21-BBA3E30E5064}")
                {
                    oFace44 = Kolco12_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность21 = oAssCompDef.Constraints.AddMateConstraint(oFace27, oFace44, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            oFace45 = Korpus_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Korpus_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Korpus_Model.SurfaceBodies[1].Faces[i].InternalName == "{A67D5A84-F0A6-9169-5C9C-363F9AE9EAE6}")
                {
                    oFace45 = Korpus_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace6 = Kolco12_Model.SurfaceBodies[1].Faces[1];
            Поверхность3 = oAssCompDef.Constraints.AddMateConstraint(oFace45, oFace6, 0.29, InferredTypeEnum.kNoInference, InferredTypeEnum.kInferredPoint);
            //Вставка в сборку Пробки 23
            ComponentOccurrence Probka23_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\23. Пробка.ipt", oPositionMatrix);
            oFace46 = Korpus_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Korpus_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Korpus_Model.SurfaceBodies[1].Faces[i].InternalName == "{60210093-1163-24A6-3093-0805EA6E4C1A}")
                {
                    oFace46 = Korpus_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace47 = Probka23_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Probka23_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Probka23_Model.SurfaceBodies[1].Faces[i].InternalName == "{1D42DAB9-89FB-7CCD-C4AD-CD708B5EA36D}")
                {
                    oFace47 = Probka23_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность4 = oAssCompDef.Constraints.AddInsertConstraint(oFace46, oFace47, true, -1.6);
            //Вставка в сборку Пружины 6
            ComponentOccurrence Pruzhina6_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\6. Пружина.ipt", oPositionMatrix);
            oFace50 = Pruzhina6_Model.SurfaceBodies[1].Edges[1] as Face;
            for (int i = 1; i <= Pruzhina6_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Pruzhina6_Model.SurfaceBodies[1].Faces[i].InternalName == "{A39DE579-ADB4-A593-26FB-63A3B0C53FEF}")
                {
                    oFace50 = Pruzhina6_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace51 = Tarelka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Tarelka_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Tarelka_Model.SurfaceBodies[1].Faces[i].InternalName == "{C414203D-913C-B3B7-4EB5-5AD7DE07B8D6}")
                {
                    oFace51 = Tarelka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }

            oFace48 = Tarelka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Tarelka_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Tarelka_Model.SurfaceBodies[1].Faces[i].InternalName == "{4620FA6E-87F2-1AB1-D5F3-907185DFCFB9}")
                {
                    oFace48 = Tarelka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace49 = Pruzhina6_Model.SurfaceBodies[1].Edges[1] as Face;
            for (int i = 1; i <= Pruzhina6_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Pruzhina6_Model.SurfaceBodies[1].Faces[i].InternalName == "{E4CC3656-EC80-67E1-2B8F-A0A3AB67ADF8}")
                {
                    oFace49 = Pruzhina6_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность33 = oAssCompDef.Constraints.AddMateConstraint(oFace51, oFace50, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            Поверхность32 = oAssCompDef.Constraints.AddMateConstraint(oFace48, oFace49, -0.01, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку Пружины 7
            ComponentOccurrence Pruzhina7_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\7. Пружина.ipt", oPositionMatrix);
            oFace52 = Pruzhina7_Model.SurfaceBodies[1].Edges[1] as Face;
            for (int i = 1; i <= Pruzhina7_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Pruzhina7_Model.SurfaceBodies[1].Faces[i].InternalName == "{A39DE579-ADB4-A593-26FB-63A3B0C53FEF}")
                {
                    oFace52 = Pruzhina7_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace54 = Tarelka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Tarelka_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Tarelka_Model.SurfaceBodies[1].Faces[i].InternalName == "{F7663AD1-2CC3-F17C-64F1-5A37D1D841A7}")
                {
                    oFace54 = Tarelka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace55 = Pruzhina7_Model.SurfaceBodies[1].Edges[1] as Face;
            for (int i = 1; i <= Pruzhina7_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Pruzhina7_Model.SurfaceBodies[1].Faces[i].InternalName == "{E4CC3656-EC80-67E1-2B8F-A0A3AB67ADF8}")
                {
                    oFace55 = Pruzhina7_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность34 = oAssCompDef.Constraints.AddMateConstraint(oFace51, oFace52, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            Поверхность35 = oAssCompDef.Constraints.AddMateConstraint(oFace54, oFace55, -0.01, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
            //Вставка в сборку Пружины 18
            ComponentOccurrence Pruzhina18_Model = oAssDoc.ComponentDefinition.Occurrences.Add(FBD.SelectedPath + "\\18. Пружина.ipt", oPositionMatrix);
            oFace53 = Pruzhina18_Model.SurfaceBodies[1].Edges[1] as Face;
            for (int i = 1; i <= Pruzhina18_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Pruzhina18_Model.SurfaceBodies[1].Faces[i].InternalName == "{A39DE579-ADB4-A593-26FB-63A3B0C53FEF}")
                {
                    oFace53 = Pruzhina18_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace56 = Probka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Probka_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Probka_Model.SurfaceBodies[1].Faces[i].InternalName == "{4B014BA6-8AA6-FE46-D65B-301A3B8EE385}")
                {
                    oFace56 = Probka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }

            oFace57 = Probka_Model.SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= Probka_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Probka_Model.SurfaceBodies[1].Faces[i].InternalName == "{4D9DED6C-0B2F-13C3-7203-A670AB2CF507}")
                {
                    oFace57 = Probka_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            oFace58 = Pruzhina18_Model.SurfaceBodies[1].Edges[1] as Face;
            for (int i = 1; i <= Pruzhina18_Model.SurfaceBodies[1].Faces.Count; i++)
            {
                if (Pruzhina18_Model.SurfaceBodies[1].Faces[i].InternalName == "{E4CC3656-EC80-67E1-2B8F-A0A3AB67ADF8}")
                {
                    oFace58 = Pruzhina18_Model.SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Поверхность36 = oAssCompDef.Constraints.AddMateConstraint(oFace56, oFace53, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
            Поверхность37 = oAssCompDef.Constraints.AddMateConstraint(oFace57, oFace58, -0.01, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);


            Inventor.ContentCenter oContentCenter = ThisApplication.ContentCenter; //получаем ссылку на ContentCenter
            Inventor.ContentTreeViewNode oContentNode = default;
            Inventor.ContentTreeViewNode oContentNode1 = default;
            oContentNode1 = oContentCenter.TreeViewTopNode.ChildNodes["Крепежные изделия"].ChildNodes["Болты  Винты"].ChildNodes["С углублением под ключ"];

            long AllPoint1, NowPoint1;    //этими переменными реализован счётчик количества строк в библиотеке
            NowPoint1 = 0;

            foreach (Inventor.ContentFamily oFamily in oContentNode1.Families)//тут бегаем по семействам в папке/подпапке узла
            {
                AllPoint1 = oFamily.TableRows.Count;
                if (oFamily.DisplayName == "Винт ГОСТ 11738-84")
                {
                    foreach (Inventor.ContentTableRow ContTableRow in oFamily.TableRows)//тут бегаем по строкам таблицы
                    {
                        NowPoint1 += 1;
                        if (DOtvBolt == 8)
                        {
                            if (NowPoint1 == 50)
                            {
                                
                                var ContTableRow_Designation = ContTableRow.GetCellValue("DESIGNATION");
                                Inventor.MemberManagerErrorsEnum failureCM = default;
                                string MessageProbl = "You_have_a_Problem";
                                var CreateMem1 = oFamily.CreateMember(ContTableRow, out failureCM, out MessageProbl);
                                ComponentOccurrence Vint = oAssDoc.ComponentDefinition.Occurrences.Add(CreateMem1, oPositionMatrix);//создание вхождения

                                oFace63 = Krishka_Model.SurfaceBodies[1].Faces[1];
                                for (int i = 1; i <= Krishka_Model.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Krishka_Model.SurfaceBodies[1].Faces[i].InternalName == "{F8100852-5D1F-2F02-EE30-478330D05BB0}")
                                    {
                                        oFace63 = Krishka_Model.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace64 = Vint.SurfaceBodies[1].Faces[1];
                                for (int i = 1; i <= Vint.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Vint.SurfaceBodies[1].Faces[i].InternalName == "{C330C981-36E8-CEDA-7982-6347B0894ECF}")
                                    {
                                        oFace64 = Vint.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                Поверхность31 = oAssCompDef.Constraints.AddInsertConstraint(oFace64, oFace63, true, 0);

                                ObjectCollection objCollectionVint = ThisApplication.TransientObjects.CreateObjectCollection();
                                objCollectionVint.Add(Vint);
                                CircularOccurrencePattern CircPat;
                                CircPat = oAssCompDef.OccurrencePatterns.AddCircularPattern(objCollectionVint, oAssCompDef.WorkAxes[2], true, 60+"degree", 6);
                            }
                        }
                        if (DOtvBolt == 10)
                        {
                            if (NowPoint1 == 80)
                            {
                                
                                var ContTableRow_Designation = ContTableRow.GetCellValue("DESIGNATION");
                                Inventor.MemberManagerErrorsEnum failureCM = default;
                                string MessageProbl = "You_have_a_Problem";
                                var CreateMem1 = oFamily.CreateMember(ContTableRow, out failureCM, out MessageProbl);
                                ComponentOccurrence Vint = oAssDoc.ComponentDefinition.Occurrences.Add(CreateMem1, oPositionMatrix);//создание вхождения
                                oFace63 = Krishka_Model.SurfaceBodies[1].Faces[1];
                                for (int i = 1; i <= Krishka_Model.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Krishka_Model.SurfaceBodies[1].Faces[i].InternalName == "{F8100852-5D1F-2F02-EE30-478330D05BB0}")
                                    {
                                        oFace63 = Krishka_Model.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace64 = Vint.SurfaceBodies[1].Faces[1];
                                for (int i = 1; i <= Vint.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Vint.SurfaceBodies[1].Faces[i].InternalName == "{C330C981-36E8-CEDA-7982-6347B0894ECF}")
                                    {
                                        oFace64 = Vint.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                Поверхность31 = oAssCompDef.Constraints.AddInsertConstraint(oFace64, oFace63, true, 0);

                                ObjectCollection objCollectionVint = ThisApplication.TransientObjects.CreateObjectCollection();
                                objCollectionVint.Add(Vint);
                                CircularOccurrencePattern CircPat;
                                CircPat = oAssCompDef.OccurrencePatterns.AddCircularPattern(objCollectionVint, oAssCompDef.WorkAxes[2], true, 60 + "degree", 6);
                            }
                        }
                        if (DOtvBolt == 12)
                        {
                            if (NowPoint1 == 111)
                            {
                                
                                var ContTableRow_Designation = ContTableRow.GetCellValue("DESIGNATION");
                                Inventor.MemberManagerErrorsEnum failureCM = default;
                                string MessageProbl = "You_have_a_Problem";
                                var CreateMem1 = oFamily.CreateMember(ContTableRow, out failureCM, out MessageProbl);
                                ComponentOccurrence Vint = oAssDoc.ComponentDefinition.Occurrences.Add(CreateMem1, oPositionMatrix);//создание вхождения

                            }
                        }
                        if (DOtvBolt == 14)
                        {
                            if (NowPoint1 == 144)
                            {
                                
                                var ContTableRow_Designation = ContTableRow.GetCellValue("DESIGNATION");
                                Inventor.MemberManagerErrorsEnum failureCM = default;
                                string MessageProbl = "You_have_a_Problem";
                                var CreateMem1 = oFamily.CreateMember(ContTableRow, out failureCM, out MessageProbl);
                                ComponentOccurrence Vint = oAssDoc.ComponentDefinition.Occurrences.Add(CreateMem1, oPositionMatrix);//создание вхождения
                                oFace63 = Krishka_Model.SurfaceBodies[1].Faces[1];
                                for (int i = 1; i <= Krishka_Model.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Krishka_Model.SurfaceBodies[1].Faces[i].InternalName == "{F8100852-5D1F-2F02-EE30-478330D05BB0}")
                                    {
                                        oFace63 = Krishka_Model.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace64 = Vint.SurfaceBodies[1].Faces[1];
                                for (int i = 1; i <= Vint.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Vint.SurfaceBodies[1].Faces[i].InternalName == "{C330C981-36E8-CEDA-7982-6347B0894ECF}")
                                    {
                                        oFace64 = Vint.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                Поверхность31 = oAssCompDef.Constraints.AddInsertConstraint(oFace64, oFace63, true, 0);

                                ObjectCollection objCollectionVint = ThisApplication.TransientObjects.CreateObjectCollection();
                                objCollectionVint.Add(Vint);
                                CircularOccurrencePattern CircPat;
                                CircPat = oAssCompDef.OccurrencePatterns.AddCircularPattern(objCollectionVint, oAssCompDef.WorkAxes[2], true, 60 + "degree", 6);
                            }
                        }
                    }
                }
            }

            oContentNode = oContentCenter.TreeViewTopNode.ChildNodes["Крепежные изделия"].ChildNodes["Гайки"].ChildNodes["Шестигранные"];

            //ComponentOccurrence Gaika = oAssDoc.ComponentDefinition.Occurrences.Add(oContentNode, oPositionMatrix);

            long AllPoint, NowPoint;    //этими переменными реализован счётчик количества строк в библиотеке
            NowPoint = 0;

            foreach (Inventor.ContentFamily oFamily in oContentNode.Families)//тут бегаем по семействам в папке/подпапке узла
            {
                
                if (oFamily.DisplayName == "Гайка ГОСТ 2526-70")
                {
                    foreach (Inventor.ContentTableRow ContTableRow in oFamily.TableRows)//тут бегаем по строкам таблицы
                    {
                        NowPoint += 1;
                        if (DVint == 16) { 
                        if (NowPoint == 10) { 
                        var ContTableRow_Designation = ContTableRow.GetCellValue("DESIGNATION");
                        Inventor.MemberManagerErrorsEnum failureCM = default;
                        string MessageProbl = "You_have_a_Problem";
                        var CreateMem = oFamily.CreateMember(ContTableRow, out failureCM, out MessageProbl);
                        ComponentOccurrence Gaika = oAssDoc.ComponentDefinition.Occurrences.Add(CreateMem, oPositionMatrix);//создание вхождения

                            oFace59 = Krishka_Model.SurfaceBodies[1].Edges[1] as Face;
                            for (int i = 1; i <= Krishka_Model.SurfaceBodies[1].Faces.Count; i++)
                            {
                                if (Krishka_Model.SurfaceBodies[1].Faces[i].InternalName == "{2E716E67-D9C4-A22F-9519-083F5E7AA48B}")
                                {
                                    oFace59 = Krishka_Model.SurfaceBodies[1].Faces[i] as Face;
                                    break;
                                }
                            }
                            oFace60 = Gaika.SurfaceBodies[1].Edges[1] as Face;
                            for (int i = 1; i <= Gaika.SurfaceBodies[1].Faces.Count; i++)
                            {
                                if (Gaika.SurfaceBodies[1].Faces[i].InternalName == "{0F6E14BD-07B3-0DAA-966D-7B80E83827B8}")
                                {
                                    oFace60 = Gaika.SurfaceBodies[1].Faces[i] as Face;
                                    break;
                                }
                            }
                            oFace61 = Gaika.SurfaceBodies[1].Edges[1] as Face;
                            for (int i = 1; i <= Gaika.SurfaceBodies[1].Faces.Count; i++)
                            {
                                if (Gaika.SurfaceBodies[1].Faces[i].InternalName == "{C5765F4A-308D-0C09-9A9F-A80F2E6B92D4}")
                                {
                                    oFace61 = Gaika.SurfaceBodies[1].Faces[i] as Face;
                                    break;
                                }
                            }
                            oFace62 = Vint_Model.SurfaceBodies[1].Edges[1] as Face;
                            for (int i = 1; i <= Vint_Model.SurfaceBodies[1].Faces.Count; i++)
                            {
                                if (Vint_Model.SurfaceBodies[1].Faces[i].InternalName == "{99F2EE8C-853D-EA0D-A959-7565D1DAAAA9}")
                                {
                                    oFace62 = Vint_Model.SurfaceBodies[1].Faces[i] as Face;
                                    break;
                                }
                            }
                            Поверхность38 = oAssCompDef.Constraints.AddMateConstraint(oFace59, oFace60, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
                            Поверхность39 = oAssCompDef.Constraints.AddMateConstraint(oFace61, oFace62, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                        }
                        }
                        if (DVint == 14)
                        {
                            if (NowPoint == 8)
                            {
                                var ContTableRow_Designation = ContTableRow.GetCellValue("DESIGNATION");
                                Inventor.MemberManagerErrorsEnum failureCM = default;
                                string MessageProbl = "You_have_a_Problem";
                                var CreateMem = oFamily.CreateMember(ContTableRow, out failureCM, out MessageProbl);
                                ComponentOccurrence Gaika = oAssDoc.ComponentDefinition.Occurrences.Add(CreateMem, oPositionMatrix);//создание вхождения

                                oFace59 = Krishka_Model.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Krishka_Model.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Krishka_Model.SurfaceBodies[1].Faces[i].InternalName == "{2E716E67-D9C4-A22F-9519-083F5E7AA48B}")
                                    {
                                        oFace59 = Krishka_Model.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace60 = Gaika.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Gaika.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Gaika.SurfaceBodies[1].Faces[i].InternalName == "{0F6E14BD-07B3-0DAA-966D-7B80E83827B8}")
                                    {
                                        oFace60 = Gaika.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace61 = Gaika.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Gaika.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Gaika.SurfaceBodies[1].Faces[i].InternalName == "{C5765F4A-308D-0C09-9A9F-A80F2E6B92D4}")
                                    {
                                        oFace61 = Gaika.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace62 = Vint_Model.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Vint_Model.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Vint_Model.SurfaceBodies[1].Faces[i].InternalName == "{99F2EE8C-853D-EA0D-A959-7565D1DAAAA9}")
                                    {
                                        oFace62 = Vint_Model.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                Поверхность38 = oAssCompDef.Constraints.AddMateConstraint(oFace59, oFace60, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
                                Поверхность39 = oAssCompDef.Constraints.AddMateConstraint(oFace61, oFace62, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            }
                        }
                        if (DVint == 18)
                        {
                            if (NowPoint == 12)
                            {
                                var ContTableRow_Designation = ContTableRow.GetCellValue("DESIGNATION");
                                Inventor.MemberManagerErrorsEnum failureCM = default;
                                string MessageProbl = "You_have_a_Problem";
                                var CreateMem = oFamily.CreateMember(ContTableRow, out failureCM, out MessageProbl);
                                ComponentOccurrence Gaika = oAssDoc.ComponentDefinition.Occurrences.Add(CreateMem, oPositionMatrix);//создание вхождения

                                oFace59 = Krishka_Model.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Krishka_Model.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Krishka_Model.SurfaceBodies[1].Faces[i].InternalName == "{2E716E67-D9C4-A22F-9519-083F5E7AA48B}")
                                    {
                                        oFace59 = Krishka_Model.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace60 = Gaika.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Gaika.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Gaika.SurfaceBodies[1].Faces[i].InternalName == "{0F6E14BD-07B3-0DAA-966D-7B80E83827B8}")
                                    {
                                        oFace60 = Gaika.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace61 = Gaika.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Gaika.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Gaika.SurfaceBodies[1].Faces[i].InternalName == "{C5765F4A-308D-0C09-9A9F-A80F2E6B92D4}")
                                    {
                                        oFace61 = Gaika.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace62 = Vint_Model.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Vint_Model.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Vint_Model.SurfaceBodies[1].Faces[i].InternalName == "{99F2EE8C-853D-EA0D-A959-7565D1DAAAA9}")
                                    {
                                        oFace62 = Vint_Model.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                Поверхность38 = oAssCompDef.Constraints.AddMateConstraint(oFace59, oFace60, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
                                Поверхность39 = oAssCompDef.Constraints.AddMateConstraint(oFace61, oFace62, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            }
                        }
                        if (DVint == 20)
                        {
                            if (NowPoint == 14)
                            {
                                var ContTableRow_Designation = ContTableRow.GetCellValue("DESIGNATION");
                                Inventor.MemberManagerErrorsEnum failureCM = default;
                                string MessageProbl = "You_have_a_Problem";
                                var CreateMem = oFamily.CreateMember(ContTableRow, out failureCM, out MessageProbl);
                                ComponentOccurrence Gaika = oAssDoc.ComponentDefinition.Occurrences.Add(CreateMem, oPositionMatrix);//создание вхождения

                                oFace59 = Krishka_Model.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Krishka_Model.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Krishka_Model.SurfaceBodies[1].Faces[i].InternalName == "{2E716E67-D9C4-A22F-9519-083F5E7AA48B}")
                                    {
                                        oFace59 = Krishka_Model.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace60 = Gaika.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Gaika.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Gaika.SurfaceBodies[1].Faces[i].InternalName == "{0F6E14BD-07B3-0DAA-966D-7B80E83827B8}")
                                    {
                                        oFace60 = Gaika.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace61 = Gaika.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Gaika.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Gaika.SurfaceBodies[1].Faces[i].InternalName == "{C5765F4A-308D-0C09-9A9F-A80F2E6B92D4}")
                                    {
                                        oFace61 = Gaika.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                oFace62 = Vint_Model.SurfaceBodies[1].Edges[1] as Face;
                                for (int i = 1; i <= Vint_Model.SurfaceBodies[1].Faces.Count; i++)
                                {
                                    if (Vint_Model.SurfaceBodies[1].Faces[i].InternalName == "{99F2EE8C-853D-EA0D-A959-7565D1DAAAA9}")
                                    {
                                        oFace62 = Vint_Model.SurfaceBodies[1].Faces[i] as Face;
                                        break;
                                    }
                                }
                                Поверхность38 = oAssCompDef.Constraints.AddMateConstraint(oFace59, oFace60, 0, InferredTypeEnum.kNoInference, InferredTypeEnum.kNoInference);
                                Поверхность39 = oAssCompDef.Constraints.AddMateConstraint(oFace61, oFace62, 0, InferredTypeEnum.kInferredLine, InferredTypeEnum.kInferredLine);
                            }
                        }

                    }
                }
            }
            

            MessageBox.Show("Сборка завершена!");
            oAssemblyDocName = oAssDoc;

        }

        private void comboBox30_TextChanged(object sender, EventArgs e)
        {
            LShtok = Convert.ToDouble(comboBox30.Text);
        }

        private void comboBox29_TextChanged(object sender, EventArgs e)
        {
            LShtokRuchka = Convert.ToDouble(comboBox29.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Inventor Assembly Document|*.iam";
            saveFileDialog1.Title = Text;
            saveFileDialog1.FileName = oAssemblyDocName.DisplayName;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrWhiteSpace(saveFileDialog1.FileName))
                {
                    oAssemblyDocName.SaveAs(saveFileDialog1.FileName, false);
                }
            }

        }

        private void comboBox32_TextChanged(object sender, EventArgs e)
        {
            DFlancaVnutr = Convert.ToDouble(comboBox32.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Inventor Part Document|*.ipt";
            openFileDialog1.Title = "Открыть файл регулятора давления";

            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrWhiteSpace(openFileDialog1.FileName))
                {

                    if ("1. Опора.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["1. Опора"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["1. Опора"].DisplayName = "1. Опора";
                        oFileName["1. Опора"] = openFileDialog1.FileName;
                    }
                    if ("3. Винт.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["3. Винт"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["3. Винт"].DisplayName = "3. Винт";
                        oFileName["3. Винт"] = openFileDialog1.FileName;
                    }
                    if ("4. Штифт.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["4. Штифт"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["4. Штифт"].DisplayName = "4. Штифт";
                        oFileName["4. Штифт"] = openFileDialog1.FileName;
                    }
                    if ("5. Упор.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["5. Упор"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["5. Упор"].DisplayName = "5. Упор";
                        oFileName["5. Упор"] = openFileDialog1.FileName;
                    }
                    if ("6. Пружина.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["6. Пружина"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["6. Пружина"].DisplayName = "6. Пружина";
                        oFileName["6. Пружина"] = openFileDialog1.FileName;
                    }
                    if ("7. Пружина.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["7. Пружина"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["7. Пружина"].DisplayName = "7. Пружина";
                        oFileName["7. Пружина"] = openFileDialog1.FileName;
                    }
                    if ("8. Крышка.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["8. Крышка"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["8. Крышка"].DisplayName = "8. Крышка";
                        oFileName["8. Крышка"] = openFileDialog1.FileName;
                    }
                    if ("10. Тарелка.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["10. Тарелка"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["10. Тарелка"].DisplayName = "10. Тарелка";
                        oFileName["10. Тарелка"] = openFileDialog1.FileName;
                    }
                    if ("11. Диафрагма.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["11. Диафрагма"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["11. Диафрагма"].DisplayName = "11. Диафрагма";
                        oFileName["11. Диафрагма"] = openFileDialog1.FileName;
                    }
                    if ("12. Кольцо 80х70.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["12. Кольцо 80х70"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["12. Кольцо 80х70"].DisplayName = "12. Кольцо 80х70";
                        oFileName["12. Кольцо 80х70"] = openFileDialog1.FileName;
                    }
                    if ("13. Корпус.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["13. Корпус"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["13. Корпус"].DisplayName = "13. Корпус";
                        oFileName["13. Корпус"] = openFileDialog1.FileName;
                    }
                    if ("14. Втулка.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["14. Втулка"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["14. Втулка"].DisplayName = "14. Втулка";
                        oFileName["14. Втулка"] = openFileDialog1.FileName;
                    }
                    if ("15. Шток.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["15. Шток"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["15. Шток"].DisplayName = "15. Шток";
                        oFileName["15. Шток"] = openFileDialog1.FileName;
                    }
                    if ("16. Прокладка.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["16. Прокладка"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["16. Прокладка"].DisplayName = "16. Прокладка";
                        oFileName["16. Прокладка"] = openFileDialog1.FileName;
                    }
                    if ("17. Кольцо 55х48.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["17. Кольцо 55х48"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["17. Кольцо 55х48"].DisplayName = "17. Кольцо 55х48";
                        oFileName["17. Кольцо 55х48"] = openFileDialog1.FileName;
                    }
                    if ("18. Пружина.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["18. Пружина"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["18. Пружина"].DisplayName = "18. Пружина";
                        oFileName["18. Пружина"] = openFileDialog1.FileName;
                    }
                    if ("19. Клапан.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["19. Клапан"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["19. Клапан"].DisplayName = "19. Клапан";
                        oFileName["19. Клапан"] = openFileDialog1.FileName;
                    }
                    if ("20. Пробка.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["20. Пробка"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["20. Пробка"].DisplayName = "20. Пробка";
                        oFileName["20. Пробка"] = openFileDialog1.FileName;
                    }
                    if ("21. Шайба.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["21. Шайба"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["21. Шайба"].DisplayName = "21. Шайба";
                        oFileName["21. Шайба"] = openFileDialog1.FileName;
                    }
                    if ("22. Угольник.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["22. Угольник"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["22. Угольник"].DisplayName = "22. Угольник";
                        oFileName["22. Угольник"] = openFileDialog1.FileName;
                    }
                    if ("23. Пробка.ipt" == openFileDialog1.SafeFileName)
                    {
                        oPartDoc["23. Пробка"] = (PartDocument)ThisApplication.Documents.Open(
                        openFileDialog1.FileName, true);
                        oPartDoc["23. Пробка"].DisplayName = "23. Пробка";
                        oFileName["23. Пробка"] = openFileDialog1.FileName;
                    }

                    MessageBox.Show("Файл открыт. Имя файла: " + openFileDialog1.SafeFileName, "Сообщение");
                }
            }

        }

        private void comboBox24_TextChanged(object sender, EventArgs e)
        {
            LYgol = Convert.ToDouble(comboBox24.Text);
        }

        private void comboBox28_TextChanged(object sender, EventArgs e)
        {
            HRezbProbka = Convert.ToDouble(comboBox28.Text);
        }

        private void comboBox31_TextChanged(object sender, EventArgs e)
        {
            DShtok = Convert.ToDouble(comboBox31.Text);
        }

        private void comboBox25_TextChanged(object sender, EventArgs e)
        {
            LTrYgol = Convert.ToDouble(comboBox25.Text);
        }

        private void comboBox27_TextChanged(object sender, EventArgs e)
        {
            HProbka = Convert.ToDouble(comboBox27.Text);
        }

        private void comboBox26_TextChanged_1(object sender, EventArgs e)
        {
            DKolco = Convert.ToDouble(comboBox26.Text);
        }

        private void comboBox20_TextChanged(object sender, EventArgs e)
        {
            DShtift = Convert.ToDouble(comboBox20.Text);
        }

        private void comboBox19_TextChanged(object sender, EventArgs e)
        {
            HShtift = Convert.ToDouble(comboBox19.Text);
        }

        private void comboBox22_TextChanged(object sender, EventArgs e)
        {
            ShirYgol = Convert.ToDouble(comboBox22.Text);
        }

        private void comboBox23_TextChanged(object sender, EventArgs e)
        {
            HYgol = Convert.ToDouble(comboBox23.Text);
        }

        private void comboBox21_TextChanged(object sender, EventArgs e)
        {
            HVtulka = Convert.ToDouble(comboBox21.Text);
        }

        private void comboBox14_TextChanged(object sender, EventArgs e)
        {
            ShirStKr = Convert.ToDouble(comboBox14.Text);
        }

        private void comboBox13_TextChanged(object sender, EventArgs e)
        {
            DVint = Convert.ToDouble(comboBox13.Text);
        }

        private void comboBox18_TextChanged(object sender, EventArgs e)
        {
            HRezbVint = Convert.ToDouble(comboBox18.Text);
        }

        private void comboBox17_TextChanged(object sender, EventArgs e)
        {
            HVint = Convert.ToDouble(comboBox17.Text);
        }

        //private void button6_Click(object sender, EventArgs e)
        //{
        //    UIevents.OnSelect += click;
        //}

        private void comboBox12_TextChanged(object sender, EventArgs e)
        {
            HFlancaKr = Convert.ToDouble(comboBox12.Text);
        }

        private void comboBox8_TextChanged(object sender, EventArgs e)
        {
            RVtullkaSmall = Convert.ToDouble(comboBox8.Text);
        }

        private void comboBox9_TextChanged(object sender, EventArgs e)
        {
            HFlanca = Convert.ToDouble(comboBox9.Text);
        }

        private void comboBox11_TextChanged(object sender, EventArgs e)
        {
            DOtvBolt = Convert.ToDouble(comboBox11.Text);
        }

        private void comboBox10_TextChanged(object sender, EventArgs e)
        {
            DOtvTr = Convert.ToDouble(comboBox10.Text);
        }

        private void comboBox7_TextChanged(object sender, EventArgs e)
        {
            RVtullka = Convert.ToDouble(comboBox7.Text);
        }

        private void comboBox6_TextChanged(object sender, EventArgs e)
        {
            DOsn = Convert.ToDouble(comboBox6.Text);
        }

        private void comboBox5_TextChanged(object sender, EventArgs e)
        {
            HKorp = Convert.ToDouble(comboBox5.Text);
        }

        private void comboBox4_TextChanged(object sender, EventArgs e)
        {
            DFlancaVnesh = Convert.ToDouble(comboBox4.Text);
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            DKolca = Convert.ToDouble(comboBox3.Text);
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            DKlap = Convert.ToDouble(comboBox2.Text);
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            DGasPr = Convert.ToDouble(comboBox1.Text);
        }

        private void click(ObjectsEnumerator JustSelectedEntities,
        ref ObjectCollection MoreSelectedEntities,
        SelectionDeviceEnum SelectionDevice,
        Inventor.Point ModelPosition, Point2d ViewPosition,
        Inventor.View View)
        {
            object obj = JustSelectedEntities[1];
            if (obj is Inventor.Face)
            {
                Inventor.Face obj1 = obj as Inventor.Face;
                string path = @"C:\Temp\interalName.txt";
                string text = String.Join("", obj1.InternalName);

                // полная перезапись файла 
                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(path, false))
                {
                    writer.WriteLine(text);
                }
            }
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            //DesignProject.WorkspasePath();

            //DirectoryInfo dir = new DirectoryInfo("е:\\Регулятор давления\\");
            //foreach (FileInfo f in dir.GetFiles())
            //{
            //    f.Delete();
            //}
            

            //Перевод размеров в см
            DGasPr = DGasPr / 2 / 10;
            DOsn = DOsn / 2 / 10;
            DKlap = DKlap / 2 / 10;
            RVtullka = RVtullka / 10;
            RVtullkaSmall = RVtullkaSmall / 10;
            DKolca = DKolca / 2 / 10;
            DFlancaVnesh = DFlancaVnesh/2 / 10;
            DFlancaVnutr = DFlancaVnutr / 2 / 10;
            HFlanca = HFlanca / 10;
            HKorp = HKorp / 10;
            DOtvTr = DOtvTr / 2 / 10;
            DOtvBolt = DOtvBolt / 2 / 10;
            HFlancaKr = HFlancaKr / 10;
            DVint = DVint / 2 / 10;
            ShirStKr = ShirStKr / 10;
            HKr = HKr / 10;
            HGorlKr = HGorlKr / 10;
            HVint = HVint / 10;
            HRezbVint = HRezbVint / 10;
            DShtift = DShtift / 2 / 10;
            HShtift = HShtift / 10;
            HVtulka = HVtulka / 10;
            HYgol = HYgol / 10;
            ShirYgol = ShirYgol / 10;
            LYgol = LYgol / 10;
            LTrYgol = LTrYgol / 10;
            DKolco = DKolco / 10 / 2;
            HProbka = HProbka / 10;
            HRezbProbka = HRezbProbka / 10;
            DShtok = DShtok / 2 / 10;
            LShtok = LShtok / 10;
            LShtokRuchka = LShtokRuchka / 10;
            HPr6 = HPr6 / 10;
            DPr6 = DPr6 / 2 / 10;
            dPr6 = dPr6 / 2 / 10;
            HPr7 = HPr7 / 10;
            DPr7 = DPr7 / 2 / 10;
            dPr7 = dPr7 / 2 / 10;
            HPr18 = HPr18 / 10;
            DPr18 = DPr18 / 2 / 10;
            dPr18 = dPr18 / 2 / 10;
            DTar = DTar / 2 / 10;
            HTar = HTar / 10;
            hTar = hTar / 10;
            d1Tar = d1Tar /2 / 10;
            d2Tar = d2Tar / 2 / 10;
            DKolca17 = DKolca17 / 2 / 10;
            dKolca17 = dKolca17 / 2 / 10;
            DYpor = DYpor / 2 / 10;
            HYpor = HYpor / 10;
            d1Ypor = d1Ypor / 2 / 10;
            d2Ypor = d2Ypor / 2 / 10;
            h1Ypor = h1Ypor / 10;
            h2Ypor = h2Ypor / 10;
            DOpor = DOpor / 10 / 2;
            HOpor = HOpor / 10;
            d1Opor = d1Opor / 10 / 2;
            h1Opor = h1Opor / 10;
            h2Opor = h2Opor / 10;

            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show(FBD.SelectedPath);
            }
            DirectoryInfo dir = new DirectoryInfo(FBD.SelectedPath);
            foreach (FileInfo f in dir.GetFiles())
            {
                f.Delete();
            }

            //Построение детали 1.Опора
            //Объявление локальных переменных
            //Вставка программного кода по пространственному моделированию детали
            Имя_нового_документа("1. Опора");
            oPartDoc["1. Опора"].DisplayName = "1. Опора";
            PlanarSketch oSketch = oCompDef["1. Опора"].Sketches.Add(oCompDef["1. Опора"].WorkPlanes[3]);
            SketchPoint[] point = new SketchPoint[27];
            SketchLine[] lines = new SketchLine[25];
            SketchArc[] arcs = new SketchArc[5];
            SketchCircle[] Окружность = new SketchCircle[3];

            //Определение координат точек для твердотельного основания
            point[0] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(0, 0), false);
            point[1] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(DOpor-h1Opor/2, 0), false);
            point[2] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(DOpor - h1Opor / 2, h1Opor/2), false);
            point[3] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(DOpor - h1Opor / 2, h1Opor), false);
            point[4] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(d1Opor+0.1, h1Opor), false);
            point[5] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(d1Opor+0.1, h1Opor+0.1), false);
            point[6] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(d1Opor, h1Opor + 0.1), false);
            point[7] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(d1Opor, HOpor-0.1), false);
            point[8] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(d1Opor-0.1, HOpor), false);
            point[9] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(DShtok+0.05, HOpor), false);
            point[10] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(DShtok + 0.05, HOpor-h2Opor), false);
            point[11] = oSketch.SketchPoints.Add(oTransGeom["1. Опора"].CreatePoint2d(0, HOpor - h2Opor-0.05), false);
            //Построение замкнутого контура твердотельного основания
            lines[0] = oSketch.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[2] = oSketch.SketchLines.AddByTwoPoints(point[6], point[7]);
            lines[3] = oSketch.SketchLines.AddByTwoPoints(point[7], point[8]);
            lines[4] = oSketch.SketchLines.AddByTwoPoints(point[8], point[9]);
            lines[5] = oSketch.SketchLines.AddByTwoPoints(point[9], point[10]);
            lines[6] = oSketch.SketchLines.AddByTwoPoints(point[10], point[11]);
            lines[7] = oSketch.SketchLines.AddByTwoPoints(point[11], point[0]);
            arcs[0] = oSketch.SketchArcs.AddByCenterStartEndPoint(oTransGeom["1. Опора"].CreatePoint2d(
            point[2].Geometry.X, point[2].Geometry.Y), point[1], point[3]);
            arcs[1] = oSketch.SketchArcs.AddByCenterStartEndPoint(oTransGeom["1. Опора"].CreatePoint2d(
            point[5].Geometry.X, point[5].Geometry.Y), point[6], point[4]);
            //Принять эскиз 
            oTrans["1. Опора"].End();
            //Выбор функции твердотельного построения
            Profile oProfile = (Profile)oSketch.Profiles.AddForSolid();
            //Вращение эскиза для получения твердотельной модели
            RevolveFeature revolvefeature = oCompDef["1. Опора"].Features.
            RevolveFeatures.AddFull(oProfile, lines[7],
            PartFeatureOperationEnum.kJoinOperation);
            oFileName["1. Опора"] = "1. Опора";
            oPartDoc["1. Опора"].SaveAs(FBD.SelectedPath + "\\1. Опора.ipt", false);
            //Save_Model("1. Опора", "1. Опора");


            //Построение детали 13. Корпус
            Имя_нового_документа("13. Корпус");
            oPartDoc["13. Корпус"].DisplayName = "13. Корпус";
            PlanarSketch oSketch21 = oCompDef["13. Корпус"].Sketches.Add(oCompDef["13. Корпус"].WorkPlanes[3]);

            point[0] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr, 1.4), false);
            point[1] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DKlap, 1.4), false);
            point[2] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DKlap, 0.4), false);
            point[3] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn - 0.4, 0.4), false);
            point[4] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn - 0.4, 0), false);
            point[5] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn, 0.4), false);
            point[6] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 1.05, 5), false);
            point[7] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 1.35, HKorp - HFlanca), false);
            point[8] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DFlancaVnesh, HKorp - HFlanca), false);
            point[9] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DFlancaVnesh, HKorp), false);
            point[10] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DFlancaVnutr, HKorp), false);
            point[11] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DFlancaVnutr, HKorp - 0.3), false);
            point[12] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DKolca, HKorp - 0.3), false);
            point[13] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DKolca, HKorp - 0.6), false);
            point[14] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DKolca - DKolco, HKorp - 0.6), false);
            point[15] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d((DKolca - DKolco) * 0.96, HKorp - 0.2), false);
            point[16] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn, HKorp - 0.2), false);
            point[17] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn, HKorp - 0.9), false);
            point[18] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 0.67, HKorp - 1.5), false);
            point[19] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(RVtullka, HKorp - 1.5), false);
            point[20] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(RVtullka, HKorp - 1.7), false);
            point[21] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(RVtullkaSmall, HKorp - 1.7), false);
            point[22] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(RVtullkaSmall, 5), false);
            point[23] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr, 5), false);
            point[24] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, 0), false);
            point[25] = oSketch21.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, 5), false);
            lines[0] = oSketch21.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch21.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch21.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch21.SketchLines.AddByTwoPoints(point[3], point[4]);
            arcs[0] = oSketch21.SketchArcs.AddByCenterStartEndPoint(oTransGeom["13. Корпус"].CreatePoint2d(
            point[3].Geometry.X, point[3].Geometry.Y), point[4], point[5]);
            lines[4] = oSketch21.SketchLines.AddByTwoPoints(point[5], point[6]);
            lines[5] = oSketch21.SketchLines.AddByTwoPoints(point[6], point[7]);
            lines[6] = oSketch21.SketchLines.AddByTwoPoints(point[7], point[8]);
            lines[7] = oSketch21.SketchLines.AddByTwoPoints(point[8], point[9]);
            lines[8] = oSketch21.SketchLines.AddByTwoPoints(point[9], point[10]);
            lines[9] = oSketch21.SketchLines.AddByTwoPoints(point[10], point[11]);
            lines[10] = oSketch21.SketchLines.AddByTwoPoints(point[11], point[12]);
            lines[11] = oSketch21.SketchLines.AddByTwoPoints(point[12], point[13]);
            lines[12] = oSketch21.SketchLines.AddByTwoPoints(point[13], point[14]);
            lines[13] = oSketch21.SketchLines.AddByTwoPoints(point[14], point[15]);
            lines[14] = oSketch21.SketchLines.AddByTwoPoints(point[15], point[16]);
            lines[15] = oSketch21.SketchLines.AddByTwoPoints(point[16], point[17]);
            lines[16] = oSketch21.SketchLines.AddByTwoPoints(point[17], point[18]);
            lines[17] = oSketch21.SketchLines.AddByTwoPoints(point[18], point[19]);
            lines[18] = oSketch21.SketchLines.AddByTwoPoints(point[19], point[20]);
            lines[19] = oSketch21.SketchLines.AddByTwoPoints(point[20], point[21]);
            lines[20] = oSketch21.SketchLines.AddByTwoPoints(point[21], point[22]);
            lines[21] = oSketch21.SketchLines.AddByTwoPoints(point[22], point[23]);
            lines[22] = oSketch21.SketchLines.AddByTwoPoints(point[23], point[0]);
            lines[23] = oSketch21.SketchLines.AddByTwoPoints(point[24], point[25]);
            oTrans["13. Корпус"].End();
            Profile oProfile21 = (Profile)oSketch21.Profiles.AddForSolid();
            // Вращение эскиза для получения твердотельной модели
            RevolveFeature revolvefeature17 = oCompDef["13. Корпус"].Features.
            RevolveFeatures.AddFull(oProfile21, lines[23],
            PartFeatureOperationEnum.kJoinOperation);

            PlanarSketch oSketch22 = oCompDef["13. Корпус"].Sketches.Add(oCompDef["13. Корпус"].WorkPlanes[3]);
            Transaction oTrans2 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample2");
            point[0] = oSketch22.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr, HRezbProbka), false);
            point[1] = oSketch22.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DKlap + 0.1, HRezbProbka), false);
            point[2] = oSketch22.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DKlap + 0.1, HRezbProbka + 0.6), false);
            point[3] = oSketch22.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr + 0.25, HRezbProbka + 0.6), false);
            point[4] = oSketch22.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr + 0.15, HRezbProbka + 0.5), false);
            point[5] = oSketch22.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr, HRezbProbka + 0.5), false);
            point[6] = oSketch22.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, 0), false);
            point[7] = oSketch22.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, 2), false);
            lines[0] = oSketch22.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch22.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch22.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch22.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch22.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch22.SketchLines.AddByTwoPoints(point[5], point[0]);
            lines[6] = oSketch22.SketchLines.AddByTwoPoints(point[6], point[7]);
            oTrans2.End();
            Profile oProfile22 = (Profile)oSketch22.Profiles.AddForSolid();
            RevolveFeature revolvefeature18 = oCompDef["13. Корпус"].Features.
                RevolveFeatures.AddFull(oProfile22, lines[6],
                PartFeatureOperationEnum.kCutOperation);

            WorkPlane oWorkPlane3 = oCompDef["13. Корпус"].WorkPlanes.AddByPlaneAndOffset(
            oCompDef["13. Корпус"].WorkPlanes[2], HKorp - HFlanca + " см", false);
            oWorkPlane3.Visible = false;
            PlanarSketch oSketch23 = oCompDef["13. Корпус"].Sketches.Add(oWorkPlane3, false);
            Transaction oTrans3 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample2");
            point[0] = oSketch23.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 1.67, DOsn * 0.79), false);
            point[1] = oSketch23.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 1.67, -DOsn * 0.79), false);
            point[2] = oSketch23.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(-DOsn * 1.67, -DOsn * 0.79), false);
            point[3] = oSketch23.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(-DOsn * 1.67, DOsn * 0.79), false);
            point[4] = oSketch23.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, 0), false);
            point[5] = oSketch23.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 0.61, DOsn * 0.79), false);
            point[6] = oSketch23.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 0.61, -DOsn * 0.79), false);
            point[7] = oSketch23.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(-DOsn * 0.61, DOsn * 0.79), false);
            point[8] = oSketch23.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(-DOsn * 0.61, -DOsn * 0.79), false);
            lines[0] = oSketch23.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch23.SketchLines.AddByTwoPoints(point[1], point[6]);
            lines[2] = oSketch23.SketchLines.AddByTwoPoints(point[5], point[0]);
            lines[3] = oSketch23.SketchLines.AddByTwoPoints(point[8], point[2]);
            lines[4] = oSketch23.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[5] = oSketch23.SketchLines.AddByTwoPoints(point[3], point[7]);
            arcs[0] = oSketch23.SketchArcs.AddByCenterStartEndPoint(oTransGeom["13. Корпус"].CreatePoint2d(
            point[4].Geometry.X, point[4].Geometry.Y), point[6], point[5]);
            arcs[1] = oSketch23.SketchArcs.AddByCenterStartEndPoint(oTransGeom["13. Корпус"].CreatePoint2d(
            point[4].Geometry.X, point[4].Geometry.Y), point[7], point[8]);
            oTrans3.End();
            Profile oProfile23 = (Profile)oSketch23.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef5 = oCompDef["13. Корпус"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile23,/*Длина в см*/4.3,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kJoinOperation,/*Эскиз*/oProfile23);

            EdgeCollection EdgeCollection15 = ThisApplication.TransientObjects.CreateEdgeCollection();
            FilletFeature oFillet3 = default(FilletFeature);
            EdgeCollection15.Add(oExtrudeDef5.Faces[1].Edges[6]);
            EdgeCollection15.Add(oExtrudeDef5.Faces[6].Edges[4]);
            EdgeCollection15.Add(oExtrudeDef5.Faces[7].Edges[2]);
            EdgeCollection15.Add(oExtrudeDef5.Faces[7].Edges[4]);
            FilletDefinition oFilletDef4 = oCompDef["13. Корпус"].Features.FilletFeatures.CreateFilletDefinition();
            oFilletDef4.AddConstantRadiusEdgeSet(EdgeCollection15, 24 + "мм");
            oFillet3 = oCompDef["13. Корпус"].Features.FilletFeatures.Add(oFilletDef4, false);

            PlanarSketch oSketch24 = oCompDef["13. Корпус"].Sketches.Add(oCompDef["13. Корпус"].WorkPlanes[3]);
            Transaction oTrans4 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample4");
            point[0] = oSketch24.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(-DOsn * 1.67, 4), false);
            point[1] = oSketch24.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(-DOsn * 1.67, DOtvTr + 4), false);
            point[2] = oSketch24.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(-1.5, DOtvTr * 0.87 + 4), false);
            point[3] = oSketch24.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(-1.5, 4), false);
            lines[0] = oSketch24.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch24.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch24.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch24.SketchLines.AddByTwoPoints(point[3], point[0]);
            oTrans4.End();
            Profile oProfile24 = (Profile)oSketch24.Profiles.AddForSolid();
            RevolveFeature revolvefeature19 = oCompDef["13. Корпус"].Features.
                RevolveFeatures.AddFull(oProfile24, lines[3],
                PartFeatureOperationEnum.kCutOperation);

            PlanarSketch oSketch25 = oCompDef["13. Корпус"].Sketches.Add(oCompDef["13. Корпус"].WorkPlanes[3]);
            Transaction oTrans5 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample5");
            point[0] = oSketch25.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 1.67, 4), false);
            point[1] = oSketch25.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 1.67, DOtvTr + 4), false);
            point[2] = oSketch25.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DOsn * 1.67 - 3.2, DOtvTr * 0.87 + 4), false);
            point[3] = oSketch25.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(1.5, DOtvTr * 0.54 + 4), false);
            point[4] = oSketch25.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, DOtvTr * 0.54 + 4), false);
            point[5] = oSketch25.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, 4), false);
            lines[0] = oSketch25.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch25.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch25.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch25.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch25.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch25.SketchLines.AddByTwoPoints(point[5], point[0]);
            oTrans5.End();
            Profile oProfile25 = (Profile)oSketch25.Profiles.AddForSolid();
            RevolveFeature revolvefeature20 = oCompDef["13. Корпус"].Features.
                RevolveFeatures.AddFull(oProfile25, lines[5],
                PartFeatureOperationEnum.kCutOperation);

            WorkPlane oWorkPlane4 = oCompDef["13. Корпус"].WorkPlanes.AddByPlaneAndOffset(
            oCompDef["13. Корпус"].WorkPlanes[2], 20 + " мм", false);
            oWorkPlane4.Visible = false;
            PlanarSketch oSketch26 = oCompDef["13. Корпус"].Sketches.Add(oWorkPlane4, false);
            Transaction oTrans6 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample6");
            point[0] = oSketch26.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, 0), false);
            point[1] = oSketch26.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr * 0.75, DGasPr * 1.299), false);
            point[2] = oSketch26.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr * 1.05, DGasPr * 1.819), false);
            point[3] = oSketch26.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr * 1.05, -DGasPr * 1.819), false);
            point[4] = oSketch26.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(DGasPr * 0.75, -DGasPr * 1.299), false);
            lines[0] = oSketch26.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[1] = oSketch26.SketchLines.AddByTwoPoints(point[3], point[4]);
            arcs[0] = oSketch26.SketchArcs.AddByCenterStartEndPoint(oTransGeom["13. Корпус"].CreatePoint2d(
            point[0].Geometry.X, point[0].Geometry.Y), point[4], point[1]);
            arcs[1] = oSketch26.SketchArcs.AddByCenterStartEndPoint(oTransGeom["13. Корпус"].CreatePoint2d(
            point[0].Geometry.X, point[0].Geometry.Y), point[3], point[2]);
            oTrans6.End();
            Profile oProfile26 = (Profile)oSketch26.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef6 = oCompDef["13. Корпус"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile26,/*Длина в см*/DOtvTr * 0.87 + 2,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kCutOperation,/*Эскиз*/oProfile26);

            WorkPlane oWorkPlane5 = oCompDef["13. Корпус"].WorkPlanes.AddByPlaneAndOffset(
            oCompDef["13. Корпус"].WorkPlanes[2], HKorp - 1.5 + " см", false);
            oWorkPlane5.Visible = false;
            PlanarSketch oSketch27 = oCompDef["13. Корпус"].Sketches.Add(oWorkPlane5, false);
            Transaction oTrans7 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample7");
            point[0] = oSketch27.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(-RVtullka * 1.87, 0), false);
            Окружность[0] = oSketch27.SketchCircles.AddByCenterRadius(point[0], 0.1);
            oTrans7.End();
            Profile oProfile27 = (Profile)oSketch27.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef7 = oCompDef["13. Корпус"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile27,/*Длина в см*/2,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kCutOperation,/*Эскиз*/oProfile27);

            PlanarSketch oSketch28 = oCompDef["13. Корпус"].Sketches.Add(oCompDef["13. Корпус"].WorkPlanes[1]);
            Transaction oTrans8 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample8");
            point[0] = oSketch28.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(4, DOsn * 1.05), false);
            point[1] = oSketch28.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(4.515, DOsn * 1.05), false);
            point[2] = oSketch28.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(4.425, 1.6), false);
            point[3] = oSketch28.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(4.3, 1.475), false);
            point[4] = oSketch28.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(4, 1.475), false);
            lines[0] = oSketch28.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch28.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch28.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch28.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch28.SketchLines.AddByTwoPoints(point[4], point[0]);
            oTrans8.End();
            Profile oProfile28 = (Profile)oSketch28.Profiles.AddForSolid();
            RevolveFeature revolvefeature21 = oCompDef["13. Корпус"].Features.
                RevolveFeatures.AddFull(oProfile28, lines[4],
                PartFeatureOperationEnum.kCutOperation);
            //Зеркальное отображение
            ObjectCollection objCollection = ThisApplication.TransientObjects.CreateObjectCollection();
            MirrorFeature mirrorFeature;
            objCollection.Add(revolvefeature21);
            mirrorFeature = oCompDef["13. Корпус"].Features.MirrorFeatures.Add(objCollection, oCompDef["13. Корпус"].WorkPlanes[3], false,
                PatternComputeTypeEnum.kAdjustToModelCompute);

            PlanarSketch oSketch29 = oCompDef["13. Корпус"].Sketches.Add(oCompDef["13. Корпус"].WorkPlanes[3]);
            Transaction oTrans9 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample9");
            point[0] = oSketch29.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, 4), false);
            Окружность[0] = oSketch29.SketchCircles.AddByCenterRadius(point[0], 0.3);
            oTrans9.End();
            Profile oProfile29 = (Profile)oSketch29.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef8 = oCompDef["13. Корпус"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile29,/*Длина в см*/5,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kCutOperation,/*Эскиз*/oProfile29);

            WorkPlane oWorkPlane6 = oCompDef["13. Корпус"].WorkPlanes.AddByPlaneAndOffset(
            oCompDef["13. Корпус"].WorkPlanes[2], HKorp + " см", false);
            oWorkPlane6.Visible = false;
            PlanarSketch oSketch30 = oCompDef["13. Корпус"].Sketches.Add(oWorkPlane6, false);
            Transaction oTrans10 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample10");
            point[0] = oSketch30.SketchPoints.Add(oTransGeom["13. Корпус"].CreatePoint2d(0, DFlancaVnutr + DOtvBolt + 0.05), false);
            Окружность[0] = oSketch30.SketchCircles.AddByCenterRadius(point[0], DOtvBolt);
            oTrans10.End();
            Profile oProfile30 = (Profile)oSketch30.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef9 = oCompDef["13. Корпус"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile30,/*Длина в см*/HFlanca,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kCutOperation,/*Эскиз*/oProfile30);
            //Построение кругового массива отверстий
            WorkAxis Axis = oCompDef["13. Корпус"].WorkAxes[2];
            ObjectCollection objCollection1 = ThisApplication.TransientObjects.CreateObjectCollection();
            objCollection1.Add(oExtrudeDef9);
            CircularPatternFeature CircularPatternFeature = oCompDef["13. Корпус"].Features.CircularPatternFeatures.Add(objCollection1, Axis,
                false, 6, 360 + "degree", true, PatternComputeTypeEnum.kIdenticalCompute);
            oFileName["13. Корпус"] = "13. Корпус";
            oPartDoc["13. Корпус"].SaveAs(FBD.SelectedPath + "\\13. Корпус.ipt", false);

            Имя_нового_документа("8. Крышка");
            oPartDoc["8. Крышка"].DisplayName = "8. Крышка";
            PlanarSketch oSketch31 = oCompDef["8. Крышка"].Sketches.Add(oCompDef["8. Крышка"].WorkPlanes[3]);
            point[0] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DFlancaVnutr, 0), false);
            point[1] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DFlancaVnutr, 0.7), false);
            point[2] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DFlancaVnesh, 0.7), false);
            point[3] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DFlancaVnesh, HFlancaKr + 0.7), false);
            point[4] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DOsn * 1.03, HFlancaKr + 0.7), false);
            point[5] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DOsn * 0.86, HKr - HGorlKr), false);
            point[6] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DVint + ShirStKr, HKr - HGorlKr), false);
            point[7] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DVint + ShirStKr, HKr), false);
            point[8] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DVint, HKr), false);
            point[9] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DVint, HKr - HGorlKr - ShirStKr), false);
            point[10] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DOsn * 0.86 - ShirStKr, HKr - HGorlKr - ShirStKr), false);
            point[11] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DOsn * 1.03 - ShirStKr, 1.4), false);
            point[12] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DTar, 1.4), false);
            point[13] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DTar, 0.3), false);
            point[14] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DTar*1.24, 0.3), false);
            point[15] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DTar * 1.24, 0.2), false);
            point[16] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DTar * 1.24 + 0.1, 0.2), false);
            point[17] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DTar * 1.24 + 0.2, 0.1), false);
            point[18] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DTar * 1.24 + 0.3, 0.1), false);
            point[19] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(DTar * 1.24 + 0.3, 0), false);
            point[20] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(0, 0), false);
            point[21] = oSketch31.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(0, 5), false);
            lines[0] = oSketch31.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch31.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch31.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch31.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch31.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch31.SketchLines.AddByTwoPoints(point[5], point[6]);
            lines[6] = oSketch31.SketchLines.AddByTwoPoints(point[6], point[7]);
            lines[7] = oSketch31.SketchLines.AddByTwoPoints(point[7], point[8]);
            lines[8] = oSketch31.SketchLines.AddByTwoPoints(point[8], point[9]);
            lines[9] = oSketch31.SketchLines.AddByTwoPoints(point[9], point[10]);
            lines[10] = oSketch31.SketchLines.AddByTwoPoints(point[10], point[11]);
            lines[11] = oSketch31.SketchLines.AddByTwoPoints(point[11], point[12]);
            lines[12] = oSketch31.SketchLines.AddByTwoPoints(point[12], point[13]);
            lines[13] = oSketch31.SketchLines.AddByTwoPoints(point[13], point[14]);
            arcs[0] = oSketch31.SketchArcs.AddByCenterStartEndPoint(oTransGeom["8. Крышка"].CreatePoint2d(
            point[15].Geometry.X, point[15].Geometry.Y), point[16], point[14]);
            lines[14] = oSketch31.SketchLines.AddByTwoPoints(point[16], point[17]);
            arcs[1] = oSketch31.SketchArcs.AddByCenterStartEndPoint(oTransGeom["8. Крышка"].CreatePoint2d(
            point[18].Geometry.X, point[18].Geometry.Y), point[17], point[19]);
            lines[15] = oSketch31.SketchLines.AddByTwoPoints(point[19], point[0]);
            lines[16] = oSketch31.SketchLines.AddByTwoPoints(point[20], point[21]);
            oTrans["8. Крышка"].End();
            Profile oProfile31 = (Profile)oSketch31.Profiles.AddForSolid();
            RevolveFeature revolvefeature22 = oCompDef["8. Крышка"].Features.
                RevolveFeatures.AddFull(oProfile31, lines[16],
                PartFeatureOperationEnum.kJoinOperation);
            EdgeCollection EdgeCollection17 = ThisApplication.TransientObjects.CreateEdgeCollection();
            FilletFeature oFillet5 = default(FilletFeature);
            EdgeCollection17.Add(revolvefeature22.Faces[14].Edges[1]);
            EdgeCollection17.Add(revolvefeature22.Faces[14].Edges[2]);
            FilletDefinition oFilletDef5 = oCompDef["8. Крышка"].Features.FilletFeatures.CreateFilletDefinition();
            oFilletDef5.AddConstantRadiusEdgeSet(EdgeCollection17, 5 + "мм");
            oFillet5 = oCompDef["8. Крышка"].Features.FilletFeatures.Add(oFilletDef5, false);
            EdgeCollection EdgeCollection18 = ThisApplication.TransientObjects.CreateEdgeCollection();
            //FilletFeature oFillet6 = default(FilletFeature);
            FilletDefinition oFilletDef6 = oCompDef["8. Крышка"].Features.FilletFeatures.CreateFilletDefinition();
            EdgeCollection18.Add(revolvefeature22.Faces[15].Edges[2]);
            oFilletDef6.AddConstantRadiusEdgeSet(EdgeCollection18, 3 + "мм");
            oFillet5 = oCompDef["8. Крышка"].Features.FilletFeatures.Add(oFilletDef6, false);

            PlanarSketch oSketch32 = oCompDef["8. Крышка"].Sketches.Add(oCompDef["8. Крышка"].WorkPlanes[3]);
            oTrans["8. Крышка"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch32.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(0, 0.33 * HKr), false);
            Окружность[0] = oSketch32.SketchCircles.AddByCenterRadius(point[0], 0.15);
            oTrans["8. Крышка"].End();
            Profile oProfile32 = (Profile)oSketch32.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef10 = oCompDef["8. Крышка"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile32,/*Длина в см*/DOsn * 1.3,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kCutOperation,/*Эскиз*/oProfile32);

            WorkPlane oWorkPlane7 = oCompDef["8. Крышка"].WorkPlanes.AddByPlaneAndOffset(
            oCompDef["8. Крышка"].WorkPlanes[2], 0.7 + HFlancaKr + " см", false);
            oWorkPlane7.Visible = false;
            PlanarSketch oSketch33 = oCompDef["8. Крышка"].Sketches.Add(oWorkPlane7, false);
            oTrans["8. Крышка"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch33.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(0, DFlancaVnutr + DOtvBolt + 0.05), false);
            Окружность[0] = oSketch33.SketchCircles.AddByCenterRadius(point[0], 0.675);
            oTrans["8. Крышка"].End();
            Profile oProfile33 = (Profile)oSketch33.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef11 = oCompDef["8. Крышка"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile33,/*Длина в см*/0.7,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kCutOperation,/*Эскиз*/oProfile33);

            WorkPlane oWorkPlane8 = oCompDef["8. Крышка"].WorkPlanes.AddByPlaneAndOffset(
            oCompDef["8. Крышка"].WorkPlanes[2], HFlancaKr + " см", false);
            oWorkPlane8.Visible = false;
            PlanarSketch oSketch34 = oCompDef["8. Крышка"].Sketches.Add(oWorkPlane8, false);
            oTrans["8. Крышка"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch34.SketchPoints.Add(oTransGeom["8. Крышка"].CreatePoint2d(0, DFlancaVnutr + DOtvBolt + 0.05), false);
            Окружность[0] = oSketch34.SketchCircles.AddByCenterRadius(point[0], DOtvBolt);
            oTrans["8. Крышка"].End();
            Profile oProfile34 = (Profile)oSketch34.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef12 = oCompDef["8. Крышка"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile34,/*Длина в см*/HFlancaKr,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kCutOperation,/*Эскиз*/oProfile34);

            //Построение кругового массива отверстий
            WorkAxis Axis1 = oCompDef["8. Крышка"].WorkAxes[2];
            ObjectCollection objCollection2 = ThisApplication.TransientObjects.CreateObjectCollection();
            objCollection2.Add(oExtrudeDef11);
            objCollection2.Add(oExtrudeDef12);
            CircularPatternFeature CircularPatternFeature1 = oCompDef["8. Крышка"].Features.CircularPatternFeatures.Add(objCollection2, Axis1,
                false, 6, 360 + "degree", true, PatternComputeTypeEnum.kIdenticalCompute);
            oPartDoc["8. Крышка"].SaveAs(FBD.SelectedPath + "\\8. Крышка.ipt", false);

            //EdgeCollection EdgeCollection16 = ThisApplication.TransientObjects.CreateEdgeCollection();
            //FilletFeature oFillet4 = default(FilletFeature);
            //EdgeCollection16.Add(oExtrudeDef5.Faces[5].Edges[5]);
            //EdgeCollection15.Add(oExtrudeDef5.Faces[8].Edges[1]);
            //FilletDefinition oFilletDef5 = oCompDef["13. Корпус"].Features.FilletFeatures.CreateFilletDefinition();
            //oFilletDef5.AddConstantRadiusEdgeSet(EdgeCollection16, 1.5 + "мм");
            //oFillet3 = oCompDef["13. Корпус"].Features.FilletFeatures.Add(oFilletDef5, false);

            //EdgeCollection EdgeCollection15 = ThisApplication.TransientObjects.CreateEdgeCollection();
            //EdgeCollection15.Add(oExtrudeDef5.Faces[3].Edges[2]);
            //EdgeCollection15.Add(oExtrudeDef5.Faces[2].Edges[2]);
            //EdgeCollection15.Add(oExtrudeDef5.Faces[2].Edges[3]);
            //EdgeCollection15.Add(oExtrudeDef5.Faces[2].Edges[4]);
            //oCompDef["13. Корпус"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection15, 1 + "мм", true);

            //Построение детали 3. Винт
            Имя_нового_документа("3. Винт");
            oPartDoc["3. Винт"].DisplayName = "3. Винт";
            PlanarSketch oSketch1 = oCompDef["3. Винт"].Sketches.Add(oCompDef["3. Винт"].WorkPlanes[3]);
            point[0] = oSketch1.SketchPoints.Add(oTransGeom["3. Винт"].CreatePoint2d(0, 0), false);
            point[1] = oSketch1.SketchPoints.Add(oTransGeom["3. Винт"].CreatePoint2d(0, DVint - 0.2), false);
            point[2] = oSketch1.SketchPoints.Add(oTransGeom["3. Винт"].CreatePoint2d(DVint - 0.2, DVint - 0.2), false);
            point[3] = oSketch1.SketchPoints.Add(oTransGeom["3. Винт"].CreatePoint2d(DVint, DVint), false);
            point[4] = oSketch1.SketchPoints.Add(oTransGeom["3. Винт"].CreatePoint2d(DVint, 4.5), false);
            point[5] = oSketch1.SketchPoints.Add(oTransGeom["3. Винт"].CreatePoint2d(DVint, HVint - 0.1), false);
            point[6] = oSketch1.SketchPoints.Add(oTransGeom["3. Винт"].CreatePoint2d(DVint - 0.1, HVint), false);
            point[7] = oSketch1.SketchPoints.Add(oTransGeom["3. Винт"].CreatePoint2d(0, HVint), false);
            lines[0] = oSketch1.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[1] = oSketch1.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[2] = oSketch1.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[3] = oSketch1.SketchLines.AddByTwoPoints(point[5], point[6]);
            lines[4] = oSketch1.SketchLines.AddByTwoPoints(point[6], point[7]);
            lines[4] = oSketch1.SketchLines.AddByTwoPoints(point[7], point[0]);
            arcs[0] = oSketch1.SketchArcs.AddByCenterStartEndPoint(oTransGeom["3. Винт"].CreatePoint2d(
            point[1].Geometry.X, point[1].Geometry.Y), point[0], point[2]);
            oTrans["3. Винт"].End();
            Profile oProfile1 = (Profile)oSketch1.Profiles.AddForSolid();
            // Вращение эскиза для получения твердотельной модели
            RevolveFeature revolvefeature1 = oCompDef["3. Винт"].Features.
            RevolveFeatures.AddFull(oProfile1, lines[4],
            PartFeatureOperationEnum.kJoinOperation);
            // Создание плоскости построения "oWorkPlane"
            WorkPlane oWorkPlane = oCompDef["3. Винт"].WorkPlanes.AddByPlaneAndOffset(
            oCompDef["3. Винт"].WorkPlanes[3], DVint + " см", false);
            oWorkPlane.Visible = false;
            //Выбор рабочей плоскости oWorkPlane и создание эскиза на плоскости "oSketch1"
            PlanarSketch oSketch2 = oCompDef["3. Винт"].Sketches.Add(oWorkPlane, false);
            oTrans["3. Винт"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch2.SketchPoints.Add(oTransGeom["3. Винт"].CreatePoint2d(0, 0.87 * HVint), false);
            Окружность[0] = oSketch2.SketchCircles.AddByCenterRadius(point[0], DShtift);
            oTrans["3. Винт"].End();
            Profile oProfile2 = (Profile)oSketch2.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef = oCompDef["3. Винт"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile2,/*Длина в см*/DVint * 2,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kCutOperation,/*Эскиз*/oProfile2);
            //Резьба

            EdgeCollection EdgeCollection = ThisApplication.TransientObjects.CreateEdgeCollection();
            Face Face = oCompDef["3. Винт"].SurfaceBodies[1].Faces[1];
            for (int i = 1; i <= oCompDef["3. Винт"].SurfaceBodies[1].Faces.Count; i++)
            {
                if (oCompDef["3. Винт"].SurfaceBodies[1].Faces[i].InternalName == "{99F2EE8C-853D-EA0D-A959-7565D1DAAAA9}")
                {
                    Face = oCompDef["3. Винт"].SurfaceBodies[1].Faces[i] as Face;
                    break;
                }
            }
            Edge StartEdge_1 = Face.Edges[5];
            EdgeCollection.Add(StartEdge_1);
            ThreadFeatures ThreadFeatures = oCompDef["3. Винт"].Features.ThreadFeatures;
            StandardThreadInfo stInfo = ThreadFeatures.CreateStandardThreadInfo(false, true,
            "ISO Metric profile", "M16x1.25", "6g");
            ThreadInfo ThreadInfo = (ThreadInfo)stInfo;
            ThreadFeatures.Add(Face, StartEdge_1, ThreadInfo, false, false, HRezbVint + " см", 0);
            oPartDoc["3. Винт"].SaveAs(FBD.SelectedPath + "\\3. Винт.ipt", false);

            //Построение детали 4. Штифт
            Имя_нового_документа("4. Штифт");
            oPartDoc["4. Штифт"].DisplayName = "4. Штифт";
            PlanarSketch oSketch3 = oCompDef["4. Штифт"].Sketches.Add(oCompDef["4. Штифт"].WorkPlanes[3]);
            point[0] = oSketch3.SketchPoints.Add(oTransGeom["4. Штифт"].CreatePoint2d(0, 0), false);
            point[1] = oSketch3.SketchPoints.Add(oTransGeom["4. Штифт"].CreatePoint2d(0, DShtift), false);
            point[2] = oSketch3.SketchPoints.Add(oTransGeom["4. Штифт"].CreatePoint2d(DShtift, DShtift), false);
            point[3] = oSketch3.SketchPoints.Add(oTransGeom["4. Штифт"].CreatePoint2d(DShtift, HShtift - DShtift), false);
            point[4] = oSketch3.SketchPoints.Add(oTransGeom["4. Штифт"].CreatePoint2d(0, HShtift - DShtift), false);
            point[5] = oSketch3.SketchPoints.Add(oTransGeom["4. Штифт"].CreatePoint2d(0, HShtift), false);
            lines[0] = oSketch3.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[1] = oSketch3.SketchLines.AddByTwoPoints(point[5], point[0]);
            arcs[0] = oSketch3.SketchArcs.AddByCenterStartEndPoint(oTransGeom["4. Штифт"].CreatePoint2d(
            point[1].Geometry.X, point[1].Geometry.Y), point[0], point[2]);
            arcs[1] = oSketch3.SketchArcs.AddByCenterStartEndPoint(oTransGeom["4. Штифт"].CreatePoint2d(
            point[4].Geometry.X, point[4].Geometry.Y), point[3], point[5]);
            oTrans["4. Штифт"].End();
            Profile oProfile3 = (Profile)oSketch3.Profiles.AddForSolid();
            RevolveFeature revolvefeature2 = oCompDef["4. Штифт"].Features.
            RevolveFeatures.AddFull(oProfile3, lines[1],
            PartFeatureOperationEnum.kJoinOperation);
            oPartDoc["4. Штифт"].SaveAs(FBD.SelectedPath + "\\4. Штифт.ipt", false);

            //Построение детали 5. Упор
            Имя_нового_документа("5. Упор");
            oPartDoc["5. Упор"].DisplayName = "5. Упор";
            PlanarSketch oSketch4 = oCompDef["5. Упор"].Sketches.Add(oCompDef["5. Упор"].WorkPlanes[3]);
            point[0] = oSketch4.SketchPoints.Add(oTransGeom["5. Упор"].CreatePoint2d(0.7, 0), false);
            point[1] = oSketch4.SketchPoints.Add(oTransGeom["5. Упор"].CreatePoint2d(DYpor, 0), false);
            point[2] = oSketch4.SketchPoints.Add(oTransGeom["5. Упор"].CreatePoint2d(DYpor, HYpor-h1Ypor), false);
            point[3] = oSketch4.SketchPoints.Add(oTransGeom["5. Упор"].CreatePoint2d(d1Ypor, HYpor - h1Ypor), false);
            point[4] = oSketch4.SketchPoints.Add(oTransGeom["5. Упор"].CreatePoint2d(d1Ypor, HYpor), false);
            point[5] = oSketch4.SketchPoints.Add(oTransGeom["5. Упор"].CreatePoint2d(d2Ypor, HYpor), false);
            point[6] = oSketch4.SketchPoints.Add(oTransGeom["5. Упор"].CreatePoint2d(d2Ypor, HYpor-h2Ypor), false);
            point[7] = oSketch4.SketchPoints.Add(oTransGeom["5. Упор"].CreatePoint2d(0, HYpor - h2Ypor), false);
            point[8] = oSketch4.SketchPoints.Add(oTransGeom["5. Упор"].CreatePoint2d(0, HYpor - h1Ypor), false);
            lines[0] = oSketch4.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch4.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch4.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch4.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch4.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch4.SketchLines.AddByTwoPoints(point[5], point[6]);
            lines[6] = oSketch4.SketchLines.AddByTwoPoints(point[6], point[7]);
            lines[7] = oSketch4.SketchLines.AddByTwoPoints(point[7], point[8]);
            lines[8] = oSketch4.SketchLines.AddByTwoPoints(point[8], point[0]);

            oTrans["5. Упор"].End();

            Profile oProfile4 = (Profile)oSketch4.Profiles.AddForSolid();
            RevolveFeature revolvefeature3 = oCompDef["5. Упор"].Features.
            RevolveFeatures.AddFull(oProfile4, lines[7],
            PartFeatureOperationEnum.kJoinOperation);
            //фаски
            EdgeCollection EdgeCollection1 = ThisApplication.TransientObjects.CreateEdgeCollection();
            Face Face1 = revolvefeature3.SideFaces[3];
            Edge StartEdge_2 = Face1.Edges[1];
            EdgeCollection1.Add(StartEdge_2);
            oCompDef["5. Упор"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection1, 1 + "мм", true);

            //вторая фаска(меньше строк)
            EdgeCollection1.Add(revolvefeature3.Faces[3].Edges[2]);
            oCompDef["5. Упор"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection1, 1 + "мм", true);

            FilletFeature oFillet = default(FilletFeature);

            //скругления
            EdgeCollection1.Add(revolvefeature3.Faces[6].Edges[1]);
            FilletDefinition oFilletDef = oCompDef["5. Упор"].Features.FilletFeatures.CreateFilletDefinition();
            oFilletDef.AddConstantRadiusEdgeSet(EdgeCollection1, 0.5 + "мм");
            oFillet = oCompDef["5. Упор"].Features.FilletFeatures.Add(oFilletDef, false);

            EdgeCollection1.Add(revolvefeature3.Faces[5].Edges[1]);
            FilletDefinition oFilletDef1 = oCompDef["5. Упор"].Features.FilletFeatures.CreateFilletDefinition();
            oFilletDef1.AddConstantRadiusEdgeSet(EdgeCollection1, 1 + "мм");
            oFillet = oCompDef["5. Упор"].Features.FilletFeatures.Add(oFilletDef1, false);
            oPartDoc["5. Упор"].SaveAs(FBD.SelectedPath + "\\5. Упор.ipt", false);

            //Построение детали 10.Тарелка
            Имя_нового_документа("10. Тарелка");
            oPartDoc["10. Тарелка"].DisplayName = "10. Тарелка";
            PlanarSketch oSketch5 = oCompDef["10. Тарелка"].Sketches.Add(oCompDef["10. Тарелка"].WorkPlanes[3]);
            point[0] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(0, 0), false);
            point[1] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(DTar, 0), false);
            point[2] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(DTar, HTar), false);
            point[3] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(d1Tar, HTar), false);
            point[4] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(d1Tar, HTar-hTar), false);
            point[5] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(d2Tar*1.2, HTar - hTar), false);
            point[6] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(d2Tar * 1.2, HTar), false);
            point[7] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(d2Tar, HTar), false);
            point[8] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(d2Tar, HTar - hTar), false);
            point[9] = oSketch5.SketchPoints.Add(oTransGeom["10. Тарелка"].CreatePoint2d(0, HTar - hTar), false);
            lines[0] = oSketch5.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch5.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch5.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch5.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch5.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch5.SketchLines.AddByTwoPoints(point[5], point[6]);
            lines[6] = oSketch5.SketchLines.AddByTwoPoints(point[6], point[7]);
            lines[7] = oSketch5.SketchLines.AddByTwoPoints(point[7], point[8]);
            lines[8] = oSketch5.SketchLines.AddByTwoPoints(point[8], point[9]);
            lines[9] = oSketch5.SketchLines.AddByTwoPoints(point[9], point[0]);

            oTrans["10. Тарелка"].End();

            Profile oProfile5 = (Profile)oSketch5.Profiles.AddForSolid();
            RevolveFeature revolvefeature4 = oCompDef["10. Тарелка"].Features.
            RevolveFeatures.AddFull(oProfile5, lines[9],
            PartFeatureOperationEnum.kJoinOperation);

            //фаски
            EdgeCollection EdgeCollection2 = ThisApplication.TransientObjects.CreateEdgeCollection();
            EdgeCollection2.Add(revolvefeature4.Faces[1].Edges[1]);
            EdgeCollection2.Add(revolvefeature4.Faces[2].Edges[1]);
            EdgeCollection2.Add(revolvefeature4.Faces[5].Edges[1]);
            EdgeCollection2.Add(revolvefeature4.Faces[7].Edges[2]);
            oCompDef["10. Тарелка"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection2, 0.5 + "мм", true);
            oPartDoc["10. Тарелка"].SaveAs(FBD.SelectedPath + "\\10. Тарелка.ipt", false);

            //Построение детали 11. Диафрагма
            Имя_нового_документа("11. Диафрагма");
            oPartDoc["11. Диафрагма"].DisplayName = "11. Диафрагма";
            PlanarSketch oSketch6 = oCompDef["11. Диафрагма"].Sketches.Add(oCompDef["11. Диафрагма"].WorkPlanes[3]);
            point[0] = oSketch6.SketchPoints.Add(oTransGeom["11. Диафрагма"].CreatePoint2d(0, 0), false);
            point[1] = oSketch6.SketchPoints.Add(oTransGeom["11. Диафрагма"].CreatePoint2d(DFlancaVnutr, 0), false);
            point[2] = oSketch6.SketchPoints.Add(oTransGeom["11. Диафрагма"].CreatePoint2d(DFlancaVnutr, 0.11), false);
            point[3] = oSketch6.SketchPoints.Add(oTransGeom["11. Диафрагма"].CreatePoint2d(0, 0.11), false);
            lines[0] = oSketch6.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch6.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch6.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch6.SketchLines.AddByTwoPoints(point[3], point[0]);
            oTrans["11. Диафрагма"].End();
            Profile oProfile6 = (Profile)oSketch6.Profiles.AddForSolid();
            RevolveFeature revolvefeature5 = oCompDef["11. Диафрагма"].Features.
            RevolveFeatures.AddFull(oProfile6, lines[3],
            PartFeatureOperationEnum.kJoinOperation);
            oPartDoc["11. Диафрагма"].SaveAs(FBD.SelectedPath + "\\11. Диафрагма.ipt", false);

            //Построение детали 12. Кольцо 80х70
            Имя_нового_документа("12. Кольцо 80х70");
            oPartDoc["12. Кольцо 80х70"].DisplayName = "12. Кольцо 80х70";
            PlanarSketch oSketch7 = oCompDef["12. Кольцо 80х70"].Sketches.Add(oCompDef["12. Кольцо 80х70"].WorkPlanes[3]);
            point[0] = oSketch7.SketchPoints.Add(oTransGeom["12. Кольцо 80х70"].CreatePoint2d(DKolca - DKolco / 2, 0), false);
            point[1] = oSketch7.SketchPoints.Add(oTransGeom["12. Кольцо 80х70"].CreatePoint2d(0, 0), false);
            point[2] = oSketch7.SketchPoints.Add(oTransGeom["12. Кольцо 80х70"].CreatePoint2d(0, 2), false);
            lines[0] = oSketch7.SketchLines.AddByTwoPoints(point[1], point[2]);
            Окружность[0] = oSketch7.SketchCircles.AddByCenterRadius(point[0], DKolco);
            oTrans["12. Кольцо 80х70"].End();
            Profile oProfile7 = (Profile)oSketch7.Profiles.AddForSolid();
            RevolveFeature revolvefeature6 = oCompDef["12. Кольцо 80х70"].Features.
            RevolveFeatures.AddFull(oProfile7, lines[0],
            PartFeatureOperationEnum.kJoinOperation);
            oPartDoc["12. Кольцо 80х70"].SaveAs(FBD.SelectedPath + "\\12. Кольцо 80х70.ipt", false);
            //Построение детали 14. Втулка
            Имя_нового_документа("14. Втулка");
            oPartDoc["14. Втулка"].DisplayName = "14. Втулка";
            PlanarSketch oSketch8 = oCompDef["14. Втулка"].Sketches.Add(oCompDef["14. Втулка"].WorkPlanes[3]);
            point[0] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(DShtok, 0), false);
            point[1] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(RVtullka, 0), false);
            point[2] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(RVtullka, 0.3), false);
            point[3] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(RVtullkaSmall * 0.9, 0.3), false);
            point[4] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(RVtullkaSmall * 0.9, 0.4), false);
            point[5] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(RVtullkaSmall, 0.4), false);
            point[6] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(RVtullkaSmall, HVtulka), false);
            point[7] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(DShtok, HVtulka), false);
            point[8] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(0, 0), false);
            point[9] = oSketch8.SketchPoints.Add(oTransGeom["14. Втулка"].CreatePoint2d(0, 2), false);
            lines[0] = oSketch8.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch8.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch8.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch8.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch8.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch8.SketchLines.AddByTwoPoints(point[5], point[6]);
            lines[6] = oSketch8.SketchLines.AddByTwoPoints(point[6], point[7]);
            lines[7] = oSketch8.SketchLines.AddByTwoPoints(point[7], point[0]);
            lines[8] = oSketch8.SketchLines.AddByTwoPoints(point[8], point[9]);
            oTrans["14. Втулка"].End();
            Profile oProfile8 = (Profile)oSketch8.Profiles.AddForSolid();
            RevolveFeature revolvefeature7 = oCompDef["14. Втулка"].Features.
            RevolveFeatures.AddFull(oProfile8, lines[8],
            PartFeatureOperationEnum.kJoinOperation);
            //фаски
            EdgeCollection EdgeCollection3 = ThisApplication.TransientObjects.CreateEdgeCollection();
            EdgeCollection3.Add(revolvefeature7.Faces[3].Edges[1]);
            EdgeCollection3.Add(revolvefeature7.Faces[7].Edges[2]);
            oCompDef["14. Втулка"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection3, 0.2 + "мм", true);
            EdgeCollection EdgeCollection4 = ThisApplication.TransientObjects.CreateEdgeCollection();
            EdgeCollection4.Add(revolvefeature7.Faces[3].Edges[1]);
            oCompDef["14. Втулка"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection4, 0.5 + "мм", true);
            oPartDoc["14. Втулка"].SaveAs(FBD.SelectedPath + "\\14. Втулка.ipt", false);
            //Построение детали 15. Шток
            Имя_нового_документа("15. Шток");
            oPartDoc["15. Шток"].DisplayName = "15. Шток";
            PlanarSketch oSketch9 = oCompDef["15. Шток"].Sketches.Add(oCompDef["15. Шток"].WorkPlanes[3]);
            point[0] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(0, 0), false);
            point[1] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(DShtok, 0), false);
            point[2] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(DShtok, LShtokRuchka * 0.81), false);
            point[3] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(DShtok * 1.4, LShtokRuchka * 0.81), false);
            point[4] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(DShtok * 1.4, LShtokRuchka), false);
            point[5] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(DShtok * 2.8, LShtokRuchka), false);
            point[6] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(DShtok * 2.8, LShtokRuchka * 1.03), false);
            point[7] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(DShtok, LShtokRuchka * 1.33), false);
            point[8] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(DShtok, LShtok), false);
            point[9] = oSketch9.SketchPoints.Add(oTransGeom["15. Шток"].CreatePoint2d(0, LShtok), false);
            lines[0] = oSketch9.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch9.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch9.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch9.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch9.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch9.SketchLines.AddByTwoPoints(point[5], point[6]);
            lines[6] = oSketch9.SketchLines.AddByTwoPoints(point[6], point[7]);
            lines[7] = oSketch9.SketchLines.AddByTwoPoints(point[7], point[8]);
            lines[8] = oSketch9.SketchLines.AddByTwoPoints(point[8], point[9]);
            lines[9] = oSketch9.SketchLines.AddByTwoPoints(point[9], point[0]);
            oTrans["15. Шток"].End();
            Profile oProfile9 = (Profile)oSketch9.Profiles.AddForSolid();
            RevolveFeature revolvefeature8 = oCompDef["15. Шток"].Features.
            RevolveFeatures.AddFull(oProfile9, lines[9],
            PartFeatureOperationEnum.kJoinOperation);

            //скругления
            EdgeCollection EdgeCollection5 = ThisApplication.TransientObjects.CreateEdgeCollection();
            FilletFeature oFillet1 = default(FilletFeature);
            EdgeCollection5.Add(revolvefeature8.Faces[7].Edges[1]);
            FilletDefinition oFilletDef2 = oCompDef["15. Шток"].Features.FilletFeatures.CreateFilletDefinition();
            oFilletDef2.AddConstantRadiusEdgeSet(EdgeCollection5, DShtok + "см");
            oFillet1 = oCompDef["15. Шток"].Features.FilletFeatures.Add(oFilletDef2, false);
            oPartDoc["15. Шток"].SaveAs(FBD.SelectedPath + "\\15. Шток.ipt", false);
            //Построение детали 16. Прокладка
            Имя_нового_документа("16. Прокладка");
            oPartDoc["16. Прокладка"].DisplayName = "16. Прокладка";
            PlanarSketch oSketch10 = oCompDef["16. Прокладка"].Sketches.Add(oCompDef["16. Прокладка"].WorkPlanes[3]);
            point[0] = oSketch10.SketchPoints.Add(oTransGeom["16. Прокладка"].CreatePoint2d(DShtok * 1.5, 0), false);
            point[1] = oSketch10.SketchPoints.Add(oTransGeom["16. Прокладка"].CreatePoint2d(DGasPr * 1.3, 0), false);
            point[2] = oSketch10.SketchPoints.Add(oTransGeom["16. Прокладка"].CreatePoint2d(DGasPr * 1.3, LShtokRuchka - LShtokRuchka * 0.81 - 0.03), false);
            point[3] = oSketch10.SketchPoints.Add(oTransGeom["16. Прокладка"].CreatePoint2d(DShtok * 1.5, LShtokRuchka - LShtokRuchka * 0.81 - 0.03), false);
            point[4] = oSketch10.SketchPoints.Add(oTransGeom["16. Прокладка"].CreatePoint2d(0, 0), false);
            point[5] = oSketch10.SketchPoints.Add(oTransGeom["16. Прокладка"].CreatePoint2d(0, 2), false);
            lines[0] = oSketch10.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch10.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch10.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch10.SketchLines.AddByTwoPoints(point[3], point[0]);
            lines[4] = oSketch10.SketchLines.AddByTwoPoints(point[4], point[5]);
            oTrans["16. Прокладка"].End();
            Profile oProfile10 = (Profile)oSketch10.Profiles.AddForSolid();
            RevolveFeature revolvefeature9 = oCompDef["16. Прокладка"].Features.
            RevolveFeatures.AddFull(oProfile10, lines[4],
            PartFeatureOperationEnum.kJoinOperation);
            oPartDoc["16. Прокладка"].SaveAs(FBD.SelectedPath + "\\16. Прокладка.ipt", false);
            //Построение детали 17. Кольцо 55х48
            Имя_нового_документа("17. Кольцо 55х48");
            oPartDoc["17. Кольцо 55х48"].DisplayName = "17. Кольцо 55х48";
            PlanarSketch oSketch11 = oCompDef["17. Кольцо 55х48"].Sketches.Add(oCompDef["17. Кольцо 55х48"].WorkPlanes[3]);
            point[0] = oSketch11.SketchPoints.Add(oTransGeom["17. Кольцо 55х48"].CreatePoint2d(DKolca17, 0), false);
            point[1] = oSketch11.SketchPoints.Add(oTransGeom["17. Кольцо 55х48"].CreatePoint2d(0, 0), false);
            point[2] = oSketch11.SketchPoints.Add(oTransGeom["17. Кольцо 55х48"].CreatePoint2d(0, 2), false);
            lines[0] = oSketch11.SketchLines.AddByTwoPoints(point[1], point[2]);
            Окружность[0] = oSketch11.SketchCircles.AddByCenterRadius(point[0], dKolca17);
            oTrans["17. Кольцо 55х48"].End();
            Profile oProfile11 = (Profile)oSketch11.Profiles.AddForSolid();
            RevolveFeature revolvefeature10 = oCompDef["17. Кольцо 55х48"].Features.
            RevolveFeatures.AddFull(oProfile11, lines[0],
            PartFeatureOperationEnum.kJoinOperation);
            oPartDoc["17. Кольцо 55х48"].SaveAs(FBD.SelectedPath + "\\17. Кольцо 55х48.ipt", false);

            //Построение детали 19. Клапан
            Имя_нового_документа("19. Клапан");
            oPartDoc["19. Клапан"].DisplayName = "19. Клапан";
            PlanarSketch oSketch12 = oCompDef["19. Клапан"].Sketches.Add(oCompDef["19. Клапан"].WorkPlanes[3]);
            point[0] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(DGasPr * 1.3, 0), false);
            point[1] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(DGasPr * 1.5, 0), false);
            point[2] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(DGasPr * 1.5, 0.5), false);
            point[3] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(DGasPr * 0.6, 0.5), false);
            point[4] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(DGasPr * 0.6, 1.4), false);
            point[5] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(DShtok, 1.4), false);
            point[6] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(DShtok, 0.3), false);
            point[7] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(DGasPr * 1.3, 0.3), false);
            point[8] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(0, 0), false);
            point[9] = oSketch12.SketchPoints.Add(oTransGeom["19. Клапан"].CreatePoint2d(0, 2), false);
            lines[0] = oSketch12.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch12.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch12.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch12.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch12.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch12.SketchLines.AddByTwoPoints(point[5], point[6]);
            lines[6] = oSketch12.SketchLines.AddByTwoPoints(point[6], point[7]);
            lines[7] = oSketch12.SketchLines.AddByTwoPoints(point[7], point[0]);
            lines[8] = oSketch12.SketchLines.AddByTwoPoints(point[8], point[9]);
            oTrans["19. Клапан"].End();
            Profile oProfile12 = (Profile)oSketch12.Profiles.AddForSolid();
            RevolveFeature revolvefeature11 = oCompDef["19. Клапан"].Features.
            RevolveFeatures.AddFull(oProfile12, lines[8],
            PartFeatureOperationEnum.kJoinOperation);
            //фаски
            EdgeCollection EdgeCollection6 = ThisApplication.TransientObjects.CreateEdgeCollection();
            EdgeCollection6.Add(revolvefeature11.Faces[3].Edges[1]);
            EdgeCollection6.Add(revolvefeature11.Faces[7].Edges[1]);
            EdgeCollection6.Add(revolvefeature11.Faces[7].Edges[2]);
            oCompDef["19. Клапан"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection6, 0.5 + "мм", true);
            EdgeCollection EdgeCollection7 = ThisApplication.TransientObjects.CreateEdgeCollection();
            EdgeCollection7.Add(revolvefeature11.Faces[7].Edges[2]);
            oCompDef["19. Клапан"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection7, 1 + "мм", true);
            FilletFeature oFillet2 = default(FilletFeature);
            EdgeCollection EdgeCollection8 = ThisApplication.TransientObjects.CreateEdgeCollection();
            EdgeCollection8.Add(revolvefeature11.Faces[2].Edges[1]);
            FilletDefinition oFilletDef3 = oCompDef["19. Клапан"].Features.FilletFeatures.CreateFilletDefinition();
            oFilletDef3.AddConstantRadiusEdgeSet(EdgeCollection8, 0.5 + "мм");
            oFillet2 = oCompDef["19. Клапан"].Features.FilletFeatures.Add(oFilletDef3, false);
            oPartDoc["19. Клапан"].SaveAs(FBD.SelectedPath + "\\19. Клапан.ipt", false);
            //Построение детали 20. Пробка
            Имя_нового_документа("20. Пробка");
            oPartDoc["20. Пробка"].DisplayName = "20. Пробка";
            PlanarSketch oSketch13 = oCompDef["20. Пробка"].Sketches.Add(oCompDef["20. Пробка"].WorkPlanes[3]);
            point[0] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(0, 0), false);
            point[1] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DOsn * 1.01, 0), false);
            point[2] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DOsn * 1.16, (HProbka - HRezbProbka) * 0.215), false);
            point[3] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DOsn * 1.16, (HProbka - HRezbProbka) * 0.785), false);
            point[4] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DOsn * 1.01, HProbka - HRezbProbka), false);
            point[5] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKolca17-dKolca17+0.05, HProbka - HRezbProbka), false);
            point[6] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKolca17 - dKolca17 + 0.05, HProbka - HRezbProbka + 0.05), false);
            point[7] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKolca17 - dKolca17, HProbka - HRezbProbka + 0.05), false);
            point[8] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKolca17 - dKolca17, 0.953), false);
            //point[9] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(2.394, 0.853), false);
            //point[10] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(2.323, 0.923), false);
            point[11] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKlap, 1), false);
            point[12] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKlap, HProbka - 0.15), false);
            point[13] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKlap - 0.15, HProbka), false);
            point[14] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKlap * 0.73 + 0.05, HProbka), false);
            point[15] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKlap * 0.73, HProbka - 0.05), false);
            point[16] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DKlap * 0.73, HProbka * 0.6), false);
            point[17] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(1, HProbka * 0.38), false);
            point[18] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(1, 0.2), false);
            point[19] = oSketch13.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(0, 0.2), false);
            lines[0] = oSketch13.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch13.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch13.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch13.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch13.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch13.SketchLines.AddByTwoPoints(point[7], point[8]);
            lines[6] = oSketch13.SketchLines.AddByTwoPoints(point[8], point[11]);
            lines[7] = oSketch13.SketchLines.AddByTwoPoints(point[11], point[12]);
            lines[8] = oSketch13.SketchLines.AddByTwoPoints(point[12], point[13]);
            lines[9] = oSketch13.SketchLines.AddByTwoPoints(point[13], point[14]);
            lines[10] = oSketch13.SketchLines.AddByTwoPoints(point[14], point[15]);
            lines[11] = oSketch13.SketchLines.AddByTwoPoints(point[15], point[16]);
            lines[12] = oSketch13.SketchLines.AddByTwoPoints(point[16], point[17]);
            lines[13] = oSketch13.SketchLines.AddByTwoPoints(point[17], point[18]);
            lines[14] = oSketch13.SketchLines.AddByTwoPoints(point[18], point[19]);
            lines[15] = oSketch13.SketchLines.AddByTwoPoints(point[19], point[0]);
            arcs[0] = oSketch13.SketchArcs.AddByCenterStartEndPoint(oTransGeom["20. Пробка"].CreatePoint2d(
            point[6].Geometry.X, point[6].Geometry.Y), point[7], point[5]);
            oTrans["20. Пробка"].End();
            Profile oProfile13 = (Profile)oSketch13.Profiles.AddForSolid();
            RevolveFeature revolvefeature12 = oCompDef["20. Пробка"].Features.
            RevolveFeatures.AddFull(oProfile13, lines[15],
            PartFeatureOperationEnum.kJoinOperation);
            EdgeCollection EdgeCollection9 = ThisApplication.TransientObjects.CreateEdgeCollection();
            Face FacePr = revolvefeature12.SideFaces[8];
            Edge StartEdge_3 = FacePr.Edges[2];
            EdgeCollection9.Add(StartEdge_3);
            ThreadFeatures ThreadFeatures1 = oCompDef["20. Пробка"].Features.ThreadFeatures;
            StandardThreadInfo stInfo2 = ThreadFeatures1.CreateStandardThreadInfo(false, true,
            "ISO Metric profile", "M48x1.5", "6g");
            ThreadInfo ThreadInfo1 = (ThreadInfo)stInfo2;
            ThreadFeatures1.Add(FacePr, StartEdge_3, ThreadInfo1, false, false, 0.79 * HRezbProbka + "см", 0);

            WorkPlane oWorkPlane1 = oCompDef["20. Пробка"].WorkPlanes.AddByPlaneAndOffset(
            oCompDef["20. Пробка"].WorkPlanes[2], 6 + " мм", false);
            oWorkPlane1.Visible = false;
            //Выбор рабочей плоскости oWorkPlane и создание эскиза на плоскости "oSketch14"
            PlanarSketch oSketch14 = oCompDef["20. Пробка"].Sketches.Add(oWorkPlane1, false);
            oTrans["20. Пробка"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch14.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DOsn * 1.01, DOsn * 0.58), false);
            point[1] = oSketch14.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(0, DOsn * 1.16), false);
            point[2] = oSketch14.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(-DOsn * 1.01, DOsn * 0.58), false);
            point[3] = oSketch14.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(-DOsn * 1.01, -DOsn * 0.58), false);
            point[4] = oSketch14.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(0, -DOsn * 1.16), false);
            point[5] = oSketch14.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(DOsn * 1.01, -DOsn * 0.58), false);
            point[6] = oSketch14.SketchPoints.Add(oTransGeom["20. Пробка"].CreatePoint2d(0, 0), false);
            lines[0] = oSketch14.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch14.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch14.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch14.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch14.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[6] = oSketch14.SketchLines.AddByTwoPoints(point[5], point[0]);
            Окружность[0] = oSketch14.SketchCircles.AddByCenterRadius(point[6], 7.2);
            oTrans["20. Пробка"].End();
            Profile oProfile14 = (Profile)oSketch14.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef1 = oCompDef["20. Пробка"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile14,/*Длина в см*/HProbka - HRezbProbka,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kCutOperation,/*Эскиз*/oProfile14);
            oPartDoc["20. Пробка"].SaveAs(FBD.SelectedPath + "\\20. Пробка.ipt", false);

            //Построение детали 21. Шайба
            Имя_нового_документа("21. Шайба");
            oPartDoc["21. Шайба"].DisplayName = "21. Шайба";
            PlanarSketch oSketch15 = oCompDef["21. Шайба"].Sketches.Add(oCompDef["21. Шайба"].WorkPlanes[3]);
            point[0] = oSketch15.SketchPoints.Add(oTransGeom["21. Шайба"].CreatePoint2d(0, 0), false);
            point[1] = oSketch15.SketchPoints.Add(oTransGeom["21. Шайба"].CreatePoint2d(0, 2), false);
            point[2] = oSketch15.SketchPoints.Add(oTransGeom["21. Шайба"].CreatePoint2d(DShtok * 1.4, 0), false);
            point[3] = oSketch15.SketchPoints.Add(oTransGeom["21. Шайба"].CreatePoint2d(DGasPr * 0.9, 0), false);
            point[4] = oSketch15.SketchPoints.Add(oTransGeom["21. Шайба"].CreatePoint2d(DGasPr * 0.9, 0.03), false);
            point[5] = oSketch15.SketchPoints.Add(oTransGeom["21. Шайба"].CreatePoint2d(DShtok * 1.4, 0.03), false);
            lines[0] = oSketch15.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch15.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[2] = oSketch15.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[3] = oSketch15.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[4] = oSketch15.SketchLines.AddByTwoPoints(point[5], point[2]);
            oTrans["21. Шайба"].End();
            Profile oProfile15 = (Profile)oSketch15.Profiles.AddForSolid();
            RevolveFeature revolvefeature13 = oCompDef["21. Шайба"].Features.
            RevolveFeatures.AddFull(oProfile15, lines[0],
            PartFeatureOperationEnum.kJoinOperation);
            oPartDoc["21. Шайба"].SaveAs(FBD.SelectedPath + "\\21. Шайба.ipt", false);
            //Построение детали 22. Угольник
            Имя_нового_документа("22. Угольник");
            oPartDoc["22. Угольник"].DisplayName = "22. Угольник";
            PlanarSketch oSketch16 = oCompDef["22. Угольник"].Sketches.Add(oCompDef["22. Угольник"].WorkPlanes[3]);
            point[0] = oSketch16.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(HYgol / 2, ShirYgol / 2), false);
            point[1] = oSketch16.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(-HYgol / 2, ShirYgol / 2), false);
            point[2] = oSketch16.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(-HYgol / 2, -ShirYgol / 2), false);
            point[3] = oSketch16.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(HYgol / 2, -ShirYgol / 2), false);
            lines[0] = oSketch16.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch16.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch16.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch16.SketchLines.AddByTwoPoints(point[3], point[0]);
            oTrans["22. Угольник"].End();
            Profile oProfile16 = (Profile)oSketch16.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef2 = oCompDef["22. Угольник"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile16,/*Длина в см*/0.9,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kJoinOperation,/*Эскиз*/oProfile16);
            ExtrudeFeature oExtrudeDef3 = oCompDef["22. Угольник"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile16,/*Длина в см*/0.9,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kJoinOperation,/*Эскиз*/oProfile16);

            PlanarSketch oSketch17 = oCompDef["22. Угольник"].Sketches.Add(oCompDef["22. Угольник"].WorkPlanes[3]);
            oTrans["22. Угольник"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch17.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(-HYgol / 2, ShirYgol * 0.86 / 2), false);
            point[1] = oSketch17.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(-HYgol / 2 * 0.84, ShirYgol / 2 * 0.67), false);
            point[2] = oSketch17.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(HYgol / 2 * 0.64, ShirYgol / 2 * 0.67), false);
            point[3] = oSketch17.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(HYgol / 2 * 0.96, 0), false);
            point[4] = oSketch17.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(-HYgol / 2, 0), false);
            lines[0] = oSketch17.SketchLines.AddByTwoPoints(point[4], point[0]);
            lines[1] = oSketch17.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[2] = oSketch17.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[3] = oSketch17.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[4] = oSketch17.SketchLines.AddByTwoPoints(point[3], point[4]);
            oTrans["22. Угольник"].End();
            Profile oProfile17 = (Profile)oSketch17.Profiles.AddForSolid();
            RevolveFeature revolvefeature14 = oCompDef["22. Угольник"].Features.
            RevolveFeatures.AddFull(oProfile17, lines[4],
            PartFeatureOperationEnum.kCutOperation);

            EdgeCollection EdgeCollection10 = ThisApplication.TransientObjects.CreateEdgeCollection();
            Face Face2 = revolvefeature14.SideFaces[2];
            Edge StartEdge_4 = Face2.Edges[1];
            EdgeCollection10.Add(StartEdge_4);
            ThreadFeatures ThreadFeatures2 = oCompDef["22. Угольник"].Features.ThreadFeatures;
            StandardThreadInfo stInfo3 = ThreadFeatures2.CreateStandardThreadInfo(false, true,
            "ISO Metric profile", "M12x1.5", "6g");
            ThreadInfo ThreadInfo2 = (ThreadInfo)stInfo3;
            ThreadFeatures2.Add(Face2, StartEdge_4, ThreadInfo2, false, false, HYgol * 0.46 + "см", 0);

            //фаски
            EdgeCollection EdgeCollection11 = ThisApplication.TransientObjects.CreateEdgeCollection();
            EdgeCollection11.Add(oExtrudeDef2.Faces[1].Edges[1]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[1].Edges[2]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[1].Edges[3]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[1].Edges[4]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[2].Edges[1]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[2].Edges[2]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[2].Edges[3]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[2].Edges[4]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[3].Edges[1]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[3].Edges[2]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[3].Edges[3]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[3].Edges[4]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[4].Edges[1]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[4].Edges[2]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[4].Edges[3]);
            EdgeCollection11.Add(oExtrudeDef2.Faces[4].Edges[4]);
            oCompDef["22. Угольник"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection11, 1 + "мм", true);


            WorkPlane oWorkPlane2 = oCompDef["22. Угольник"].WorkPlanes.AddByPlaneAndOffset(
            oCompDef["22. Угольник"].WorkPlanes[3], ShirYgol / 2 + " см", false);
            oWorkPlane2.Visible = false;
            //Выбор рабочей плоскости oWorkPlane и создание эскиза на плоскости "oSketch14"
            PlanarSketch oSketch18 = oCompDef["22. Угольник"].Sketches.Add(oWorkPlane2, false);
            oTrans["22. Угольник"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch18.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(HYgol / 2 * 0.46, 0), false);
            Окружность[0] = oSketch18.SketchCircles.AddByCenterRadius(point[0], 0.525);
            Окружность[1] = oSketch18.SketchCircles.AddByCenterRadius(point[0], 0.1);
            oTrans["22. Угольник"].End();

            Profile oProfile18 = (Profile)oSketch18.Profiles.AddForSolid();
            ExtrudeFeature oExtrudeDef4 = oCompDef["22. Угольник"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile18,/*Длина в см*/LTrYgol * 0.85,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kJoinOperation,/*Эскиз*/oProfile18);


            PlanarSketch oSketch19 = oCompDef["22. Угольник"].Sketches.Add(oCompDef["22. Угольник"].WorkPlanes[2]);
            oTrans["22. Угольник"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch19.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(0.025, LTrYgol * 0.85 + ShirYgol / 2), false);
            point[1] = oSketch19.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(-0.5, LTrYgol * 0.85 + ShirYgol / 2), false);
            point[2] = oSketch19.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(-0.5, LTrYgol + ShirYgol / 2), false);
            point[3] = oSketch19.SketchPoints.Add(oTransGeom["22. Угольник"].CreatePoint2d(-0.055, LTrYgol + ShirYgol / 2), false);
            lines[0] = oSketch19.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch19.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch19.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch19.SketchLines.AddByTwoPoints(point[3], point[0]);
            oTrans["22. Угольник"].End();
            Profile oProfile19 = (Profile)oSketch19.Profiles.AddForSolid();
            RevolveFeature revolvefeature15 = oCompDef["22. Угольник"].Features.
            RevolveFeatures.AddFull(oProfile19, lines[1],
            PartFeatureOperationEnum.kJoinOperation);

            //EdgeCollection EdgeCollection12 = ThisApplication.TransientObjects.CreateEdgeCollection();
            //Face Face3 = revolvefeature15.SideFaces[3];
            //Edge StartEdge_5 = Face3.Edges[1];
            //EdgeCollection12.Add(StartEdge_5);
            //ThreadFeatures ThreadFeatures3 = oCompDef["22. Угольник"].Features.ThreadFeatures;
            //StandardThreadInfo stInfo4 = ThreadFeatures3.CreateStandardThreadInfo(false, true,
            //"ISO Taper External", "R1/8", "");
            //ThreadInfo ThreadInfo3 = (ThreadInfo)stInfo4;
            //ThreadFeatures3.Add(Face2, StartEdge_5, ThreadInfo3, false, true);
            oPartDoc["22. Угольник"].SaveAs(FBD.SelectedPath + "\\22. Угольник.ipt", false);

            //Построение детали 23. Пробка
            Имя_нового_документа("23. Пробка");
            oPartDoc["23. Пробка"].DisplayName = "23. Пробка";
            PlanarSketch oSketch20 = oCompDef["23. Пробка"].Sketches.Add(oCompDef["23. Пробка"].WorkPlanes[3]);

            point[0] = oSketch20.SketchPoints.Add(oTransGeom["23. Пробка"].CreatePoint2d(0.075, 0), false);
            point[1] = oSketch20.SketchPoints.Add(oTransGeom["23. Пробка"].CreatePoint2d(0.515, 0), false);
            point[2] = oSketch20.SketchPoints.Add(oTransGeom["23. Пробка"].CreatePoint2d(0.425, 0.9), false);
            point[3] = oSketch20.SketchPoints.Add(oTransGeom["23. Пробка"].CreatePoint2d(0, 0.9), false);
            point[4] = oSketch20.SketchPoints.Add(oTransGeom["23. Пробка"].CreatePoint2d(0, 0.3), false);
            point[5] = oSketch20.SketchPoints.Add(oTransGeom["23. Пробка"].CreatePoint2d(0.075, 0.3), false);
            lines[0] = oSketch20.SketchLines.AddByTwoPoints(point[0], point[1]);
            lines[1] = oSketch20.SketchLines.AddByTwoPoints(point[1], point[2]);
            lines[2] = oSketch20.SketchLines.AddByTwoPoints(point[2], point[3]);
            lines[3] = oSketch20.SketchLines.AddByTwoPoints(point[3], point[4]);
            lines[4] = oSketch20.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[5] = oSketch20.SketchLines.AddByTwoPoints(point[5], point[0]);
            oTrans["23. Пробка"].End();
            Profile oProfile20 = (Profile)oSketch20.Profiles.AddForSolid();
            RevolveFeature revolvefeature16 = oCompDef["23. Пробка"].Features.
            RevolveFeatures.AddFull(oProfile20, lines[3],
            PartFeatureOperationEnum.kJoinOperation);

            EdgeCollection EdgeCollection13 = ThisApplication.TransientObjects.CreateEdgeCollection();
            EdgeCollection13.Add(revolvefeature16.Faces[1].Edges[2]);
            EdgeCollection13.Add(revolvefeature16.Faces[1].Edges[1]);
            oCompDef["23. Пробка"].Features.ChamferFeatures.AddUsingDistance(EdgeCollection13, 1 + "мм", true);

            //EdgeCollection EdgeCollection14 = ThisApplication.TransientObjects.CreateEdgeCollection();
            //Face Face4 = revolvefeature16.SideFaces[1];
            //Edge StartEdge_5 = Face4.Edges[1];
            //EdgeCollection14.Add(StartEdge_5);// SketchPoint[] point = new SketchPoint[20];
            ////ThreadFeatures ThreadFeatures3 = oCompDef["23. Пробка"].Features.ThreadFeatures;
            //ThreadTableQuery ThreadFeatures3 = oCompDef["23. Пробка"].Features.Application.GeneralOptions.ThreadTableQuery;
            ////GeneralOptions oGeneralOptions = ThisApplication.GeneralOptions;
            //////GeneralOptions oGeneralOptions = oCompDef["23. Пробка"].GeneralOptions;
            ////ThreadTableQuery oThreadTable = oGeneralOptions.ThreadTableQuery;
            ////String  oTypes;
            //////oTypes = oThreadTable.GetAvailableThreadTypes;
            //////oTypes[0] = oCompDef["23. Пробка"].GetAvailableThreadTypes();
            ////String oSize;
            ////oSize = oThreadTable.GetAvailableThreadSizes(false, "NPT");
            ////String oDestignation = (String)oThreadTable.GetAvailableDesignations(false, "NPT", oSize);
            ////oDestignation = 
            //TaperedThreadInfo stInfo3 = (TaperedThreadInfo)ThreadFeatures3.CreateThreadInfo(false, true, "NPT", "R1/8");
            //TaperedThreadInfo ThreadInfo3 = stInfo3;
            //ThreadFeatures3.CreateThreadInfo.ThreadInfo3(Face4, StartEdge_5, ThreadInfo3, false, true);
            oPartDoc["23. Пробка"].SaveAs(FBD.SelectedPath + "\\23. Пробка.ipt", false);

            //Создание детали 6. Пружина
            Имя_нового_документа("6. Пружина");
            oPartDoc["6. Пружина"].DisplayName = "6. Пружина";
            PlanarSketch oSketch35 = oCompDef["6. Пружина"].Sketches.Add(oCompDef["6. Пружина"].WorkPlanes[3]);
            point[0] = oSketch35.SketchPoints.Add(oTransGeom["6. Пружина"].CreatePoint2d(DPr6-dPr6, dPr6), false);
            Окружность[0] = oSketch35.SketchCircles.AddByCenterRadius(point[0], dPr6);
            oTrans["6. Пружина"].End();
            Profile oProfile35 = (Profile)oSketch35.Profiles.AddForSolid();
            CoilFeature coil = oCompDef["6. Пружина"].Features.
                CoilFeatures.AddByRevolutionAndHeight(oProfile35, oCompDef["6. Пружина"].WorkAxes[2], nPr6, HPr6-dPr6,
                    PartFeatureOperationEnum.kJoinOperation);
            PlanarSketch oSketch38 = oCompDef["6. Пружина"].Sketches.Add(oCompDef["6. Пружина"].WorkPlanes[2]);
            oTrans["6. Пружина"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch38.SketchPoints.Add(oTransGeom["6. Пружина"].CreatePoint2d(0, 0), false);
            Окружность[0] = oSketch38.SketchCircles.AddByCenterRadius(point[0], 0.01);
            Profile oProfile38 = (Profile)oSketch38.Profiles.AddForSolid();
            oTrans["6. Пружина"].End();
            ExtrudeFeature oExtrudeDef13 = oCompDef["6. Пружина"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile38,/*Длина в см*/0.01,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kJoinOperation,/*Эскиз*/oProfile38);
            oPartDoc["6. Пружина"].SaveAs(FBD.SelectedPath + "\\6. Пружина.ipt", false);

            //Создание детали 7. Пружина
            Имя_нового_документа("7. Пружина");
            oPartDoc["7. Пружина"].DisplayName = "7. Пружина";
            PlanarSketch oSketch36 = oCompDef["7. Пружина"].Sketches.Add(oCompDef["7. Пружина"].WorkPlanes[3]);
            point[0] = oSketch36.SketchPoints.Add(oTransGeom["7. Пружина"].CreatePoint2d(DPr7-dPr7, dPr7), false);
            Окружность[0] = oSketch36.SketchCircles.AddByCenterRadius(point[0], dPr7);
            oTrans["7. Пружина"].End();
            Profile oProfile36 = (Profile)oSketch36.Profiles.AddForSolid();
            CoilFeature coil1 = oCompDef["7. Пружина"].Features.
                CoilFeatures.AddByRevolutionAndHeight(oProfile36, oCompDef["7. Пружина"].WorkAxes[2], nPr7, HPr7-dPr7,
                    PartFeatureOperationEnum.kJoinOperation, ClockwiseRotation: true, FlatStartType: true, StartTransitionAngle: 90 + "degree",
                    StartFlatAngle: 0 + "degree", FlatEndType: true, EndTransitionAngle: 90 + "degree", EndFlatAngle: 90 + "degree");
            PlanarSketch oSketch39 = oCompDef["7. Пружина"].Sketches.Add(oCompDef["7. Пружина"].WorkPlanes[2]);
            oTrans["7. Пружина"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch39.SketchPoints.Add(oTransGeom["7. Пружина"].CreatePoint2d(0, 0), false);
            Окружность[0] = oSketch39.SketchCircles.AddByCenterRadius(point[0], 0.01);
            Profile oProfile39 = (Profile)oSketch39.Profiles.AddForSolid();
            oTrans["7. Пружина"].End();
            ExtrudeFeature oExtrudeDef14 = oCompDef["7. Пружина"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile39,/*Длина в см*/0.01,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kJoinOperation,/*Эскиз*/oProfile39);
            oPartDoc["7. Пружина"].SaveAs(FBD.SelectedPath + "\\7. Пружина.ipt", false);
            //Создание детали 18. Пружина
            Имя_нового_документа("18. Пружина");
            oPartDoc["18. Пружина"].DisplayName = "18. Пружина";
            PlanarSketch oSketch37 = oCompDef["18. Пружина"].Sketches.Add(oCompDef["18. Пружина"].WorkPlanes[3]);
            point[0] = oSketch37.SketchPoints.Add(oTransGeom["18. Пружина"].CreatePoint2d(DPr18-dPr18, dPr18), false);
            Окружность[0] = oSketch37.SketchCircles.AddByCenterRadius(point[0], dPr18);
            oTrans["18. Пружина"].End();
            Profile oProfile37 = (Profile)oSketch37.Profiles.AddForSolid();
            CoilFeature coil2 = oCompDef["18. Пружина"].Features.
                CoilFeatures.AddByRevolutionAndHeight(oProfile37, oCompDef["18. Пружина"].WorkAxes[2], nPr18, HPr18- dPr18,
                    PartFeatureOperationEnum.kJoinOperation);
            PlanarSketch oSketch40 = oCompDef["18. Пружина"].Sketches.Add(oCompDef["18. Пружина"].WorkPlanes[2]);
            oTrans["18. Пружина"] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            point[0] = oSketch40.SketchPoints.Add(oTransGeom["18. Пружина"].CreatePoint2d(0, 0), false);
            Окружность[0] = oSketch40.SketchCircles.AddByCenterRadius(point[0], 0.01);
            Profile oProfile40 = (Profile)oSketch40.Profiles.AddForSolid();
            oTrans["18. Пружина"].End();
            ExtrudeFeature oExtrudeDef15 = oCompDef["18. Пружина"].Features.ExtrudeFeatures.AddByDistanceExtent(
            /*Эскиз*/oProfile40,/*Длина в см*/0.01,/*Направление вдоль оси*/
            PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
            /*Операция*/PartFeatureOperationEnum.kJoinOperation,/*Эскиз*/oProfile40);
            oPartDoc["18. Пружина"].SaveAs(FBD.SelectedPath + "\\18. Пружина.ipt", false);
            MessageBox.Show("Создание деталей завершено!", "Сообщение");
            //Перевод размеров в мм
            DGasPr = DGasPr * 2 * 10;
            DOsn = DOsn * 2 * 10;
            DKlap = DKlap * 2 * 10;
            RVtullka = RVtullka * 10;
            RVtullkaSmall = RVtullkaSmall * 10;
            DKolca = DKolca * 2 * 10;
            DFlancaVnesh = DFlancaVnesh * 2 * 10;
            DFlancaVnutr = DFlancaVnutr * 2 * 10;
            HFlanca = HFlanca * 10;
            HKorp = HKorp * 10;
            DOtvTr = DOtvTr * 2 * 10;
            DOtvBolt = DOtvBolt * 2 * 10;
            HFlancaKr = HFlancaKr * 10;
            DVint = DVint * 2 * 10;
            ShirStKr = ShirStKr * 10;
            HKr = HKr * 10;
            HGorlKr = HGorlKr * 10;
            HVint = HVint * 10;
            DShtift = DShtift * 2 * 10;
            HShtift = HShtift * 10;
            HVtulka = HVtulka * 10;
            HYgol = HYgol * 10;
            ShirYgol = ShirYgol * 10;
            LYgol = LYgol * 10;
            LTrYgol = LTrYgol * 10;
            DKolco = DKolco * 10 * 2;
            HProbka = HProbka * 10;
            HRezbProbka = HRezbProbka * 10;
            DShtok = DShtok * 2 * 10;
            LShtok = LShtok * 10;
            LShtokRuchka = LShtokRuchka * 10;
            HPr6 = HPr6 * 10;
            DPr6 = DPr6 * 2 * 10;
            dPr6 = dPr6 * 2 * 10;
            HPr7 = HPr7 * 10;
            DPr7 = DPr7 * 2 * 10;
            dPr7 = dPr7 * 2 * 10;
            HPr18 = HPr18 * 10;
            DPr18 = DPr18 * 2 * 10;
            dPr18 = dPr18 * 2 * 10;
            DTar = DTar * 2 * 10;
            HTar = HTar * 10;
            hTar = hTar * 10;
            d1Tar = d1Tar * 2 * 10;
            d2Tar = d2Tar * 2 * 10;
            DKolca17 = DKolca17 * 2 * 10;
            dKolca17 = dKolca17 * 2 * 10;
            DYpor = DYpor * 2 * 10;
            HYpor = HYpor * 10;
            d1Ypor = d1Ypor * 2 * 10;
            d2Ypor = d2Ypor * 2 * 10;
            h1Ypor = h1Ypor * 10;
            h2Ypor = h2Ypor * 10;
            DOpor = DOpor * 10 * 2;
            HOpor = HOpor * 10;
            d1Opor = d1Opor * 10 * 2;
            h1Opor = h1Opor * 10;
            h2Opor = h2Opor * 10;
        }

        


        private void Save_Model(string oPartDocName, string Text)
        {
            saveFileDialog1.Filter = "Inventor Part Document|*.ipt";
            saveFileDialog1.Title = Text;
            saveFileDialog1.FileName = oPartDoc[oPartDocName].DisplayName;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrWhiteSpace(saveFileDialog1.FileName))
                {
                    oPartDoc[oPartDocName].SaveAs(saveFileDialog1.FileName, false);
                    oFileName[oPartDocName] = saveFileDialog1.FileName;
                }
            }
        }
    }
}

