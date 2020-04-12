using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelAPI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Provo a scrivere in un file excel
            //Apro il file excel
            //OpenFile();
            // WriteData();

            /* Rendo protetto con password un file excel
            Excel ex = new Excel(@"C:\Users\fabio\source\repos\PianificazioneScadenze\PianificazioneScadenze\bin\Debug\prova2.xlsx", 1);
            // ex.ProtectSheet("password");
            ex.UnprotectSheet("password");
            ex.SaveAs(@"C:\Users\fabio\source\repos\PianificazioneScadenze\PianificazioneScadenze\bin\Debug\prova2.xlsx");
            ex.Close();
            /*
            Excel ex = new Excel(@"C:\Users\fabio\source\repos\PianificazioneScadenze\PianificazioneScadenze\bin\Debug\prova.xlsx", 1);
            ex.SelectWorksheet(2);
            ex.WriteToCell(1, 1, "this is sheet two");
            ex.DeleteWorksheet(1);
            ex.SaveAs(@"C:\Users\fabio\source\repos\PianificazioneScadenze\PianificazioneScadenze\bin\Debug\prova2.xlsx");
            ex.Close();
            */
            /*
            ex.CreateNewFile();
            ex.CreateNewSheet();
            ex.SaveAs(@"C:\Users\fabio\source\repos\PianificazioneScadenze\PianificazioneScadenze\bin\Debug\createnewfile.xlsx");
            ex.Close();*/
            Excel ex = new Excel(@"C:\Users\fabio\source\repos\ExcelAPI\ExcelAPI\bin\Debug\test", 1);
            MessageBox.Show(ex.ReadCell(1, 1));
            int colonne = ex.CountColumn();
            int righe = ex.CountRows();

            string[,] read = ex.ReadRange(1, 1, righe, colonne);
            ex.Save();
            ex.Close();

            Excel ex2 = new Excel(@"C:\Users\fabio\source\repos\ExcelAPI\ExcelAPI\bin\Debug\test1", 1);
            ex2.WriteRange(1, 1, righe, colonne, read);
            ex2.SaveAs(@"C:\Users\fabio\source\repos\ExcelAPI\ExcelAPI\bin\Debug\test2");
            ex2.Close();

            //Excel ex = new Excel(@"C:\Users\fabio\source\repos\ExcelAPI\ExcelAPI\bin\Debug\test", 1);
            //MessageBox.Show("colonne " + ex.CountColumn() + "righe" + ex.CountRows());
            //ex.Save();
            //ex.Close();

        }
        /* Importazione file excel e di conseguenza lettura del file */
        public void OpenFile()
        {
            Excel excel = new Excel(@"C:\Users\fabio\source\repos\ExcelAPI\ExcelAPI\bin\Debug\test", 1);
            MessageBox.Show(excel.ReadCell(0, 0));
        }

        //Provo a scrivere in una cella e a salvare il file excel che leggo e a salvarne un altro
        public void WriteData()
        {
            Excel excel = new Excel(@"C:\Users\fabio\source\repos\ExcelAPI\ExcelAPI\bin\Debug\test", 1);
            excel.WriteToCell(1, 1, "Provo il metodo write to cell");
            excel.Save();
            excel.SaveAs(@"C:\Users\fabio\source\repos\ExcelAPI\ExcelAPI\bin\Debug\test");
            excel.Close();
        }
    }
}
