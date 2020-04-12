using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

/* Per poter utilizzare Microsoft.Office.Interop aggiungere riferimento -> COM -> microsoft.office.interop.excel 
 * Se ci fossero problemi con la le librerie disinstallare in maniera completa Office,
 * disinstallare tramite il tool disponibile sul sito Microsoft e reinstallare da zero. */
namespace ExcelAPI
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel()
        {

        }
        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
            MessageBox.Show("prova");
        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        }

        // Creare un nuovo foglio/sheet
        public void CreateNewSheet()
        {
            Worksheet temptsheet = wb.Worksheets.Add(After: ws);
        }

        // Legge una singola cella (riga, colonna)
        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2;
            else
                return "";
        }

        // Legge tutto il foglio di excel ritorna un array di stringhe (prima riga da dove iniziare a leggere, prima colonna, ultima riga , ultima colonna
        public string[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnstring = new string[endi - starti, endy - starty];
            for (int p = 1; p <= endi - starti; p++)
            {
                for (int q = 1; q <= endy - starty; q++)
                {
                    returnstring[p - 1, q - 1] = holder[p, q].ToString();
                }
            }
            return returnstring;
        }

        // conta il numero di righe
        public int CountRows()
        {
            return ws.UsedRange.Rows.Count;
        }

        public int CountColumn()
        {
            // conta il numero di colonne
            return ws.UsedRange.Columns.Count;
        }

        // NON FUNZIONA RICONTROLLARE Scrive un array di stringhe all'interno di un foglio excel
        public void WriteRange(int starti, int starty, int endi, int endy, string[,] writestring)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = writestring;
        }

        // Scrive in una determinata cella (numero riga, numero cella, stringa)
        public void WriteToCell(int i, int j, string s)
        {
            ws.Cells[i, j].Value2 = s;
        }

        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        // Seleziona il foglio di lavoro in cui scrivere
        public void SelectWorksheet(int SheetNumber)
        {
            this.ws = wb.Worksheets[SheetNumber];
        }

        // Elimina un foglio di lavoro
        public void DeleteWorksheet(int SheetNumber)
        {
            wb.Worksheets[SheetNumber].Delete();
        }
        // Mette la protezione al file excel
        public void ProtectSheet()
        {
            ws.Protect();
        }

        public void ProtectSheet(string Password)
        {
            ws.Protect(Password);
        }
        // Toglie la protezione al file excel
        public void UnprotectSheet()
        {
            ws.Unprotect();
        }

        public void UnprotectSheet(string Password)
        {
            ws.Unprotect(Password);
        }

        public void Close()
        {
            wb.Close();
        }
    }
}
