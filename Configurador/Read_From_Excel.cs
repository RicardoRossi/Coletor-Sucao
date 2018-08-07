using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab
using Configurador;
using System.Collections.Generic;

namespace LerExcel
{
    public class Read_From_Excel
    {
        private List<Coletor> coletores;

        //Construtor
        public Read_From_Excel()
        {
            coletores = new List<Coletor>();
        }

        public List<Coletor> getColetores()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\col_sucao.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            int i = 0;
            int j = 0;
            for (i = 2; i <= rowCount; i++)
            {
                var c = new Coletor();
                c.CodigoColetor = xlRange.Cells[i, 1].Value2.ToString();
                c.DescricaoColetor = xlRange.Cells[i, 2].Value2.ToString();
                c.QuantidadeCompressor = xlRange.Cells[i, 3].Value2.ToString();
                c.DiametroTuboAcoColetor = xlRange.Cells[i, 4].Value2.ToString();
                c.CodigoTuboAcoColetor = xlRange.Cells[i, 5].Value2.ToString();
                c.QuantidadeRamalRack = xlRange.Cells[i, 6].Value2.ToString();
                c.DiametroSuccaoRack = xlRange.Cells[i, 7].Value2.ToString();
                c.CodigoBolsaSoldaSuccaoRack = xlRange.Cells[i, 8].Value2.ToString();
                c.DiametroEncaixeSuccaoRack = xlRange.Cells[i, 9].Value2.ToString();
                c.DiametroSuccaoCompressor = xlRange.Cells[i, 10].Value2.ToString();
                c.CodigoBolsaSoldaSuccaoCompressor = xlRange.Cells[i, 11].Value2.ToString();
                c.DiametroEncaixeSuccaoCompressor = xlRange.Cells[i, 12].Value2.ToString();
                c.ArquivoColetorTemplate= xlRange.Cells[i, 13].Value2.ToString();

                AddColetor(c);
            }

            void AddColetor(Coletor c)
            {
                coletores.Add(c);
            }

            //string celulaAtual = xlRange.Cells[2, 3].Value2.ToString();

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return coletores;
        }


    }
}

