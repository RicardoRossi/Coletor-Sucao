using SldWorks;
using SwConst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using LerExcel;

namespace Configurador
{
    public partial class Form1 : Form
    {
        SldWorks.SldWorks swApp;
        ModelDoc2 swModel;
        PartDoc swPart;
        AssemblyDoc swAssembly;
        ModelDocExtension swExt;
        CustomPropertyManager swCustomMgr;
        Dimension myDimension;
        SheetMetalFeatureData sMetal;
        Feature feature;
        Feature swSubFeature;


        // Construtor da classe
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                swApp = (SldWorks.SldWorks)Marshal.GetActiveObject("SldWorks.application");
            }
            catch
            {
                MessageBox.Show("Erro ao conectar no Solidworks");
                return;
            }

            swModel = (ModelDoc2)swApp.ActiveDoc;
            if (swModel == null)
            {
                MessageBox.Show("Não há documento aberto");
            }

            swExt = swModel.Extension;

            var excel = new Read_From_Excel();

            // O metodo retorna uma lista de coletores
            List<Coletor> coletores = excel.getColetores();

            // Converto a Lista retornada para array para acessar pelo indice.
            Array c = (coletores.ToArray());

            //Coletor coletorZero = (Coletor)c.GetValue(0);

            //string msg = "";
            //foreach (Coletor coletor in coletores)
            //{                
            //    msg += coletor.codigoColetor + "\n";
            //}
            //MessageBox.Show(msg);

            for (int i = 0; i < c.Length; i++)
            {
                Coletor coletor = (Coletor)c.GetValue(i);
                string qtCP = coletor.QuantidadeCompressor;

                OpenColetorTemplate(coletor);
            }

            //Replace(swModel);
            swModel.EditRebuild3();

        }

        private void OpenColetorTemplate(Coletor c)
        {
            int errors = 0;
            int warnings = 0;

            swApp.OpenDoc6(c.ArquivoColetorTemplate, (int)swDocumentTypes_e.swDocASSEMBLY,
                (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", (int)errors, (int)warnings);
        }

        private void Replace(ModelDoc2 swModel)
        {
            swAssembly = (AssemblyDoc)swModel;

            swAssembly.ReplaceComponents(@"C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\CONEXOES\2047907.SLDPRT",
            "", true, true);
        }

        private void SalvarPDF()
        {
            string nome = Path.GetFileNameWithoutExtension(swModel.GetPathName()); // Pega o nome sem extensão do full path do nome original com extensão.
            int Error = 0;
            int Warnings = 0;
            bool bRet;
            bRet = swExt.SaveAs($@"C:\Users\54808\Documents\{nome}.PDF", (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref Error, ref Warnings);

            // Converte um enum do tipo int para a string do enum deixando mais claro o erro
            // pois será retornada um msg e não um int.
            swFileSaveError_e e = (swFileSaveError_e)Error;
            MessageBox.Show(e.ToString());
        }

        private void AlterarDimensao(double dimensao)
        {
            myDimension = swModel.Parameter("comprimento@comprimento@Part1.Part");
            myDimension.SystemValue = dimensao / 1000; // Converte pra metros.
            swModel.EditRebuild3();
        }

        private void InserirPropriedade()
        {
            swExt = swModel.Extension;
            swCustomMgr = swExt.CustomPropertyManager[""];
            swCustomMgr.Add3("Descrição", (int)swCustomInfoType_e.swCustomInfoText, "Grade ventilador ZA",
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
        }

        private void LerArquivo()
        {
            try
            {
                // O using é equivalente a um try finally que chama um Dispose
                // para liberar recursos.
                using (StreamWriter sw = new StreamWriter("properties.txt"))
                {
                    sw.WriteLine("Peso");
                    sw.WriteLine("Descrição");
                    sw.WriteLine("Material");
                    sw.WriteLine("Dimensões");
                    // sr.Close(); // Não é necessário o Closse(), pois o using
                    // faz o dispose
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            // Cria a list para armazenar as properties
            List<string> listaNomePropriedades = new List<string>();

            // Cria o stream para ler arquivo
            using (StreamReader sr = new StreamReader("properties.txt"))
            {
                // Lê o arquivo até o final da stream
                while (!sr.EndOfStream)
                {
                    listaNomePropriedades.Add(sr.ReadLine());
                }
            }

            // Lê a lista na memória até o final e constroi
            // uma mensagem para MessageBox
            string props = "";
            foreach (var item in listaNomePropriedades)
            {
                props += item + "\n";
            }
            MessageBox.Show(props);
        }
    }
}
