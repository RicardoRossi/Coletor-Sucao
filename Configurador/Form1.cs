﻿using SldWorks;
using SwConst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using LerExcel;
using System.Globalization;

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
        ModelView mView;

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

            var excel = new Read_From_Excel();

            // O metodo retorna uma lista de coletores
            List<Coletor> coletores = excel.getColetores();

            // Converto a Lista retornada para array para acessar pelo indice.
            Array c = (coletores.ToArray());

            //swApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocNONE);

            for (int i = 0; i < c.Length; i++)
            {
                Coletor coletor = (Coletor)c.GetValue(i);
                string qtCP = coletor.QuantidadeCompressor;

                OpenColetorTemplate(coletor);

                SaveAs2d(coletor);

                // Salva o 3D e troca referencia no novo 2d
                SaveAs3d(coletor);

                InserirPropriedade(coletor);

                // Replace ramal do rack
                ReplaceBolsaRack(coletor);

                // Replace ramal compressor.
                ReplaceBolsaCP(coletor);

                SaveAsTubo(coletor);

                // Salva o 2D final
                swApp.ActivateDoc(coletor.CodigoColetor + ".SLDDRW");
                swModel = swApp.ActiveDoc;

                int error = 0;
                int warnings = 0;
                swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced + (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    ref error, ref warnings);

                swModel.SaveAs(Path.ChangeExtension(swModel.GetPathName(),".PDF"));
                
                swApp.CloseAllDocuments(true);
            }

            //Replace(swModel);
            //swModel.EditRebuild3();

            //mView.EnableGraphicsUpdate = true;
        }

        private void SaveAs2d(Coletor coletor)
        {
            string codigo = coletor.CodigoColetor;
            string caminhoSalvar = @"C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\COLETOR SUCCAO\";
            string nomeCompletoArquivo2d = caminhoSalvar + codigo + ".SLDDRW";
            int error = 0;
            int warning = 0;

            swModel = swApp.ActiveDoc;
            swExt = swModel.Extension;

            //mView = swModel.ActiveView;
            //mView.EnableGraphicsUpdate = false;

            // Salva e abre o arquivo e ativa para trocar a referencia do coletor.
            int retVal = swModel.SaveAs3(nomeCompletoArquivo2d, 0, 0);
            swApp.OpenDoc6(nomeCompletoArquivo2d, (int)swDocumentTypes_e.swDocDRAWING,
              (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", error, warning);
            //swModel = swApp.ActiveDoc;
        }

        private void SaveAs3d(Coletor coletor)
        {
            string codigo = coletor.CodigoColetor;
            string caminhoSalvar = @"C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\COLETOR SUCCAO\";
            string nomeCompletoArquivo3d = caminhoSalvar + codigo + ".SLDASM";

            // Mostra o 3D
            swApp.ActivateDoc(Path.GetFileName(coletor.ArquivoTemplateDoColetor));

            // Ativa o 3D
            swModel = swApp.ActiveDoc;
            swExt = swModel.Extension;
            int retVal = swModel.SaveAs3(nomeCompletoArquivo3d, 0, 0);
        }

        private void SaveAsTubo(Coletor coletor)
        {
            string codigo = coletor.CodigoTuboAcoColetor;
            string caminhoSalvar = @"C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\COLETOR SUCCAO\";
            string nomeCompletoArquivo3d = caminhoSalvar + codigo + ".SLDPRT";

            // Mostra o 3D
            swApp.ActivateDoc(coletor.ArquivoTemplateTuboDoColetor);

            // Ativa o 3D
            swModel = swApp.ActiveDoc;
            swExt = swModel.Extension;
            int retVal = swModel.SaveAs3(nomeCompletoArquivo3d, 0, 0);

            AlterarDimensao(coletor);
        }

        private void ReplaceBolsaRack(Coletor coletor)
        {
            swApp.ActivateDoc(coletor.CodigoColetor + ".SLDASM");

            swModel = swApp.ActiveDoc;
            AssemblyDoc swAssembly = (AssemblyDoc)swModel;
            Object[] components = swAssembly.GetComponents(true);

            foreach (var componente in components)
            {
                Component2 component = (Component2)componente;

                swModel = component.GetModelDoc2();

                string nomeCompletoDoComponente = swModel.GetPathName();
                string nomeComExtensao = Path.GetFileName(nomeCompletoDoComponente);

                if (String.Equals(nomeComExtensao, "BOLSA SOLDA SUCCAO RACK TEMPLATE.SLDPRT"))
                {
                    component.Select(true);
                    break;
                }
            }

            swAssembly.ReplaceComponents($@"C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\COLETOR SUCCAO\{coletor.CodigoBolsaSoldaSuccaoRack}.SLDPRT",
                "", true, true);

            swModel = swApp.ActiveDoc;
            swModel.Save();
        }

        private void ReplaceBolsaCP(Coletor coletor)
        {
            swApp.ActivateDoc(coletor.CodigoColetor + ".SLDASM");
            swModel = swApp.ActiveDoc;

            AssemblyDoc swAssembly = (AssemblyDoc)swModel;
            Object[] components = swAssembly.GetComponents(true);

            foreach (var componente in components)
            {
                Component2 component = (Component2)componente;

                swModel = component.GetModelDoc2();

                string nomeCompletoDoComponente = swModel.GetPathName();
                string nomeComExtensao = Path.GetFileName(nomeCompletoDoComponente);

                if (String.Equals(nomeComExtensao, "BOLSA SOLDA SUCCAO CP TEMPLATE.SLDPRT"))
                {
                    component.Select(true);
                    break;
                }
            }

            swAssembly.ReplaceComponents($@"C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\COLETOR SUCCAO\{coletor.CodigoBolsaSoldaSuccaoCompressor}.SLDPRT",
                "", true, true);

            swModel = swApp.ActiveDoc;
            swModel.Save();
        }

        private void OpenColetorTemplate(Coletor c)
        {
            int errors = 0;
            int warnings = 0;

            // Abre o assembly do coletor
            swApp.OpenDoc6(c.ArquivoTemplateDoColetor, (int)swDocumentTypes_e.swDocASSEMBLY,
                (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", (int)errors, (int)warnings);

            // Converte o path de sldasm para slddrw
            string path2d = c.ArquivoTemplateDoColetor.Replace("SLDASM", "SLDDRW");

            // Abre o 2d do coletor
            swApp.OpenDoc6(path2d, (int)swDocumentTypes_e.swDocDRAWING,
               (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", (int)errors, (int)warnings);
        }

        private void Replace(ModelDoc2 swModel)
        {
            swAssembly = (AssemblyDoc)swModel;

            swAssembly.ReplaceComponents(@"C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\COLETOR SUCCAO\2047907.SLDPRT",
            "", true, true);
        }

        private void SalvarPDF()
        {
            string nome = Path.GetFileNameWithoutExtension(swModel.GetPathName()); // Pega o nome sem extensão do full path do nome original com extensão.
            int Error = 0;
            int Warnings = 0;
            bool bRet;
            bRet = swExt.SaveAs($@"C:\ELETROFRIO\ENGENHARIA SMR\PRODUTOS FINAIS ELETROFRIO\MECÂNICA\RACK PADRAO\COLETOR SUCCAO\{nome}.PDF", (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref Error, ref Warnings);

            // Converte um enum do tipo int para a string do enum deixando mais claro o erro
            // pois será retornada um msg e não um int.
            swFileSaveError_e e = (swFileSaveError_e)Error;
            MessageBox.Show(e.ToString());
        }

        private void AlterarDimensao(Coletor c)
        {
            myDimension = swModel.Parameter("D1_RAMAL_RACK@Sketch3");
            string s1 = c.DiametroEncaixeSuccaoRack.Replace(",", ".");
            string s2 = c.DiametroEncaixeSuccaoCompressor.Replace(",", ".");

            double d1 = Double.Parse(s1, CultureInfo.InvariantCulture);
            double d2 = Double.Parse(s2, CultureInfo.InvariantCulture);

            myDimension.SystemValue = d1 / 1000; // Converte pra metros.
            myDimension = swModel.Parameter("D1_RAMAL_CP@Sketch4");
            myDimension.SystemValue = d2 / 1000; // Converte pra metros.
            swModel.EditRebuild3();
        }

        private void InserirPropriedade(Coletor c)
        {
            string descricao = "";

            descricao += "COL S ";
            descricao += c.DiametroTuboAcoColetor + "\" ";
            descricao += c.QuantidadeCompressor + "CP ";
            descricao += c.DiametroSuccaoCompressor + "\" X ";
            descricao += c.DiametroSuccaoRack + "\" ";

            swModel = swApp.ActiveDoc;
            ConfigurationManager configMgr;
            configMgr = swModel.ConfigurationManager;
            Configuration config = configMgr.ActiveConfiguration;
            string nomeConfig = config.Name;
            swExt = swModel.Extension;

            // Deleta prop da custom
            swCustomMgr = swExt.CustomPropertyManager[""];
            string[] nomesProp = null;
            nomesProp = (string[])swCustomMgr.GetNames();
            
            foreach (var nome in nomesProp)
            {
                swCustomMgr.Delete(nome);
            }

            // Deleta prop da personalizada
            swCustomMgr = swExt.CustomPropertyManager[nomeConfig];
            nomesProp = null;
            nomesProp = (string[])swCustomMgr.GetNames();

            foreach (var nome in nomesProp)
            {
                swCustomMgr.Delete(nome);
            }

            swCustomMgr.Add3("DESCRIÇÃO", (int)swCustomInfoType_e.swCustomInfoText, descricao,
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            swCustomMgr.Add3("PROJETISTA", (int)swCustomInfoType_e.swCustomInfoText, "RICARDO R.",
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            swCustomMgr.Add3("PROJETISTA2D", (int)swCustomInfoType_e.swCustomInfoText, "RICARDO R.",
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            swCustomMgr.Add3("GRUPO ITEM", (int)swCustomInfoType_e.swCustomInfoText, "494",
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            swCustomMgr.Add3("REVISÃO", (int)swCustomInfoType_e.swCustomInfoText, "01",
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
