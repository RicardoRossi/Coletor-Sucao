using SldWorks;
using SwConst;
using SwCommands;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;

namespace Configurador
{
    public partial class Form1 : Form
    {
        SldWorks.SldWorks swApp;
        ModelDoc2 swModel;
        ModelDocExtension swExt;
        CustomPropertyManager swCustomMgr;

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
                //return;
            }

            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (swModel == null)
            {
                MessageBox.Show("Não há documento aberto");
            }

            //swApp.SendMsgToUser("Conectado");

            swExt = swModel.Extension;
            swCustomMgr = swExt.CustomPropertyManager[""];
            swCustomMgr.Add3("Descrição", (int)swCustomInfoType_e.swCustomInfoText, "Grade ventilador ZA",
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);


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
