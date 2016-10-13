using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace ConverteDocs
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private List<string> ListaArqs;

        private void PesquisaArquivos(string NomeDir)
        {
            // cria instância de DirectoryInfo para o diretório selecionado
            DirectoryInfo DirInfo = new DirectoryInfo(NomeDir);
            try
            {
                // obtém arquivos do diretório
                FileInfo[] AFileInfo = DirInfo.GetFiles("*.doc");
                // processa arquivos, adicionando-os na ListView
                foreach (FileInfo FilInfo in AFileInfo)
                    ListaArqs.Add(FilInfo.FullName);
                // procura subdiretórios
                DirectoryInfo[] ADirInfo = DirInfo.GetDirectories();
                // chama função recursivamente
                foreach (DirectoryInfo DirecInfo in ADirInfo)
                    PesquisaArquivos(DirecInfo.FullName);
            }
            catch
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = textBox1.Text;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                textBox1.Text = folderBrowserDialog1.SelectedPath; 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                ListaArqs = new List<string>();
                PesquisaArquivos(textBox1.Text);
                checkedListBox1.Items.Clear();
                foreach(string Arq in ListaArqs)
                    checkedListBox1.Items.Add(Arq);
            }
        }

        private ApplicationClass wordApp;
        
        private void ConverteArquivo(string NomeArq)
        {
            object read_only = false;
            object visible = false;
            object confirm = true;
            object missing = System.Reflection.Missing.Value;
            object Formato = WdSaveFormat.wdFormatXMLDocument;
            object Arquivo = NomeArq;
  

            Document doc = wordApp.Documents.Open(ref Arquivo, ref missing, ref read_only, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref visible, 
                ref missing, ref missing, ref missing, ref missing);

            Arquivo = Arquivo + "x";
            doc.SaveAs(ref Arquivo, ref Formato, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            object missing = System.Reflection.Missing.Value;

            wordApp = new ApplicationClass();

            Cursor = Cursors.WaitCursor;
            try
            {
                foreach (string Nome in checkedListBox1.CheckedItems)
                    ConverteArquivo(Nome);
            }
            finally
            {
                Cursor = Cursors.Default;
                wordApp.Quit(ref missing, ref missing, ref missing);
            }
        }

    }
}