using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace AbreWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object readOnly = false;
            object visible = true;
            object missing = System.Reflection.Missing.Value;
            object NomeArq;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                NomeArq = openFileDialog1.FileName;
                // Cria uma instância do Word
                ApplicationClass wordApp = new ApplicationClass();
                // Mostra o Word
                wordApp.Visible = true;
                // Abre o documento
                Document doc = wordApp.Documents.Open(ref NomeArq, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref visible, ref missing,
                    ref missing, ref missing, ref missing);

                NomeArq = NomeArq + "x";
                object Formato = WdSaveFormat.wdFormatXMLDocument;
                // Apenas para Word anterior ao 2007
                //object Formato = -1;
                //foreach (FileConverter conv in wordApp.FileConverters)
                //    if (conv.ClassName == "MSWord12")
                //    {
                //        Formato = conv.SaveFormat;
                //        break;
                //    }
                doc.Convert();
                doc.SaveAs(ref NomeArq, ref Formato, ref missing, ref missing, ref missing, 
                    ref missing, ref missing, ref missing, ref missing, ref missing, 
                    ref missing, ref missing, ref missing, ref missing, ref missing, 
                    ref missing);
                wordApp.Quit(ref missing, ref missing, ref missing);
            }
        }
    }
}