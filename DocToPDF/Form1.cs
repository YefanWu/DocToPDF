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
using Word = NetOffice.WordApi;
using NetOffice;

namespace DocToPDF
{
    public partial class Form1 : Form
    {
        string[] docName; //Class level array for file name & extension.
        string[] docNameNoExt; //File name without path.
        string[] docPath; //File path with name and ext.
        string[] pdfPath; //Replace file ext with .pdf

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            //Clear existing file list.
            listFiles.Items.Clear();
            try
            {
                Array.Clear(docName, 0, docName.Length);
                Array.Clear(docPath, 0, docPath.Length);
            }
            catch (Exception)
            {

            }


            Stream wordDocs = null; //stream for Doc files.
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Multiselect = true; //Allow select mutiple files.
            openfile.Title = "Pick up your Word document!";
            openfile.Filter = "Word Document (*.doc, *.docx) | *.doc; *.docx";
            openfile.RestoreDirectory = true; //use the same path with last time.

            if (openfile.ShowDialog() == DialogResult.OK) //select file(s) and click on OK
            {
                try
                {
                    if ((wordDocs = openfile.OpenFile()) != null)
                    {
                        docName = openfile.SafeFileNames;
                        docPath = openfile.FileNames;
                        docNameNoExt = new string[docName.Length];
                        pdfPath = new string[docName.Length];
                        UpdateNameList(docName);
                        wordDocs.Close();
                        for (int i = 0; i < docPath.Length; i++)
                        {
                            docNameNoExt[i] = Path.GetFileNameWithoutExtension(docPath[i]);
                            pdfPath[i] = Path.GetDirectoryName(docPath[i]) + "\\" + docNameNoExt[i] + ".pdf";
                        }
                        
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }

            }

        }

        private void UpdateNameList(string[] names)
        {
            listFiles.BeginUpdate();
            listFiles.Columns.Add("File Path", 180, HorizontalAlignment.Left); //Create a default column, otherwise no data can show in the list.
            listFiles.Columns[0].Width = listFiles.ClientSize.Width; //Make the first column same size as form. So the additional blank column is gone. Also the size can adjust automatically. 
            for (int i = 0; i < names.Length; i++) //go through all the file names in array and display in the list
            {
                ListViewItem viewItem = new ListViewItem();
                viewItem.Text = names[i];
                viewItem.Checked = true;
                listFiles.Items.Add(viewItem);
            }
            listFiles.EndUpdate();
            StatusLabel1.Text = "Files loaded.";
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            listFiles.Items.Clear();
            Array.Clear(docName,0,docName.Length);
            ProgressBar.Value = 0;
            StatusLabel1.Text = "Cleared";
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            StatusLabel1.Text = "Converting...";
            WordToPDF(docPath, pdfPath);
            StatusLabel1.Text = "Completed";
        }

        private void WordToPDF(string[] docNames, string[] pdfNames)
        {
            //Initialize progress bar
            ProgressBar.Maximum = docName.Length;
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = Word.Enums.WdAlertLevel.wdAlertsNone; //Do not display alerts from Word app. 
            for (int i = 0; i < docNames.Length; i++)
            {
                Word.Document doc;
                doc = wordApplication.Documents.Open(docNames[i], false);
                //Inheret convert tool see https://msdn.microsoft.com/en-us/VBA/Word-VBA/articles/document-exportasfixedformat-method-word
                doc.ExportAsFixedFormat(pdfNames[i], Word.Enums.WdExportFormat.wdExportFormatPDF, false, 1);
                doc.Close();
                ProgressBar.Value++;
            }
            
        }
    }
}
