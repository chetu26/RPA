using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using DataTable = System.Data.DataTable;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Column = DocumentFormat.OpenXml.Wordprocessing.Column;
using DocumentFormat.OpenXml;

namespace EditWordFiles
{
    public partial class WordFileModificationForm : Form
    {
        public WordFileModificationForm()
        {
            InitializeComponent();
        }
        public string pdfFilePath = null;
        //public string companyName = null;
        public int totalAmmountIndex = 0;
        //double totalAmount = 0;
        Dictionary<string, double> companyNameAndTotalAmountDict = new Dictionary<string, double>();
        public string docFilePath = null;
        public object txtFileContent { get; private set; }

        public void BrowsePdfBtn_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog openFileDialog = new FolderBrowserDialog();

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    pdfTextBox.Text = openFileDialog.SelectedPath;
                    pdfFilePath = pdfTextBox.Text;
                    GetTextFromPDF();
                }

                //GetTextFromWord();
            }
            catch (Exception ex)
            {
                LoggingLibrary.LoggingManager.LogException(this.GetType(), ex);
                MessageBox.Show("Failure! : " + ex.Message);
            }
        }

        private void GetTextFromPDF()
        {
            try
            {
                string[] files = Directory.GetFiles(pdfFilePath, "*.pdf");
                if (files.Length < 1)
                {
                    throw new Exception("There is no PDF file in this folder");
                }
                foreach (string file in files)
                {
                    string companyName = null;
                    double totalAmount = -1;
                    PdfReader reader = new PdfReader(file);
                    int PageNum = reader.NumberOfPages;
                    string[] words;
                    string text;
                    //double value = 0;

                    for (int i = 1; i <= PageNum; i++)
                    {
                        text = PdfTextExtractor.GetTextFromPage(reader, i, new LocationTextExtractionStrategy()).Trim();
                        text = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(text)));
                        words = text.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                        if (i == 1)
                        {
                            companyName = words[0].Trim();
                        }
                        else if (i == 2)
                        {
                            for (int w = 0; w <= words.Length; w++)
                            {
                                if (words[w].Trim() == "TOTAL AMOUNT")
                                {
                                    totalAmmountIndex = w;
                                    break;
                                }
                            }
                            for (int t = totalAmmountIndex; t >= 0; t--)
                            {
                                if (double.TryParse(words[t].Trim(), out totalAmount))
                                {
                                    if (totalAmount > 0)
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if ((!string.IsNullOrEmpty(companyName)) && totalAmount >= 0)
                    {
                        companyNameAndTotalAmountDict[companyName] = totalAmount;
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingLibrary.LoggingManager.LogException(this.GetType(), ex);
                MessageBox.Show("Failure! : " + ex.Message);
            }


        }
        private void BrowseDocBtn_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Word document (*.doc*)|*.doc*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    docTextBox.Text = openFileDialog.FileName;
                }
                docFilePath = docTextBox.Text;
            }
            catch (Exception ex)
            {
                LoggingLibrary.LoggingManager.LogException(this.GetType(), ex);
                MessageBox.Show("Failure! : " + ex.Message);
            }
        }

        private void AddColumnAndData(string companyName,double totalAmount)
        {
            try
            {
                DataSet dataSet = new DataSet();
                using (var doc = WordprocessingDocument.Open(docTextBox.Text.Trim(), true))
                {
                    List<Table> tables = doc.MainDocumentPart.Document.Body.Elements<Table>().Where(tbl => tbl.InnerText.Contains("Joint Evaluation") || tbl.InnerText.Contains("Commercial Evaluation")).ToList();
                    // To get all rows from table  
                    for (int i = 0; i < tables.Count; i++)
                    {
                        List<TableRow> tableRows = tables[i].Elements<TableRow>().ToList();

                        // Create a cell.
                        for (int rowId = 0; rowId < tableRows.Count; rowId++)
                        {
                            //TableCell tc1 = new TableCell();
                            TableCell tc1 = (TableCell)tableRows[rowId].Elements<TableCell>().LastOrDefault().CloneNode(true);

                            if (!string.IsNullOrEmpty(tc1.InnerText))
                                tc1.InnerXml = tc1.InnerXml.Replace(tc1.InnerXml, tc1.InnerXml.Replace(tc1.InnerText, ""));

                            List<Paragraph> p = tc1.Elements<Paragraph>().ToList();
                            bool isAmountWritten = false;
                            foreach (var para in p)
                            {
                                List<Run> r = para.Elements<Run>().ToList();

                                int runcount = -1;
                                foreach (var run in r)
                                {
                                    runcount++;
                                    // Set the text for the run.
                                    Text t = run?.Elements<Text>().FirstOrDefault();
                                    if (t != null)
                                    {
                                        t.Text = "";
                                    }

                                    //// Specify the table cell content.
                                    if (rowId == 0)
                                    {
                                        TableCellProperties tcp = new TableCellProperties(
                                        new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50%", }
                                       );

                                        if (t != null && runcount == 0)
                                            t.Text = companyName;

                                        else if (t == null)
                                            run.Append(new Text(companyName));                                           

                                        tc1.Append(tcp);
                                    }
                                    else if (rowId == 1)
                                    {
                                        if (!isAmountWritten)
                                        {
                                            if (t != null && runcount == 0)
                                                t.Text = totalAmount.ToString("#,##0.00");
                                            else if (t == null || string.IsNullOrEmpty(t.Text))
                                            {
                                                run.Append(new Text(totalAmount.ToString("#,##0.00")));                                               
                                            }
                                            isAmountWritten = true;

                                        }
                                    }
                                    else
                                    {
                                        if (t != null && runcount == 0)
                                            t.Text = "";
                                        else if (t == null)
                                            run.Append(new Text(""));                                      
                                    }

                                   
                                }

                                if (runcount == -1)
                                {
                                    //// Specify the table cell content.
                                    if (rowId == 0)
                                    {
                                        para.Append(new Run(new Text(companyName)));                                       
                                    }
                                    else if (rowId == 1)
                                    {
                                        if (!isAmountWritten)
                                        {
                                            para.Append(new Run(new Text(totalAmount.ToString("#,##0.00"))));
                                            isAmountWritten = true;
                                        }

                                    }
                                    else
                                    {
                                        para.Append(new Run(new Text("")));                                       
                                    }                                   
                                }
                            }

                            // Append the table cell to the table row.
                            tableRows[rowId].Append(tc1);
                        }

                    }

                    doc.Save();
                }
            }
            catch (Exception ex)
            {
                LoggingLibrary.LoggingManager.LogException(this.GetType(), ex);
                MessageBox.Show("Failure! : " + ex.Message);
            }
        }


        private void ModifyBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if ((pdfTextBox.Text == "" || pdfTextBox.Text == null) || (docTextBox.Text == "" || docTextBox.Text == null))
                {
                    throw new Exception("Please select PDF folder And Template file");
                }
                else if (pdfTextBox.Text == "" || pdfTextBox.Text == null)
                {
                    throw new Exception("Please select PDF folder");
                }
                else if (docTextBox.Text == "" || docTextBox.Text == null)
                {
                    throw new Exception("Please select Template file");
                }
                else
                {
                    if (FileIsOpen(docFilePath))
                    {
                        throw new Exception("Template file is already in use.");
                    }
                    foreach (string key in companyNameAndTotalAmountDict.Keys)
                    {
                        AddColumnAndData(key, companyNameAndTotalAmountDict[key]);
                    }
                    MessageBox.Show("Transformation Successful!");
                }
            }
            catch (Exception ex)
            {
                LoggingLibrary.LoggingManager.LogException(this.GetType(), ex);
                MessageBox.Show("Failure! : " + ex.Message);
            }
        }


        public bool FileIsOpen(string path)
        {
            System.IO.FileStream a = null;

            try
            {
                a = System.IO.File.Open(path,
                System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None);
                return false;
            }
            catch (System.IO.IOException ex)
            {
                LoggingLibrary.LoggingManager.LogException(this.GetType(), ex);
                return true;
            }

            finally
            {
                if (a != null)
                {
                    a.Close();
                    a.Dispose();
                }
            }
        }
    }
}
