using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using System.Runtime.InteropServices;
using System.Drawing;

namespace WindowsFormsApp2
{
    public partial class Main : Form
    {
        protected override CreateParams CreateParams
        {
            get
            {
                const int CS_DROPSHADOW = 0x20000;
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }
        }

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        string SourceFile = String.Empty;
        string DestinationFile = String.Empty;
        string MergedPDFfile = String.Empty;
        string OPD_IPD = String.Empty;
        string defaultPath = "";

        bool ODBCMode = false;

        public Main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                defaultPath = AppDomain.CurrentDomain.BaseDirectory;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot get defualt path! " + ex.Message.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (SourceFile == String.Empty)
            {
                MessageBox.Show("Please input RPT Source file!", "RPT Source file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (DestinationFile == String.Empty)
            {
                MessageBox.Show("Please input PDF Source file!", "PDF Source file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MergedPDFfile == String.Empty)
            {
                MessageBox.Show("Please select Destination output file!", "Destination Path", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ReportDocument report = new ReportDocument();

            report.Load(SourceFile);

            if (ODBCMode)
            {
                report.SetDatabaseLogon(txtUser.Text, txtPassword.Text);
                //report.SetDatabaseLogon("ADDONESIGNATURE", "Ac#08n7H8@y38347n");
            }
            else
            {
                report.SetDatabaseLogon(txtUser.Text, txtPassword.Text, txtDatabaseUrl.Text, txtDatabaseName.Text);
            }

            foreach (Control control in flwParameters.Controls)
            {

                if (control is UserControl userControl)
                {
                    Label prmName = userControl.Controls["tableLayoutPanel1"].Controls["lblParamsName"] as Label;
                    TextBox prmValue = userControl.Controls["tableLayoutPanel1"].Controls["txtValue"] as TextBox;

                    report.SetParameterValue(prmName.Text, prmValue.Text);
                }
            }

            ExportOptions exportOptions = new ExportOptions();
            PdfRtfWordFormatOptions pdfFormatOpts = ExportOptions.CreatePdfRtfWordFormatOptions();
            exportOptions.ExportFormatOptions = pdfFormatOpts;

            string outputFile = defaultPath + "Template.pdf";

            try
            {
                report.ExportToDisk(ExportFormatType.PortableDocFormat, outputFile);

                MergeBackgroundWithPdf(outputFile);

                lblMessage.Text = "Export Succesful!";
                lblMessage.ForeColor = Color.Green;


            }
            catch (LogOnException err)
            {
                lblMessage.Text = "Export error" + err.Message.ToString();
                lblMessage.ForeColor = Color.Red;
            }
            
            report.Close();
            report.Dispose();

            
        }

        public void MergeBackgroundWithPdf(string pdfFilePath)
        {
            string originalPdfPath = pdfFilePath;
            string letterHeadPath = DestinationFile;

            string mergedFileName = $"{OPD_IPD}_Merge_{DateTime.Now.ToString("ddMMyyyy_HHmmss")}";

            try
            {
                PdfReader contentReader = new PdfReader(originalPdfPath);

                if (contentReader != null)
                {
                    using (FileStream fs = new FileStream(MergedPDFfile + mergedFileName + ".pdf", FileMode.Create))
                    {
                        PdfStamper stamper = new PdfStamper(contentReader, fs);

                        PdfReader templateReader = new PdfReader(letterHeadPath);
                        PdfImportedPage templatePage = stamper.GetImportedPage(templateReader, 1);

                        for (int i = 1; i <= contentReader.NumberOfPages; i++)
                        {
                            PdfContentByte contentPage = stamper.GetUnderContent(i);
                            if (SourceFile.Contains("IPD"))
                            {
                                contentPage.AddTemplate(templatePage, -131, -161); // IPD Offset (-131, -161)
                            }
                            else
                            {
                                contentPage.AddTemplate(templatePage, -175, -107); // OPD Offset (-175, -107)
                            }
                            
                        }

                        stamper.Close();
                        templateReader.Close();
                        contentReader.Close();
                    }
                }
                else
                {
                    Console.WriteLine("contentReader is null. Please check the initialization.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        private void btnBrowseSource_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse PDF Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "rpt",
                Filter = "rpt files (*.rpt)|*.rpt",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                lblFile1.Text = openFileDialog1.FileName;
                SourceFile = openFileDialog1.FileName;

                if (SourceFile.Contains("IPD"))
                {
                    OPD_IPD = "IPD";
                }
                else
                {
                    OPD_IPD = "OPD";
                }

                flwParameters.Controls.Clear();

                ReportDocument report = new ReportDocument();
                
                try
                {
                    report.Load(SourceFile);

                    try
                    {
                        for (int i = 0; i < report.ParameterFields.Count; i++)
                        {
                            ParametersControl prm = new ParametersControl();
                            string paramsType = report.ParameterFields[i].ParameterValueType.ToString().Split('P')[0];
                            string name = report.ParameterFields[i].Name;
                            string value = String.Empty;

                            if (report.ParameterFields[i].CurrentValues.Count == 0)
                            {
                                value = "";
                            }
                            else
                            {
                                var paramField = report.ParameterFields[i];
                                ParameterDiscreteValue parameterValue = (ParameterDiscreteValue)paramField.CurrentValues[0];
                                value = Convert.ToString(parameterValue.Value);

                            }

                            prm.ParamsName.Text = name;
                            prm.ParamsType.Text = paramsType;
                            prm.ParamsValue.Text = value;

                            flwParameters.Controls.Add(prm);
                            //prm.Dock = DockStyle.Top;
                        }

                        flwParameters.AutoScroll = true;
                        flwParameters.WrapContents = false;
                    }
                    catch (Exception paramsError)
                    {
                        MessageBox.Show("Reading parameters error : " + paramsError.ToString(), "Throw Exception Parameters", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    try
                    {
                        var db = report.Database;

                        if (db.Tables.Count > 0)
                        {
                            dbGroup.Enabled = true;

                            Database database = report.Database;

                            foreach (Table table in database.Tables)
                            {
                                TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                                ConnectionInfo connectionInfo = tableLogOnInfo.ConnectionInfo;

                                string serverName = connectionInfo.ServerName;
                                string databaseName = connectionInfo.DatabaseName;
                                string userID = connectionInfo.UserID;
                                string password = connectionInfo.Password;
                                NameValuePair2 connectionStringType = (NameValuePair2)connectionInfo.LogonProperties[0];


                                // -------- read fields name inside the table ----------
                                //string tableName = table.Name;
                                //var fields = table.Fields;
                                //foreach (var field in fields)
                                //{
                                //    var fieldName = field;
                                //}
                                // -----------------------------------------------------

                                if (connectionStringType.Name.ToString() == "DSN")
                                {
                                    txtDatabaseUrl.Enabled = false;
                                    txtDatabaseUrl.BackColor = SystemColors.Control;
                                    txtDatabaseName.Enabled = false;
                                    txtDatabaseName.BackColor = SystemColors.Control;
                                    ODBCMode = true;
                                }
                                else
                                {
                                    txtDatabaseUrl.Enabled = true;
                                    txtDatabaseUrl.BackColor = Color.White;
                                    txtDatabaseName.Enabled = true;
                                    txtDatabaseName.BackColor = Color.White;
                                    ODBCMode = false;
                                }
                            }

                        }
                        else
                        {
                            dbGroup.Enabled = false;
                        }
                    }
                    catch (Exception DatabaseError)
                    {
                        MessageBox.Show("Reading database error : " + DatabaseError.ToString(), "Throw Exception Datasorce", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                   

                }
                catch (Exception err)
                {
                    MessageBox.Show("Connot load parameters" + err.ToString());
                }
                finally 
                {
                    report.Close();
                }
            }
        }

        private void btnBrowseSource2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse PDF Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "pdf",
                Filter = "pdf files (*.pdf)|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                lblFile2.Text = openFileDialog1.FileName;
                DestinationFile = openFileDialog1.FileName;
            }
        }

        private void btnDestinationPath_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);

                    lblDestPath.Text = fbd.SelectedPath;
                    MergedPDFfile = fbd.SelectedPath + "\\";
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
    }
}
