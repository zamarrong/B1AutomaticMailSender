using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace B1AutomaticMailSender
{
    public partial class frmMain : Form
    {
        SAP sap;
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            sap = new SAP();
            MoveFiles();
            StartWork();
            Application.Exit();
        }

        private void StartWork()
        {
            ProcessDocuments("13");
            ProcessDocuments("14");
            ProcessDocuments("24");
            ProcessDocuments("203");
        }

        private void ProcessDocuments(string ObjType)
        {
            try
            {
                List<SAP.Document> documents = sap.GetSendPendingDocuments(ObjType, Properties.Settings.Default.Top);
                pb.Maximum = documents.Count;
                foreach (SAP.Document document in documents)
                {
                    Application.DoEvents();

                    lblStatus.Text = string.Format("Proccessing documents ({0}) {1} of {2}", SAP.SAPTableByObjType(ObjType), pb.Maximum, pb.Value);

                    Application.DoEvents();

                    MakePDF(document);

                    Application.DoEvents();

                    if (SendMail(MakeMail(document)))
                    {
                        sap.UpdateEDocSentStatus(document, "Y");
                    }
                    else
                    {
                        sap.UpdateEDocSentStatus(document, "E");
                        Program.WriteLog("Error sending email " + document.UUID);
                    }

                    Application.DoEvents();
                    pb.Value += 1;
                }
            }
            catch (Exception ex)
            {
                Program.WriteLog(ex.Message);
            }
            finally
            {
                pb.Value = pb.Maximum = 0;
            }
        }

        private bool SendMail(MailMessage mail)
        {
            try
            {
                lblStatus.Text = string.Format("Sending Mail to: {0}", mail.To);

                SmtpClient client = new SmtpClient();

                client.Port = Properties.Settings.Default.Port;
                client.Host = Properties.Settings.Default.Host;
                client.EnableSsl = Properties.Settings.Default.EnableSSL;
                client.Timeout = 10000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = Properties.Settings.Default.UseDefaultCredentials;
                client.Credentials = new System.Net.NetworkCredential(Properties.Settings.Default.UserName, Properties.Settings.Default.Password);

                if (mail != null)
                {
                    client.Send(mail);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Program.WriteLog(ex.Message);
                return false;
            }
        }

        private MailMessage MakeMail(SAP.Document document)
        {
            try
            {
                lblStatus.Text = string.Format("Making Mail for UUID: {0}", document.UUID);

                MailMessage mail = new MailMessage(Properties.Settings.Default.UserName, document.BusinessPartner.E_Mail, string.Format("{0} - Le ha enviado un comprobante fiscal digital ({1})", sap.CompanyName, SAP.DocNameByObjType(document.ObjType)), string.Format("Te informamos que {0} ha generado un Comprobante Fiscal Digital, el cual encontrarás adjunto en este correo en su formato XML y PDF", sap.CompanyName));
                mail.BodyEncoding = Encoding.UTF8;
                mail.IsBodyHtml = true;
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

                try
                {
                    foreach (SAP.BusinessPartner.ContactPerson contact_person in document.BusinessPartner.GetContactPersons())
                    {
                        mail.CC.Add(contact_person.E_Mail);
                    }
                } catch { }

                mail.Attachments.Add(new Attachment(string.Format(@"{0}\{1}\{2}.xml", Properties.Settings.Default.Path, document.MakePath(), document.UUID)));
                mail.Attachments.Add(new Attachment(string.Format(@"{0}\{1}\{2}.pdf", Properties.Settings.Default.Path, document.MakePath(), document.UUID)));

                return mail;
            }
            catch (Exception ex)
            {
                Program.WriteLog(ex.Message);
                return null;
            }
        }

        private void MakePDF(SAP.Document document)
        {
            try
            {
                lblStatus.Text = string.Format("Making PDF for UUID: {0}", document.UUID);

                ReportDocument cryRpt = new ReportDocument();
                cryRpt.Load(string.Format(@"{0}\{1}.rpt", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), SAP.SAPTableByObjType(document.ObjType)));
                cryRpt.DataSourceConnections[0].SetConnection(SAP.sql.DataSource, SAP.sql.InitialCatalog, SAP.sql.UserID, SAP.sql.Password);
                //Program.WriteLog(string.Format("Servidor {0} | BD {1} | UID {2} | PWD {3}", SAP.sql.DataSource, SAP.sql.InitialCatalog, SAP.sql.UserID, SAP.sql.Password));
                cryRpt.SetParameterValue("DocKey@", document.DocEntry);
                cryRpt.SetParameterValue("ObjectId@", document.ObjType);
                cryRpt.ExportToDisk(ExportFormatType.PortableDocFormat, string.Format(@"{0}\{1}\{2}.pdf", Properties.Settings.Default.Path, document.MakePath(), document.UUID));
                cryRpt.Close();
                cryRpt.Dispose();
            }
            catch (CrystalReportsException cex)
            {
                Program.WriteLog("PDF CEX: " + cex.Message);
            }
            catch (Exception ex)
            {
                Program.WriteLog("PDF EX: " + ex.Message);
            }
        }

        private void MoveFiles()
        {
            try
            {
                DirectoryInfo directory = new DirectoryInfo(@Properties.Settings.Default.Path);
                foreach (FileInfo file in directory.GetFiles())
                {
                    try
                    {
                        if (file.Extension == ".xml" || file.Extension == ".pdf")
                        {
                            SAP.Document document = sap.GetDocument(Path.GetFileNameWithoutExtension(file.Name));

                            if (document.DocEntry != 0)
                            {
                                string directory_path = string.Format(@"{0}\{1}\", Properties.Settings.Default.Path, document.MakePath());
                                string file_path = string.Format(@"{0}{1}{2}", directory_path, document.UUID, file.Extension);

                                if (!Directory.Exists(directory_path))
                                    Directory.CreateDirectory(directory_path);

                                File.Move(file.FullName, file_path);
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Program.WriteLog("MoveFiles: " + ex.Message);
            }
        }
    }
}