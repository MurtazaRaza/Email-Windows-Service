using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Timers;

namespace MailingService
{
    public partial class MailService : ServiceBase
    {
#pragma warning disable CS0618 // Type or member is obsolete
        private static readonly string FOLDER_PATH = ConfigurationSettings.AppSettings["File_Location"];
        private static readonly string USER_NAME = ConfigurationSettings.AppSettings["Username_Sender"];
        private static readonly string PASSWORD = ConfigurationSettings.AppSettings["Password_Sender"];
#pragma warning restore CS0618 // Type or member is obsolete

        private static readonly string FOLDER_PATH_SENT = FOLDER_PATH + @"Sent Mail\";
        private static readonly string FOLDER_PATH_LOG = FOLDER_PATH + @"MLog.txt";
        private String[] XMLpaths;
        System.Timers.Timer timeDelay;

        public MailService()
        {
            InitializeComponent();
            timeDelay = new System.Timers.Timer();
            timeDelay.Interval = 900000;                     //15 minutes
            timeDelay.Elapsed += new ElapsedEventHandler(ChainOfWorkerProcesses);
        }

        //Individual functions called through this function
        protected void ChainOfWorkerProcesses(object sender, ElapsedEventArgs e)
        {
            GetXMLFiles();
            ProcessXMLFiles();
            ClearPaths();
        }

        //Clear Path array to ensure that array is empty after every interval
        private void ClearPaths()
        {
            if (XMLpaths != null)
                Array.Clear(XMLpaths, 0, XMLpaths.Length);             //Clear Previously saved paths
        }

        //For debugging purposes only. Not for release.
        public void onDebug()
        {
            OnStart(null);
        }

        protected override void OnStart(string[] args)
        {
            LogService("Service starting! ");
            if(XMLpaths!=null)
                Array.Clear(XMLpaths, 0, XMLpaths.Length);             //Clear Previously saved paths
            timeDelay.Enabled = true;
        }
        
        //Gets all xml files path names from the folder path
        protected void GetXMLFiles()
        {
            try
            {
                XMLpaths = Directory.GetFiles(FOLDER_PATH, "*.xml", SearchOption.TopDirectoryOnly);
            }
            catch (Exception e)
            {
                LogService("Exception detected in GetXMLFiles()");
            }

        }

        //Gets each individual XML Node
        public static XmlNode GetXMLNode(XmlDocument doc, String nodePath)
        {
            XmlNode node = doc.DocumentElement.SelectSingleNode(nodePath);
            return node;
            
        }

        //Gets rid of unnecessary whitespaces
        public static String GetNodeString(String text)
        {
            text.Replace(" ", "");
            return text;
        }

        //Where XML is processed and if conditions are met, smtp mail is sent
        //followed by the xml mail being moved to sent folder
         protected void ProcessXMLFiles()
        {
            foreach (var path in XMLpaths)
            {
                
                XmlDocument doc = new XmlDocument();
                String To = null;
                String Subject = null;
                String Msg = null;

                bool sent = false;

                try
                {
                    doc = new XmlDocument();
                    doc.Load(path);
                }
                catch (Exception e)
                {
                    LogService("XML file io Exception on file: " + path);
                }

                try
                {
                    XmlNode node = GetXMLNode(doc, "/EmailMessage/To");
                    To = GetNodeString(node.InnerText);
                    if (To != null && To.Contains("@")) 
                    {
                        node = GetXMLNode(doc, "/EmailMessage/Subject");
                        Subject = GetNodeString(node.InnerText);
                        if (Subject != null)
                        {
                            node = GetXMLNode(doc, "/EmailMessage/MessageBody");
                            Msg = GetNodeString(node.InnerText);
                            if (Msg != null)
                            {
                                LogService("///////////////////////////////New mail\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ \n");
                                LogService("Mail should send now or somethingss wroong");
                                SmtpClient SmtpServer = new SmtpClient("smtp.live.com");
                                var mail = new MailMessage();

                                mail.From = new MailAddress(USER_NAME);
                                mail.To.Add(To);
                                mail.Subject = Subject;
                                mail.IsBodyHtml = true;
                                mail.Body = Msg;


                                try
                                {

                                    SmtpServer.Port = 587;
                                    SmtpServer.UseDefaultCredentials = false;
                                    SmtpServer.Credentials = new System.Net.NetworkCredential(USER_NAME, PASSWORD);
                                    SmtpServer.EnableSsl = true;
                                    SmtpServer.Send(mail);
                                    LogService("Mail sent!");
                                    sent = true;


                                }
                                catch (Exception inSendingMail)
                                {
                                    LogService("Invalid operations exception probably" + inSendingMail);
                                }
                                try
                                {
                                    if (sent)
                                    {
                                        LogService("Moving mail to Sent folder");
                                        if (!Directory.Exists(FOLDER_PATH_SENT))
                                        {
                                            Directory.CreateDirectory(FOLDER_PATH_SENT);

                                        }

                                        File.Copy(path, FOLDER_PATH_SENT + Path.GetFileName(path), true);
                                        File.Delete(path);
                                        LogService("Mail " + Path.GetFileName(path) + " moved successfully.");
                                    }
                                }
                                catch(Exception inCreatingOrTransferringFile)
                                {
                                    LogService("Problem creating sent folder or copying file");
                                }


                            }
                        }
                    }
                }
                catch (Exception inParsingOrSMTP)
                {
                    LogService("Exception in parsing XML or SMTP");
                }

            }
        }

        private static void LogService(string content)
        {
            try
            {


                FileStream fs = new FileStream(FOLDER_PATH_LOG, FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                sw.BaseStream.Seek(0, SeekOrigin.End);
                sw.WriteLine(content);
                sw.Flush();
                sw.Close();
            }
            catch(Exception e)
            {

            }
        }

        protected override void OnStop()
        {
            LogService("Service Stopping! ");
        }
    }
}
