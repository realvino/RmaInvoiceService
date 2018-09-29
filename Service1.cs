using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.IO;
using System.Threading;
using System.Configuration;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Text;
using System.Net;

namespace WindowsServiceCS
{
    public partial class Service1 : ServiceBase
    {
        public string EmailId, password, smtpServer;
        public bool useSSL;
        public int Port, i = 0;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.WriteToFile("Simple Service started {0}");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            this.WriteToFile("Simple Service stopped {0}");
            this.Schedular.Dispose();
        }

    private Timer Schedular;

    public void ScheduleService()
    {
       System.Timers.Timer time = new System.Timers.Timer();

      time.Start();

      time.Interval = 60000;

      time.Elapsed += SchedularCallback;
    }

    private void SchedularCallback(object sender, System.Timers.ElapsedEventArgs e)
    {

        try
        {
            var ftpdirectory = new DirectoryInfo(@"C:\ftp\data\RMA\");

            var myFile = (from f in ftpdirectory.GetFiles("*.txt")
                          orderby f.CreationTime descending
                          select f).ToList();



            if (myFile.Count != 0)
            {
                foreach (var filename in myFile)
                {
                    WriteToFile("This is the file < " + filename.Name + " > the FTP Server");

                    string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                    string DbFileName = string.Empty;
                    string CreationTime = string.Empty;
                    SqlConnection conn1 = new SqlConnection(constr);
                    conn1.Open();
                    string Dbdata = "select Filename from Temp where Filename='" + filename.Name + "'";

                    SqlCommand cmd = new SqlCommand(Dbdata, conn1);
                    SqlDataReader dr1 = cmd.ExecuteReader();
                    if (dr1.Read())
                    {
                        DbFileName = dr1["Filename"].ToString();
                    }
                    conn1.Close();
                    

                    if (DbFileName == "")
                    {

                        using (WebClient ftpClient = new WebClient())
                        {

                            using (SqlConnection connnn = new SqlConnection(constr))
                            {
                                using (SqlCommand spcmd = new SqlCommand("sp_readFile", connnn))
                                {
                                    spcmd.CommandType = CommandType.StoredProcedure;

                                    spcmd.Parameters.Add("@Path", SqlDbType.VarChar).Value = @"C:\ftp\data\RMA\" + filename.Name;
                                    spcmd.Parameters.Add("@File", SqlDbType.VarChar).Value = filename.Name;
                                    connnn.Open();
                                    spcmd.ExecuteNonQuery();

                                    SqlConnection conFinal = new SqlConnection(constr);
                                    conFinal.Open();
                                    string qurfinal = "insert into Temp (Filename) values('" + filename.Name + "')";
                                    WriteToFile("insert into Temp (Filename) values('" + filename.Name + "')");
                                    SqlCommand cmdfinal = new SqlCommand(qurfinal, conFinal);
                                    cmdfinal.ExecuteNonQuery();
                                    conFinal.Close();
                                }
                            }

                        }
                    }

                    using (SqlConnection connnn2 = new SqlConnection(constr))
                    {
                        using (SqlCommand spcmd2 = new SqlCommand("sp_BlockedProductUpdate", connnn2))
                        {
                            spcmd2.CommandType = CommandType.StoredProcedure;
                            connnn2.Open();
                            spcmd2.ExecuteNonQuery();
                        }
                    }
                }
            }

        }
        catch (Exception  ee)
        {
            WriteToFile("There was an error connecting to the FTP Server " +ee);
        }
        }
  
        private void WriteToFile(string text)
        {
            string path = "C:\\hostinglog\\RmaInvoiceTestServiceLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }



     
    }
}
