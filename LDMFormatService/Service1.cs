using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Timers;
using System.Data.SqlClient;

namespace LDMFormatService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
        string rootPathIn, rootPathOut, rootPathBackUp, rootPathBackUpIncorrect;
        RegistryKey key;
        string connetionString;
        SqlConnection conn;
        SqlCommand cmd;

        public Service1()
        {
            InitializeComponent();
            key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry32)
                .OpenSubKey(@"SOFTWARE\LDM", true);
            connetionString = @"Data Source=THBKKWNB1010214\SQLEXPRESS;Initial Catalog=dhl_ldm;User ID=sa;Password=P@ssw0rd!1234567890";
            conn = new SqlConnection(connetionString);
        }

        protected void checkDBConnectioning(string POS, string DST, string ULD_Type, string TARE, string NETT, string TOTAL, string FLAG, string FLAG2, string VOL)
        {
            try
            {
                if(conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }

                string sql = "INSERT INTO m_ldm (POS,DST,ULD_Type,TARE,NETT,TOTAL,FLAG,FLAG2,VOL) VALUES (@POS,@DST,@ULD_Type,@TARE,@NETT,@TOTAL,@FLAG,@FLAG2,@VOL)";
                cmd = new SqlCommand(sql,conn);
                cmd.Parameters.AddWithValue("@POS",POS.ToString());
                cmd.Parameters.AddWithValue("@DST", DST.ToString());
                cmd.Parameters.AddWithValue("@ULD_Type", ULD_Type.ToString());
                cmd.Parameters.AddWithValue("@TARE", TARE.ToString());
                cmd.Parameters.AddWithValue("@NETT", NETT.ToString());
                cmd.Parameters.AddWithValue("@TOTAL", TOTAL.ToString());
                cmd.Parameters.AddWithValue("@FLAG", FLAG.ToString());
                cmd.Parameters.AddWithValue("@FLAG2", FLAG2.ToString());
                cmd.Parameters.AddWithValue("@VOL", VOL.ToString());
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                WriteToFile("INFO: Data is writed to database already.");
            }
            catch (Exception ex)
            {
                WriteToFile("ERR: DB"+ex.Message);
            }
        }

        protected override void OnStart(string[] args)
        {
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 180000;
            timer.Enabled = true;
        }

        public void OnDebug()
        {
            OnStart(null);
        }

        protected override void OnStop()
        {
            timer.Stop();
        }

        public bool clearDataBeforeAddToDB(string source)
        {
            StreamWriter sw;
            string[] lines = File.ReadAllLines(Path.Combine(rootPathIn, source));
            if (lines.Length == 0 || lines[3].Split('.').Length != 3)
            {
                WriteToFile("WAR:File isn't format that want.");
                return false;
            }
            sw = File.CreateText(Path.Combine(rootPathOut,source));
            WriteToFile("DATA:"+ lines[3]);
            foreach (string line in lines.Skip(7).ToArray())
            {
                if (line.Trim().Equals(""))
                {
                    sw.Close();
                    break;
                }
                Regex re = new Regex(@"\b(\w*LOOSE|VOID\w*)\b", RegexOptions.IgnoreCase);
                if (!re.IsMatch(line))
                {
                    string txt_replace = Regex.Replace(line.Trim(), @":", "");
                    txt_replace = Regex.Replace(txt_replace.Trim(), @"\*", "");
                    txt_replace = Regex.Replace(txt_replace.Trim(), @"\s+", ",");
                    string[] txt_temp = txt_replace.Split(',');
                    //WriteToFile("DEBUG:"+txt_replace);
                    string ULD_Type = txt_temp[2].Substring(0, 3);
                    string FLAG = (txt_temp.Length == 8) ? txt_temp[txt_temp.Length - 2] : txt_temp[txt_temp.Length - 3];
                    string FLAG2 = "";
                    string VOL = txt_temp.Last();

                    re = new Regex(@"ORG...", RegexOptions.IgnoreCase);
                    Match m = re.Match(FLAG);
                    if (re.IsMatch(FLAG) && FLAG.Length > 10)
                    {
                        FLAG2 = m.Value;
                    }
                    else if (FLAG.Length < 10)
                    {
                        FLAG2 = FLAG;
                    }
                    else
                    {
                        FLAG2 = "N/A";
                    }
                    txt_replace = txt_temp[0] + "," + txt_temp[1] + "," + ULD_Type + "," + txt_temp[3] + "," + txt_temp[4] + "," + txt_temp[5] + "," + txt_temp[6] + "," + FLAG2 + "," + VOL;
                    checkDBConnectioning(txt_temp[0],txt_temp[1],ULD_Type,txt_temp[3],txt_temp[4],txt_temp[5],txt_temp[6],FLAG2,VOL);
                    sw.WriteLine(txt_replace);
                }
                else { }
            }
            sw.Close();
            WriteToFile("LOG:Cleaning and Backup is finished.");
            return true;
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            try
            {
                rootPathIn = key.GetValue("INPUT_DIRECTORY").ToString(); //@"C:\Users\hubtraining\Desktop\LDM_FORMAT\LDM_INDATA" 
                rootPathOut = key.GetValue("OUTPUT_DIRECTORY").ToString(); //@"C:\Users\hubtraining\Desktop\LDM_FORMAT\LDM_OUTDATA"
                rootPathBackUp = key.GetValue("BACKUP_DIRECTORY").ToString(); //@"C:\Users\hubtraining\Desktop\LDM_FORMAT\LDM_BACKUP"
                rootPathBackUpIncorrect = key.GetValue("BACKUP_INCORECTFORMAT_DIRECTORY").ToString(); //@"C:\Users\hubtraining\Desktop\LDM_FORMAT\LDM_BACKUP\Incorrect_Format"
            }
            catch
            {

            }
            WriteToFile("LOG:---------- Start Process ----------");
            if (!isDirectoryEmpty())
            {
                WriteToFile("LOG:Directory is has file.\n" +
                    "LOG:Going to load file for clearning.\n" +
                    "LOG:File is processing...\n");
                string[] files = Directory.GetFiles(rootPathIn, "*.txt", SearchOption.AllDirectories);
                foreach (string file in files)
                {
                    try
                    {
                        string fileName = Path.GetFileName(file);
                        bool success = clearDataBeforeAddToDB(fileName);
                        if (success == true)
                        {
                            File.Move(file,Path.Combine(rootPathBackUp,fileName));
                            WriteToFile(file+","+ Path.Combine(rootPathBackUp, fileName));
                        }
                        else
                        {
                            File.Move(file, Path.Combine(rootPathBackUpIncorrect, fileName));
                        }
                    }
                    catch (Exception ex) 
                    {
                        WriteToFile("ERR:"+ex.Message);
                    }
                }

            }
            else
            {
                WriteToFile("INFO:Directory is hasn't file.");
            }
            WriteToFile("LOG:---------- End Process ----------");
        }

        private bool isDirectoryEmpty()
        {
            return Directory.GetFileSystemEntries(rootPathIn).Length == 0;
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
    }
}