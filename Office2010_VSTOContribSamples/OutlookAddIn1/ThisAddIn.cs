using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

using System.Data;
using OutlookAddIn1.Properties;
using MySql.Data.MySqlClient;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisApplication_SaveAttachments()
        {
            Outlook.MAPIFolder inBox = this.Application.ActiveExplorer()
                .Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items inBoxItems = inBox.Items;
            Outlook.MailItem newEmail = null;
            inBoxItems = inBoxItems.Restrict("[Unread] = true");

            try
            {
                foreach (object collectionItem in inBoxItems)
                {
                    newEmail = collectionItem as Outlook.MailItem;
                    if (newEmail != null)
                    {
                        if (newEmail.Attachments.Count > 0)
                        {
                            for (int i = 1; i <= newEmail
                               .Attachments.Count; i++)
                            {
                                newEmail.Attachments[i].SaveAsFile
                                    (@"C:\TestFileSave\" +
                                    newEmail.Attachments[i].FileName);
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                string errorInfo = (string)ex.Message.Substring(0, 11);
                if (errorInfo == "Cannot save")
                {
                    MessageBox.Show(@"Create Folder C:\TestFileSave");
                }
            }
        }

        MailItem _mailItem = null;

        private void ThisApplication_ItemRead()
        {
            if (_mailItem.Subject.Contains("abcd") == true)
            {
                MessageBox.Show("Item Read：" + _mailItem.Subject);
            }
            ThisApplication_OperateDataByEmail(_mailItem);
        }

        private void ThisApplication_ItemLoad(object Item)
        {
            if (Item is MailItem)
            {
                try
                {
                    MailItem mailItem = Item as MailItem;
                    mailItem.Read += new
                       ItemEvents_10_ReadEventHandler(ThisApplication_ItemRead);
                    _mailItem = mailItem;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString(),
                      "Exception",
                      MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        
        private void ThisApplication_NewMail(string EntryIDCollection)
        {
            Outlook.ApplicationClass outLookApp = new Outlook.ApplicationClass();
            NameSpace outLookNS = outLookApp.GetNamespace("MAPI");
            MAPIFolder outLookFolder = outLookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            string storeID = outLookFolder.StoreID;
            MailItem newEmail = (MailItem)outLookNS.GetItemFromID(EntryIDCollection, storeID);

            ThisApplication_OperateDataByEmail(newEmail);
        }

        private void ThisApplication_OperateDataByEmail(MailItem newEmail)
        {
            if (newEmail != null)
            {
                string basepath = @"D:\Working\VSTOContrib\Office2010_VSTOContribSamples\OutlookAddIn1\bin\EmailAttachments\company\";
                string ship_name = null;
                string report_type = null;

                for (int i = 1; i <= newEmail.Attachments.Count; i++)
                {
                    ship_name = newEmail.Attachments[i].FileName.Substring(0, 3);
                    report_type = newEmail.Attachments[i].FileName.Substring(3, 6);

                    //保存mdb附件
                    //filename: c:\import\asp\Voyrpt\aspvoyrpt201806.mdb
                    if (!Directory.Exists(basepath + ship_name + "\\" + report_type))
                    {
                        Directory.CreateDirectory(basepath + ship_name + "\\" + report_type);
                    }
                    if (!File.Exists(basepath + ship_name + "\\" + report_type + "\\" + newEmail.Attachments[i].FileName))
                    {
                        newEmail.Attachments[i].SaveAsFile
                            (basepath + ship_name + "\\" + report_type + "\\" + newEmail.Attachments[i].FileName);
                    }
                        
                    //解析mdb文件，写数据库
                    ThisApplication_WriteToDB(newEmail.SenderName, newEmail.ReceivedTime.ToString(),
                        basepath + ship_name + "\\" + report_type + "\\" + newEmail.Attachments[i].FileName);
                }
            }

        }

        public void ThisApplication_WriteToDB(string sender, string rTime, string path)
        {
            MySqlConnection conn = new MySqlConnection((new Settings()).worldConnectionString);
            MySqlCommand cmd = new MySqlCommand();

            try
            {
                conn.Open();
                cmd.Connection = conn;

                cmd.CommandText = "select * from city where id > 10";//"INSERT INTO myTable VALUES(NULL, @number, @text)";
                cmd.Prepare();

                //cmd.Parameters.AddWithValue("@number", 1);
                //cmd.Parameters.AddWithValue("@text", "One");

                //for (int i = 1; i <= 1000; i++)
                //{
                //    cmd.Parameters["@number"].Value = i;
                //    cmd.Parameters["@text"].Value = "A string value";
                //    cmd.ExecuteNonQuery();
                //}

                writeDBUsingStored10(cmd.ExecuteReader());
                conn.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error " + ex.Number + " has occurred: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void writeDBUsingStored10(MySqlDataReader reader)
        {
            try
            {
                string s = null;
                int d = 0;
                DataTable table = reader.GetSchemaTable();

                if (reader.HasRows)//HasRows判断reader中是否有数据
                {
                    while (reader.Read())  //Read()方法读取下一条记录，如果没有下一条，返回false,则表示读取完成
                    {
                        s += "\r\n====>row: " + d.ToString() + "," + reader.GetString(0) + "," + reader.GetString(1) + ","
                            + reader.GetString(2) + "," + reader.GetString(3) + "," + reader.GetString(4);
                        d++;
                    }
                }

                //foreach (DataRow row in table.Rows)
                //{
                //    s += "row: ====> columns" + table.Columns.Count.ToString() + ",row" + d.ToString() + "\r\n";
                //    d++;
                //    foreach (DataColumn col in table.Columns)
                //    {
                //        s += "---->    " + col.ColumnName + ": " + row[col].ToString();
                //    }
                //    s += "\r\n";
                //}

                MessageBox.Show("write to db ok.");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void writeDBUsingStored11()
        {
            MySqlConnection conn = new MySqlConnection();
            conn.ConnectionString = "server=localhost;user=root;database=employees;port=3306;password=******";
            MySqlCommand cmd = new MySqlCommand();

            try
            {
                Console.WriteLine("Connecting to MySQL...");
                conn.Open();
                cmd.Connection = conn;
                cmd.CommandText = "DROP PROCEDURE IF EXISTS add_emp";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "DROP TABLE IF EXISTS emp";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "CREATE TABLE emp (empno INT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY, first_name VARCHAR(20), last_name VARCHAR(20), birthdate DATE)";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "CREATE PROCEDURE add_emp(" +
                                  "IN fname VARCHAR(20), IN lname VARCHAR(20), IN bday DATETIME, OUT empno INT)" +
                                  "BEGIN INSERT INTO emp(first_name, last_name, birthdate) " +
                                  "VALUES(fname, lname, DATE(bday)); SET empno = LAST_INSERT_ID(); END";

                cmd.ExecuteNonQuery();
            }
            catch (MySqlException ex)
            {
                Console.WriteLine("Error " + ex.Number + " has occurred: " + ex.Message);
            }
            conn.Close();
            Console.WriteLine("Connection closed.");
            try
            {
                Console.WriteLine("Connecting to MySQL...");
                conn.Open();
                cmd.Connection = conn;

                cmd.CommandText = "add_emp";
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@lname", "Jones");
                cmd.Parameters["@lname"].Direction = ParameterDirection.Input;

                cmd.Parameters.AddWithValue("@fname", "Tom");
                cmd.Parameters["@fname"].Direction = ParameterDirection.Input;

                cmd.Parameters.AddWithValue("@bday", "1940-06-07");
                cmd.Parameters["@bday"].Direction = ParameterDirection.Input;

                cmd.Parameters.AddWithValue("@empno", MySqlDbType.Int32);
                cmd.Parameters["@empno"].Direction = ParameterDirection.Output;

                cmd.ExecuteNonQuery();

                Console.WriteLine("Employee number: " + cmd.Parameters["@empno"].Value);
                Console.WriteLine("Birthday: " + cmd.Parameters["@bday"].Value);
            }
            catch (MySqlException ex)
            {
                Console.WriteLine("Error " + ex.Number + " has occurred: " + ex.Message);
            }
            conn.Close();
            Console.WriteLine("Done.");
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.NewMailEx += new ApplicationEvents_11_NewMailExEventHandler(ThisApplication_NewMail);
            this.Application.ItemLoad += new ApplicationEvents_11_ItemLoadEventHandler(ThisApplication_ItemLoad);
            MessageBox.Show("开始监听Outlook邮件！");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
