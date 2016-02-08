using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using WatiN.Core;

namespace SendShoreTelUsersHome
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            AES aes = new AES();

            //grab from the appSettings
            string username = AES.Decrypt(ConfigurationManager.AppSettings["username"]);
            string password = AES.Decrypt(ConfigurationManager.AppSettings["password"]);
            string shoreTelDirectorLogin = AES.Decrypt(ConfigurationManager.AppSettings["ShoreTelDirectorLogin"]);
            string shoreTelUserList = AES.Decrypt(ConfigurationManager.AppSettings["ShoreTelUserList"]);

            //nav to url
            IE myIE = new IE(shoreTelDirectorLogin);

            //sign in
            myIE.TextField(Find.ByName("login")).TypeText(username);
            myIE.TextField(Find.ByName("password")).TypeText(password);
            myIE.Button(Find.ByName("SUBMIT1")).Click();

            //nav to url containing userList
            IE myIE3 = new IE(shoreTelUserList);

            //change recPerPage to 3000 to ensure we grab all extensions
            myIE3.SelectList(Find.ById("RecPerPage")).Option("3000").Select();

            var userListTable = myIE3.Table(Find.By("_t", "Users"));

            //create/overwrite javascript file
            string jspath = @"C:\SendShoreTelUsersHome.js";
            File.Create(jspath).Close();
            StreamWriter sw = new StreamWriter(jspath);

            //Prepare the ports!
            sw.WriteLine("var ports = new ActiveXObject('ShoreBusDS.Ports');");

            //writing go home line for each extension
            foreach (var tableRow in userListTable.TableRows)
            {
                //sample text layout of each tr.Text: " Zac Glenn Headquarters Executives Personal 7777 7777 AB77 77 Zac Glenn Home "
                if (!string.IsNullOrEmpty(tableRow.Text))
                {
                    Match match = Regex.Match(tableRow.Text, @"\d{4}");
                    string extension = match.Value;

                    if (!string.IsNullOrEmpty(extension))
                    {
                        sw.WriteLine("ports.UserGoHome('" + extension + "');");
                    }
                }
            }
            //close file connection because exceptions arise that the file is already in use etc
            sw.Close();

            //create your cmd process to execute js file using wscript
            Process process = new Process();
            process.StartInfo.FileName = "wscript.exe";
            process.StartInfo.WorkingDirectory = @"c:\";
            process.StartInfo.Arguments = "SendShoreTelUsersHome.js";
            process.Start();

            //close everything we opened
            myIE.ForceClose();
        }
    }
}
