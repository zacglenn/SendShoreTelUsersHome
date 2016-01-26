﻿using System;
using System.Configuration;
using System.Text.RegularExpressions;
using WatiN.Core;

namespace SendShoreTelUsersHome
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //grab from the appSettings
            string username = ConfigurationManager.AppSettings["username"];
            string password = ConfigurationManager.AppSettings["password"];
            string shoreTelDirectorLogin = ConfigurationManager.AppSettings["ShoreTelDirectorLogin"];
            string shoreTelUserList = ConfigurationManager.AppSettings["ShoreTelUserList"];

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

            //Prepare the ports!
            Console.WriteLine("var sports = new ActiveXObject('ShoreBusDS.sPorts');");

            foreach (var tableRow in userListTable.TableRows)
            {
                //sample text layout of each tr.Text: " Zac Glenn Headquarters Executives Personal 7777 7777 AB77 77 Zac Glenn Home "
                if (!string.IsNullOrEmpty(tableRow.Text))
                {
                    Match match = Regex.Match(tableRow.Text, @"\d{4}");
                    string extension = match.Value;
                    Console.WriteLine("sports.UserGoHome('" + extension + "');");
                }
            }
            //close everything we opened
            myIE.ForceClose();
        }
    }
}