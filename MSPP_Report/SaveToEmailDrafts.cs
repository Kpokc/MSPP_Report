using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace MSPP_Report
{
    class SaveToEmailDrafts
    {
        public void addFileToEmail(string newPathWay, string today)
        {
            string[] emailsList = 
                { 
                "MSPP", 
                "Medtronic (Spare Parts) Report", 
                "Hi All", 
                "", /// Email to TO:
                ""  /// Emails to CC:
                };

            Application outlookApp = new Application();
            MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem);

            string bigLogo = @"V:\Warehouses\Tax Warehouse\logos\BigLogo.png";
            string bigLogoLink = "sslirl.com";

            mailItem.Attachments.Add(newPathWay); /// Excel report to attach
            mailItem.To = emailsList[3];
            mailItem.CC = emailsList[4];
            mailItem.Subject = emailsList[1] + " " + today;

            mailItem.HTMLBody = string.Format("<body>" + emailsList[2] + "," +
                                                        "<p>Please find attached " + emailsList[1] +
                                                        ".</p>" +
                                                        "<p>Regards,<br>" +
                                                        "<span style=\"font-size: 12.0pt;color:navy\"> Your Name Here </span>" +
                                                        "<br><span style=\"font-size: 8.0pt;color:#1F497D\">Source & Supply Logistics Ltd</span>" +
                                                        "<br><span style=\"font-size: 8.0pt;color:#1F497D\">IDA Business & Technology Park</span>" +
                                                        "<br><span style=\"font-size: 8.0pt;color:#1F497D\"" + ">Parkmore West</span>" +
                                                        "<br><span style=\"font-size: 8.0pt;color:#1F497D\">Galway, Ireland.</span>" +
                                                        "<br><span style=\"font-size: 9.0pt;color:navy;font-weight:bold\">Phone: </span>" + "  " +
                                                        "<span style=\"font-size: 9.0pt;color:navy;font-weight:bold\">Fax: </span>" +
                                                        "<br><span style=\"font-size: 9.0pt;color:navy;font-weight:bold\">Email:</span> " +
                                                        "<span style=\"font-size: 9.0pt;color:blue;font-weight:bold\"> Your Email Here </span>" +
                                                        "<p></body><a href=\"{1}\"><img src=\"{0}\"></a></p>", bigLogo, bigLogoLink);
            mailItem.Save();
        }
    }
}
