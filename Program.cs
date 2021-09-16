using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static MsgReader.Outlook.Storage;

namespace EML
{
    class Program
    {
         
        static void Main(string[] args)
        {

            Console.Write("Please  Enter Company:");
            string companyFolder = Console.ReadLine();
            Console.Write("Please  Enter Company output:");
            string companyFolderOut = Console.ReadLine();
            int i = 0;

            foreach (string fileName  in Directory.GetFiles(@"C:\Workspace\Work\OCRFiles\Input\"+companyFolder+"\\"))
            {
              
                string shortName = Path.GetFileName(fileName);
                shortName = shortName.Replace("from MAKRO", "");
                // Read a email .msg file
                Message message = new MsgReader.Outlook.Storage.Message(@fileName);
                // Read sender
                Console.WriteLine("Sender:" + message.Sender.Email);
                // Read sent on
                Console.WriteLine("SentOn:" + message.SentOn);
                // Read recipient to
                Console.WriteLine("recipientsTo:" + message.GetEmailRecipients(MsgReader.Outlook.RecipientType.To, false, false));
                // Read recipient cc
                Console.WriteLine("recipientsCc:" + message.GetEmailRecipients(MsgReader.Outlook.RecipientType.Cc, false, false));
                // Read subject
                Console.WriteLine("subject:" + message.Subject);
                // Read body html
                Console.WriteLine("htmlBody:" + message.BodyHtml);
                foreach (var attachment in message.Attachments)
                {
                    i++;

                    string outPutFileName = string.Empty;

                    if (attachment.GetType() == typeof(MsgReader.Outlook.Storage.Attachment))
                    {
                        var attach = (MsgReader.Outlook.Storage.Attachment)attachment;

                       // outPutFileName = Path.Combine(@"C:\Workspace\Work\OCRFiles\Output\", shortName+"_"+(attach).FileName);
                        outPutFileName = Path.Combine(@"C:\Workspace\Work\OCRFiles\Output\"+ companyFolderOut+"\\", i.ToString()+".pdf");
                       

                        File.WriteAllBytes(outPutFileName, attach.Data);

                    }
                }
         

            }
            Console.WriteLine(DateTime.Now.TimeOfDay.ToString());
            Console.ReadKey();



        }
        public void UploadFiles()
        {
            while (true)
            {
                foreach (string sFile in Directory.GetFiles(@"C:\Workspace\Work\OCRFiles\"))
                {
                    try
                    {
                        FileStream fs = File.Open(sFile, FileMode.Open, FileAccess.ReadWrite);
                        EMLReader reader = new EMLReader(fs);
                        Console.WriteLine(reader.HTMLBody);
                        fs.Close();

                        // .... Process EML file here
                       
                        Console.ReadKey();

                       // File.Delete(sFile);
                    }
                    catch (System.IO.IOException err)
                    {
                        Debug.WriteLine("File " + sFile + " is currently in use.");
                        Console.ReadKey();
                    }
                    Thread.Sleep(10);
                }
            }
        }
    }
}
