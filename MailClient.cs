using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace PrintMailExcel
{
    public class MailClient : IMailClient
    {
        public string EmailAddress { get; }
        public string Password { get; }

        public MailClient(string emailAddress, string password)
        {
            EmailAddress = emailAddress;
            Password = password;
        }

        public bool TryGetExcelFilesInfo(string basePath, out List<ExcelFile> excelFilesInfo)
        {
            excelFilesInfo = new List<ExcelFile>();

            Application application = new Application();
            NameSpace nameSpace = application.GetNamespace("MAPI");
            nameSpace.Logon(EmailAddress, Password, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            List<MailItem> unreadMails = GetMails(nameSpace, unread: true);
            List< MailItem> readMails = GetMails(nameSpace, unread: false);

            foreach (var readMail in readMails)
            {
                readMail.Delete();
            }

            foreach (var unreadMail in unreadMails)
            {
                ExcelHandling excelHandling = new ExcelHandling();

                List<string> mimeTypesExcel = new List<string>
                    {
                        "xlsx",
                        "xls",
                        "csv"
                    };

                foreach (Attachment attachment in unreadMail.Attachments)
                {
                    if (mimeTypesExcel.Any(ending => attachment.FileName.EndsWith(ending)))
                    {
                        Guid fileGuid = Guid.NewGuid();
                        string tempAttachmentFullPath = Path.Combine(basePath, $"{fileGuid}_{attachment.FileName}");
                        attachment.SaveAsFile(tempAttachmentFullPath);
                        List<ExcelSheet> excelSheetsContent = excelHandling.GetExcelTables(tempAttachmentFullPath);
                        ExcelFile tempAttachment = new ExcelFile(excelSheetsContent, attachment.FileName, fileGuid);
                        excelFilesInfo.Add(tempAttachment);
                        if (File.Exists(tempAttachmentFullPath))
                            File.Delete(tempAttachmentFullPath);
                    }
                }
            }

            return excelFilesInfo.Any();
        }

        private List<MailItem> GetMails(NameSpace nameSpace, bool unread = true)
        {
            MAPIFolder inboxFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            List<MailItem> eMails = new List<MailItem>();

            foreach (MailItem mailItem in inboxFolder.Items)
            {
                if (unread ? mailItem.UnRead : !mailItem.UnRead)
                    eMails.Add(mailItem);
            }
            return eMails;
        }


        public static void PrintExcelFilesinfoRows(List<ExcelFile> excelFilesInfo)
        {
            var strBuilder = new StringBuilder();
            foreach (var excelFile in excelFilesInfo)
            {
                foreach (var sheet in excelFile.Sheets)
                {
                    strBuilder.AppendLine($"File: {excelFile.ExcelFileName}, SheetName: {sheet.SheetName}");
                    foreach (var row in sheet.Table)
                    {
                        strBuilder.AppendLine(String.Join(" ", row));
                    }
                    strBuilder.AppendLine();
                }
                strBuilder.AppendLine();
            }

            Console.Write(strBuilder.ToString());
        }

        public List<ExcelPrint> GetRows(List<ExcelFile> excelFilesInfo)
        {
            List<ExcelPrint> printList = new List<ExcelPrint>();
            foreach (var excelFile in excelFilesInfo)
            {
                foreach (var sheet in excelFile.Sheets)
                {
                    StringBuilder strBuilder = new StringBuilder();
                    strBuilder.AppendLine($"File: {excelFile.ExcelFileName}, SheetName: {sheet.SheetName}");
                    foreach (var row in sheet.Table)
                    {
                        strBuilder.AppendLine(String.Join(" ", row));
                    }
                    strBuilder.AppendLine();

                    printList.Add(new ExcelPrint(excelFile.ExcelFileName, sheet.SheetName, strBuilder.ToString()));
                }
            }
            return printList;
        }
    }
}