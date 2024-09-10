using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Net;
using log4net.Config;
using log4net;

namespace NotificationDocument
{
    internal class Program
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));
        public static string effectiveLabel = ConfigurationSettings.AppSettings["effectiveLabel"];
        public static string connectionString = ConfigurationSettings.AppSettings["connectionString"];
        public static string excludeRole = "ExcludeNotification";
        public static int memoId = 0;
        public static List<string> excludeRoles
        {
            get
            {
                var emails = (from role in dbContext.MSTRoles
                              join userPerm in dbContext.MSTUserPermissions on role.RoleId equals userPerm.RoleId
                              where role.NameEn.ToUpper() == excludeRole.ToUpper() || role.NameTh.ToUpper() == excludeRole.ToUpper()
                              join emp in dbContext.MSTEmployees on userPerm.EmployeeId equals emp.EmployeeId
                              select emp.Email).ToList();

                return emails;
            }
        }

        public static DbContextDataContext dbContext = new DbContextDataContext(connectionString);
        public static DateTime currentDate = DateTime.Now;
        static void Main(string[] args)
        {
            XmlConfigurator.Configure();
            log.Info($"=============================================================================================================");
            var currents = new List<string>()
            {
                currentDate.ToString("dd MMM yyyy"),
                currentDate.ToString("dd/MMM/yyyy"),
                currentDate.ToString("dd MM yyyy"),
                currentDate.ToString("dd/MM/yyyy")
            };

            log.Info($"Format Current Date : {string.Join(",", currents)}");

            var emails = dbContext.ViewEmployees.Where(x => !excludeRoles.Contains(x.Email)).Select(s => s.Email).ToList();

            var memos = dbContext.TRNMemos.Where(x => x.DocumentNo.Contains("DAR") &&
            dbContext.TRNMemoForms.Any(a => x.MemoId == a.MemoId && a.obj_label == effectiveLabel && currents.Contains(a.obj_value) )).ToList();

            log.Info($"Send Memo Count : {memos.Count()}");

            foreach ( var memo in memos )
            {
                memoId = memo.MemoId;
                var emailTemplateModel = dbContext.MSTEmailTemplates.FirstOrDefault(x => x.FormState == "NotificationDoc");

                var sURLToRequest = $"{ConfigurationSettings.AppSettings["TinyUrl"]}Request?MemoID={memo.MemoId}";


                emailTemplateModel.EmailSubject = ReplaceEmail(emailTemplateModel.EmailSubject, memo, sURLToRequest);
                emailTemplateModel.EmailBody = ReplaceEmail(emailTemplateModel.EmailBody, memo, sURLToRequest);
                SendEmail(emailTemplateModel.EmailBody, emailTemplateModel.EmailSubject, emails);
            }
            log.Info($"=============================================================================================================");
        }

        public static string ReplaceEmail(string content, TRNMemo memo, string sURLToRequest)
        {
            content = content

               .Replace("[TRNMemo_DocumentNo]", memo.DocumentNo)
               .Replace("[TRNMemo_TemplateSubject]", memo.TemplateSubject)
               .Replace("[TRNMemo_RNameEn]", memo.RNameEn)
               .Replace("[TRNMemo_RequestDate]", memo.RequestDate.Value.ToString("dd MMM yyyy"))
               .Replace("[TRNActionHistory_ActorName]", dbContext.ViewEmployees.FirstOrDefault(x => x.EmployeeId.ToString() == memo.LastActionBy)?.NameEn)

               .Replace("[Effective_Date]", currentDate.ToString("dd MMM yyyy"))
               .Replace("[TRNMemo_StatusName]", memo.StatusName)

               .Replace("[TRNMemo_CompanyName]", memo.CompanyName)
               .Replace("[TRNMemo_TemplateName]", memo.TemplateName)

               .Replace("[URLToRequest]", String.Format("<a href='{0}'>Click</a>", sURLToRequest));

            return content;
        }

        public static DateTime TruncateTime(DateTime dateTime)
        {
            return new DateTime(dateTime.Year, dateTime.Month, dateTime.Day);
        }

        public static void SendEmail(string emailBody, string emailSubject, List<string> toList)
        {
            string smtpServer = ConfigurationSettings.AppSettings["SMTPServer"];
            int smtpPort = Convert.ToInt32(ConfigurationSettings.AppSettings["SMPTPort"]);
            string fromEmail = ConfigurationSettings.AppSettings["SMTPUser"];
            string fromPassword = ConfigurationSettings.AppSettings["SMTPPassword"];

            try
            {
                using (SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort))
                {
                    smtpClient.Credentials = new NetworkCredential(fromEmail, fromPassword);
                    smtpClient.EnableSsl = true;

                    MailMessage mailMessage = new MailMessage
                    {
                        From = new MailAddress(fromEmail),
                        Subject = emailSubject,
                        Body = emailBody,
                        IsBodyHtml = true
                    };

                    foreach (string recipient in toList)
                    {
                        mailMessage.To.Add(new MailAddress(recipient));
                    }

                    smtpClient.Send(mailMessage);
                    log.Info($"Send MemoId : {memoId} : Email sent successfully.");
                }
            }
            catch (Exception ex)
            {
                log.Info($"Failed to send email. Error: {ex.Message}");
            }
        }
    }
}
