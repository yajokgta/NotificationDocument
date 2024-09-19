using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Net;
using log4net.Config;
using log4net;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.IO;
using Newtonsoft.Json;

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
        public static int IntervalTime
        {
            get
            {
                var _config = ConfigurationSettings.AppSettings["IntervalTime"];
                return int.Parse(_config);
            }
        }

        public static bool ManualMode
        {
            get
            {
                var _config = ConfigurationSettings.AppSettings["ManualMode"];
                return bool.Parse(_config);
            }
        }
        public class SettingContentModel
        {
            public string ReplaceKey { get; set;}
            public string FormLabel { get; set;}
            public string Value { get; set; } = "";
        }

        public static List<SettingContentModel> ReadContents
        {
            get
            {
                var filePath = Path.Combine(Directory.GetCurrentDirectory(), "SettingContent.json");

                if (File.Exists(filePath))
                {
                    string jsonString = File.ReadAllText(filePath);
                    return JsonConvert.DeserializeObject<List<SettingContentModel>>(jsonString);
                }
                else
                {
                    Console.WriteLine("File not found: " + filePath);
                    return new List<SettingContentModel>();
                }
            }
        }

        public static List<SettingContentModel> SettingContents = ReadContents;

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

            var memos = new List<TRNMemo>();

            if (ManualMode)
            {
                var manuals = new List<string>();

                Console.WriteLine("Enter StartDate (Ex: 2024-01-31) :");
                var inputStartDate = Console.ReadLine();
                Console.WriteLine("Enter EndDate (Ex: 2024-01-31) :");
                var inputEndDate = Console.ReadLine();

                var startDate = GetDateByString(inputStartDate);
                var endDate = GetDateByString(inputEndDate);

                for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
                {
                    var addDays = new List<string>()
                    {
                        date.ToString("dd MMM yyyy"),
                        date.ToString("dd/MMM/yyyy"),
                        date.ToString("dd MM yyyy"),
                        date.ToString("dd/MM/yyyy")
                    };

                    manuals.AddRange(addDays);
                }

                memos = dbContext.TRNMemos.Where(x => x.DocumentNo.Contains("DAR") && x.StatusName == "Completed" &&
                dbContext.TRNMemoForms.Any(a => x.MemoId == a.MemoId && a.obj_label == effectiveLabel && manuals.Contains(a.obj_value))).ToList();
            }

            else
            {
                memos = dbContext.TRNMemos.Where(x => x.DocumentNo.Contains("DAR") && x.StatusName == "Completed" && x.ModifiedDate >= DateTime.Now.AddMinutes(IntervalTime) &&
                dbContext.TRNMemoForms.Any(a => x.MemoId == a.MemoId && a.obj_label == effectiveLabel && currents.Contains(a.obj_value))).ToList();
            }

            var emails = dbContext.ViewEmployees.Where(x => !excludeRoles.Contains(x.Email)).Select(s => s.Email).ToList();
            log.Info($"Send Memo Count : {memos.Count()}");
            var emailTemplateModel = dbContext.MSTEmailTemplates.FirstOrDefault(x => x.FormState == "NotificationDoc" && x.IsActive == true);

            foreach ( var memo in memos )
            {
                memoId = memo.MemoId;

                SettingContents.ForEach(x =>
                {
                    x.Value = getValueAdvanceForm(memo.MAdvancveForm, x.FormLabel);
                });

                var sURLToRequest = $"{ConfigurationSettings.AppSettings["TinyUrl"]}Request?MemoID={memo.MemoId}";

                var effectiveDate = getValueAdvanceForm(memo.MAdvancveForm, effectiveLabel);

                var emailSubject = ReplaceEmail(emailTemplateModel.EmailSubject, memo, sURLToRequest);
                var emailBody = ReplaceEmail(emailTemplateModel.EmailBody, memo, sURLToRequest);
                SendEmail(emailBody, emailSubject, emails);
            }

            log.Info($"=============================================================================================================");
        }
        public static DateTime GetDateByString(string str)
        {
            var infoDate = str.Split('-');
            return new DateTime(Convert.ToInt32(infoDate[0]), Convert.ToInt32(infoDate[1]), Convert.ToInt32(infoDate[2]));
        }

        public static string getValueAdvanceForm(string AdvanceForm, string label)
        {
            string setValue = "";
            JObject jsonAdvanceForm = JObject.Parse(AdvanceForm);
            if (jsonAdvanceForm.ContainsKey("items"))
            {
                JArray itemsArray = (JArray)jsonAdvanceForm["items"];
                foreach (JObject jItems in itemsArray)
                {
                    JArray jLayoutArray = (JArray)jItems["layout"];
                    foreach (JToken jLayout in jLayoutArray)
                    {
                        JObject jTemplate = (JObject)jLayout["template"];
                        var getLabel = (String)jTemplate["label"];
                        if (label == getLabel)
                        {
                            JObject jdata = (JObject)jLayout["data"];
                            if (jdata != null)
                            {
                                if (jdata["value"] != null) setValue = jdata["value"].ToString();
                            }
                            break;
                        }
                    }
                }
            }

            return setValue;
        }

        public static string ReplaceEmail(string content, TRNMemo memo, string sURLToRequest)
        {
            content = content
               .Replace("[TRNMemo_DocumentNo]", memo.DocumentNo)
               .Replace("[TRNMemo_TemplateSubject]", memo.TemplateSubject)
               .Replace("[TRNMemo_RNameEn]", memo.RNameEn)
               .Replace("[TRNMemo_StatusName]", memo.StatusName)
               .Replace("[TRNMemo_CompanyName]", memo.CompanyName)
               .Replace("[TRNMemo_TemplateName]", memo.TemplateName)
               .Replace("[URLToRequest]", String.Format("<a href='{0}'>Click</a>", sURLToRequest));

            //DynamicContent
            foreach(var setting in SettingContents)
            {
                content.Replace(setting.ReplaceKey, setting.Value);
            }

            return content;
        }

        public static DateTime TruncateTime(DateTime dateTime)
        {
            return new DateTime(dateTime.Year, dateTime.Month, dateTime.Day);
        }

        public static bool IsValidEmail(string email)
        {
            string emailRegex = @"^[^\s@]+@[^\s@]+\.[^\s@]+$";

            return Regex.IsMatch(email, emailRegex);
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
                        if (IsValidEmail(recipient))
                        {
                            mailMessage.To.Add(new MailAddress(recipient.Trim().ToLower()));
                        }
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
