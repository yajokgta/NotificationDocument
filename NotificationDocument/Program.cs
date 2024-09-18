using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Net;
using log4net.Config;
using log4net;
using System.Runtime.InteropServices.ComTypes;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Data.Linq;

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
                var roles = ConfigurationManager.AppSettings["ExcludeRole"].Split(',').ToList();
                var emails = (from role in dbContext.MSTRoles
                              join userPerm in dbContext.MSTUserPermissions on role.RoleId equals userPerm.RoleId
                              where roles.Contains(role.NameEn) || roles.Contains(role.NameTh)
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
        private static string DocumentNumber
        {
            get
            {
                var DocumentNumberS = ConfigurationManager.AppSettings["DocumentNumber"];
                if (!string.IsNullOrEmpty(DocumentNumberS))
                {
                    return DocumentNumberS;
                }
                return "";
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
                currentDate.ToString("dd/MM/yyyy"),
                currentDate.AddDays(-1).ToString("dd MMM yyyy"),
                currentDate.AddDays(-1).ToString("dd/MMM/yyyy"),
                currentDate.AddDays(-1).ToString("dd MM yyyy"),
                currentDate.AddDays(-1).ToString("dd/MM/yyyy")
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
                memos = dbContext.TRNMemos.Where(x => x.DocumentNo.Contains("DAR") && x.StatusName == "Completed" &&
                dbContext.TRNMemoForms.Any(a => x.MemoId == a.MemoId && a.obj_label == effectiveLabel && currents.Contains(a.obj_value))).ToList();
            }

            log.Info($"Send Memo Count : {memos.Count()}");

            var emailTemplateModel = dbContext.MSTEmailTemplates.FirstOrDefault(x => x.FormState == "NotificationDoc");

            var viewEmployeeQuery = dbContext.ViewEmployees.Where(x => x.IsActive == true);

            foreach ( var memo in memos )
            {
                var employees = new List<ViewEmployee>();
                var additionalEmployees = new List<ViewEmployee>();
                var tempCcPersonsList = new List<ViewEmployee>();

                var memoLineApproves = dbContext.TRNLineApproves.Where(y => y.MemoId == memo.MemoId).ToList();

                var buGroup = getValueAdvanceForm(memo.MAdvancveForm, "Business Group");
                var department = getValueAdvanceForm(memo.MAdvancveForm, "Department");
                var documentNumber = getValueAdvanceForm(memo.MAdvancveForm, DocumentNumber);
                var promulgation = getValueAdvanceForm(memo.MAdvancveForm, "การประกาศใช้");

                log.Info("Documentnumber : " + documentNumber);
                log.Info("MemoId : " + memo.MemoId);
                log.Info("promulgation : " + promulgation);

                var effectiveDate = getValueAdvanceForm(memo.MAdvancveForm, effectiveLabel);

                var listBU = dbContext.TRNMemoForms.Where(x => x.MemoId == memo.MemoId && x.obj_label == "หน่วยงานที่เกี่ยวข้อง" && x.col_label == "หน่วยงาน").Select(s => s.col_value).ToList();
                var listEmployee = dbContext.TRNMemoForms.Where(x => x.MemoId == memo.MemoId && x.obj_label == "กรณีเฉพาะบุคคลที่เกี่ยวข้อง" && x.col_label == "ชื่อผู้เกี่ยวข้อง").Select(s => s.col_value).ToList();

                string AdditionalEmp = "";
                var ccPersonDistinct = new List<string>();
                var ccPersonData = "";

                if (promulgation == "ทุกคนทั้งองค์กร")
                {
                    employees.AddRange(dbContext.ViewBUs.Where(x => x.BUDESC == buGroup)
                                .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                    additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.RequesterId).FirstOrDefault());
                    additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.CreatorId).FirstOrDefault());

                    var ccPersonNames = memo.CcPerson.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    ccPersonNames.ForEach(x => x.Trim());

                    employees.AddRange(viewEmployeeQuery.Where(x => ccPersonNames.Contains(x.NameEn) || ccPersonNames.Contains(x.NameTh)).ToList());

                    AdditionalEmp = string.Join(",", additionalEmployees.Select(nameth => nameth.NameTh));
                    ccPersonDistinct = employees.Where(employee => !additionalEmployees.Select(ae => ae.NameTh).Contains(employee.NameTh)).Select(employee => employee.NameTh).ToList();
                    ccPersonData = string.Join(",", ccPersonDistinct);
                }
                else if (promulgation == "เฉพาะหน่วยงาน")
                {
                    if (!string.IsNullOrEmpty(buGroup))
                    {
                        foreach (var lineapprove in memoLineApproves)
                        {
                            var emp = viewEmployeeQuery.Where(v => v.EmployeeId == lineapprove.EmployeeId).FirstOrDefault();
                            employees.Add(emp);
                        }

                        employees.AddRange(dbContext.ViewBUs.Where(x => x.BUDESC == buGroup && x.DepartmentNameEn.Contains(department) || x.DepartmentNameTh.Contains(department))
                                    .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.RequesterId).FirstOrDefault());
                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.CreatorId).FirstOrDefault());

                        var ccPersonNames = memo.CcPerson.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        ccPersonNames.ForEach(x => x.Trim());

                        employees.AddRange(viewEmployeeQuery.Where(x => ccPersonNames.Contains(x.NameEn) || ccPersonNames.Contains(x.NameTh)).ToList());

                        AdditionalEmp = string.Join(",", additionalEmployees.Select(nameth => nameth.NameTh));
                        ccPersonDistinct = employees.Where(employee => !additionalEmployees.Select(ae => ae.NameTh).Contains(employee.NameTh)).Select(employee => employee.NameTh).ToList();
                        ccPersonData = string.Join(",", ccPersonDistinct);
                    }
                    else
                    {
                        foreach (var lineapprove in memoLineApproves)
                        {
                            var emp = viewEmployeeQuery.Where(v => v.EmployeeId == lineapprove.EmployeeId).FirstOrDefault();
                            employees.Add(emp);
                        }

                        var adds = dbContext.ViewBUs.Where(x => x.DepartmentNameEn.Contains(department) || x.DepartmentNameTh.Contains(department))
                            .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList();

                        employees.AddRange(adds);

                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.RequesterId).FirstOrDefault());
                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.CreatorId).FirstOrDefault());

                        var ccPersonNames = memo.CcPerson.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        ccPersonNames.ForEach(x => x.Trim());

                        employees.AddRange(viewEmployeeQuery.Where(x => ccPersonNames.Contains(x.NameEn) || ccPersonNames.Contains(x.NameTh)).ToList());

                        AdditionalEmp = string.Join(",", additionalEmployees.Select(nameth => nameth.NameTh));
                        ccPersonDistinct = employees.Where(employee => !additionalEmployees.Select(ae => ae.NameTh).Contains(employee.NameTh)).Select(employee => employee.NameTh).ToList();
                        ccPersonData = string.Join(",", ccPersonDistinct);
                    }

                }
                else if (promulgation == "เฉพาะบุคคล")
                {
                    if (!string.IsNullOrEmpty(buGroup))
                    {
                        foreach (var lineapprove in memoLineApproves)
                        {
                            var emps = viewEmployeeQuery.Where(v => v.EmployeeId == lineapprove.EmployeeId).ToList();
                            employees.AddRange(emps);
                        }

                        employees.AddRange(dbContext.ViewBUs.Where(x => x.BUDESC == buGroup && x.DepartmentNameEn.Contains(department) || x.DepartmentNameTh.Contains(department))
                            .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());
                        employees.AddRange(viewEmployeeQuery.Where(x => listEmployee.Contains(x.NameEn) || listEmployee.Contains(x.NameTh)).ToList());

                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.RequesterId).FirstOrDefault());
                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.CreatorId).FirstOrDefault());

                        var ccPersonNames = memo.CcPerson.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        ccPersonNames.ForEach(x => x.Trim());

                        employees.AddRange(viewEmployeeQuery.Where(x => ccPersonNames.Contains(x.NameEn) || ccPersonNames.Contains(x.NameTh)).ToList());

                        AdditionalEmp = string.Join(",", additionalEmployees.Select(nameth => nameth.NameTh));
                        ccPersonDistinct = employees.Where(employee => !additionalEmployees.Select(ae => ae.NameTh).Contains(employee.NameTh)).Select(employee => employee.NameTh).ToList();
                        ccPersonData = string.Join(",", ccPersonDistinct);
                    }
                    else
                    {
                        foreach (var lineapprove in memoLineApproves)
                        {
                            var emps = viewEmployeeQuery.Where(v => v.EmployeeId == lineapprove.EmployeeId).ToList();
                            employees.AddRange(emps);
                        }

                        employees.AddRange(dbContext.ViewBUs.Where(x => x.DepartmentNameEn.Contains(department) || x.DepartmentNameTh.Contains(department))
                            .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        employees.AddRange(viewEmployeeQuery.Where(x => listEmployee.Contains(x.NameEn) || listEmployee.Contains(x.NameTh)).ToList());

                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.RequesterId).FirstOrDefault());
                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.CreatorId).FirstOrDefault());

                        var ccPersonNames = memo.CcPerson.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        ccPersonNames.ForEach(x => x.Trim());

                        employees.AddRange(viewEmployeeQuery.Where(x => ccPersonNames.Contains(x.NameEn) || ccPersonNames.Contains(x.NameTh)).ToList());

                        AdditionalEmp = string.Join(",", additionalEmployees.Select(nameth => nameth.NameTh));
                        ccPersonDistinct = employees.Where(employee => !additionalEmployees.Select(ae => ae.NameTh).Contains(employee.NameTh)).Select(employee => employee.NameTh).ToList();
                        ccPersonData = string.Join(",", ccPersonDistinct);

                    }
                    
                }
                else if (promulgation == "--select--")
                {
                    if (!string.IsNullOrEmpty(buGroup))
                    {
                        foreach (var lineapprove in memoLineApproves)
                        {
                            var emps = viewEmployeeQuery.Where(v => v.EmployeeId == lineapprove.EmployeeId).ToList();
                            employees.AddRange(emps);
                        }

                        employees.AddRange(dbContext.ViewBUs.Where(x => x.BUDESC == buGroup && x.DepartmentNameEn.Contains(department) || x.DepartmentNameTh.Contains(department))
                            .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        employees.AddRange(dbContext.MSTDepartments.Where(x => listBU.Contains(x.NameEn) || listBU.Contains(x.NameTh))
                                    .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        employees.AddRange(viewEmployeeQuery.Where(x => listEmployee.Contains(x.NameEn) || listEmployee.Contains(x.NameTh)).ToList());

                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.RequesterId).FirstOrDefault());
                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.CreatorId).FirstOrDefault());

                        var ccPersonNames = memo.CcPerson.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        ccPersonNames.ForEach(x => x.Trim());

                        employees.AddRange(viewEmployeeQuery.Where(x => ccPersonNames.Contains(x.NameEn) || ccPersonNames.Contains(x.NameTh)).ToList());

                        AdditionalEmp = string.Join(",", additionalEmployees.Select(nameth => nameth.NameTh));
                        ccPersonDistinct = employees.Where(employee => !additionalEmployees.Select(ae => ae.NameTh).Contains(employee.NameTh)).Select(employee => employee.NameTh).ToList();
                        ccPersonData = string.Join(",", ccPersonDistinct);
                    }
                    else
                    {
                        foreach (var lineapprove in memoLineApproves)
                        {
                            var emps = viewEmployeeQuery.Where(v => v.EmployeeId == lineapprove.EmployeeId).ToList();
                            employees.AddRange(emps);
                        }

                        employees.AddRange(dbContext.ViewBUs.Where(x => x.DepartmentNameEn.Contains(department) || x.DepartmentNameTh.Contains(department))
                           .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        employees.AddRange(viewEmployeeQuery.Where(x => listEmployee.Contains(x.NameEn) || listEmployee.Contains(x.NameTh)).ToList());

                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.RequesterId).FirstOrDefault());
                        additionalEmployees.Add(viewEmployeeQuery.Where(e => e.EmployeeId == memo.CreatorId).FirstOrDefault());

                        var ccPersonNames = memo.CcPerson.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        ccPersonNames.ForEach(x => x.Trim());

                        employees.AddRange(viewEmployeeQuery.Where(x => ccPersonNames.Contains(x.NameEn) || ccPersonNames.Contains(x.NameTh)).ToList());

                        AdditionalEmp = string.Join(",", additionalEmployees.Select(nameth => nameth.NameTh));
                        ccPersonDistinct = employees.Where(employee => !additionalEmployees.Select(ae => ae.NameTh).Contains(employee.NameTh)).Select(employee => employee.NameTh).ToList();
                        ccPersonData = string.Join(",", ccPersonDistinct);
                    }
                }
                else
                {
                    log.Info("Invalid promulgation");
                }
                
                SendEmail(employees, memo, documentNumber, ccPersonData, AdditionalEmp);
                log.Info("All Email :" + employees.Count);
                log.Info("--------------------------");

            }
            log.Info($"=============================================================================================================");
        }

        public static void SendEmail(List<ViewEmployee> employees, TRNMemo memo, String Documentnumber, string ccPersonData, string additionalEmployees)
        {
            employees.RemoveAll(x => excludeRoles.Contains(x.Email));

            log.Info($"Send : {string.Join("|", employees.Select(s => s.Email))}");
            var AllEmployee = dbContext.ViewEmployees.Where(x => x.IsActive == true);
            DateTime sentDatetime = DateTime.Now;
            string formattedDate = sentDatetime.ToString("dddd, MMMM dd, yyyy h:mm:ss tt");
            string requestDateTimeString = memo.RequestDate.HasValue ? memo.RequestDate.Value.ToString("M/d/yyyy h:mm:ss tt") : "-";
            string toPersons = additionalEmployees;
            string[] topersonNames = toPersons.Split(',');
            string CcPerson = ccPersonData;
            string[] ccpersonNames = CcPerson.Split(',');
            List<string> toEmailStrings = new List<string>();
            List<string> CcEmailStrings = new List<string>();
            int lastActionByEmployeeId = int.Parse(memo.LastActionBy);
            List<TRNActionHistory> Actionprocess = new List<TRNActionHistory>();

            string nameTh = "-";
            ViewEmployee EmployeeIdbyLast = AllEmployee.FirstOrDefault(x => x.EmployeeId == lastActionByEmployeeId);
            if (EmployeeIdbyLast != null)
            {
                nameTh = EmployeeIdbyLast.NameTh;
            }

            string actionProcess = null;
            string actionDate = null;
            string actionComment = null;
            TRNActionHistory latestAction = dbContext.TRNActionHistories.Where(action => action.MemoId == memo.MemoId).OrderByDescending(action => action.ActionDate).FirstOrDefault();
            if (latestAction != null)
            {
                actionProcess = latestAction.ActionProcess;
                actionDate = latestAction.ActionDate.HasValue ? latestAction.ActionDate.Value.ToString("dd/MM/yyyy") : "-";
                actionComment = latestAction.Comment;
            }

            foreach (string topersonName in topersonNames)
            {
                ViewEmployee TomatchingEmployee = AllEmployee.FirstOrDefault(e => e.NameTh == topersonName.Trim());
                if (TomatchingEmployee != null)
                {
                    string email = TomatchingEmployee.Email;
                    string nameEn = TomatchingEmployee.NameEn;
                    string toEmailString = $"{nameEn}&lt;<a href='mailto:{email}'>{email}</a>&gt;";
                    toEmailStrings.Add(toEmailString);
                }
            }

            foreach (string ccpersonName in ccpersonNames)
            {
                ViewEmployee CcmatchingEmployee = AllEmployee.FirstOrDefault(e => e.NameTh == ccpersonName.Trim());
                if (CcmatchingEmployee != null)
                {
                    string email = CcmatchingEmployee.Email;
                    string nameEn = CcmatchingEmployee.NameEn;
                    string ccEmailString = $"{nameEn}&lt;<a href='mailto:{email}'>{email}</a>&gt;";
                    CcEmailStrings.Add(ccEmailString);
                }
            }

            string memoDocumentCode = string.Empty;
            if (memo.DocumentCode != null)
            {
                memoDocumentCode = $" #{Documentnumber}";
            }

            string smtpServer = ConfigurationSettings.AppSettings["SMTPServer"];
            int smtpPort = Convert.ToInt32(ConfigurationSettings.AppSettings["SMPTPort"]);
            string fromEmail = ConfigurationSettings.AppSettings["SMTPUser"];
            string fromPassword = ConfigurationSettings.AppSettings["SMTPPassword"];

            string toEmailsString = string.Join(", ", toEmailStrings);
            toEmailsString = string.IsNullOrEmpty(toEmailsString) ? "-" : toEmailsString;
            string ccEmailsString = string.Join("; ", CcEmailStrings);
            ccEmailsString = string.IsNullOrEmpty(ccEmailsString) ? "-" : ccEmailsString;
            actionProcess = string.IsNullOrEmpty(actionProcess) ? "-" : actionProcess;
            actionComment = string.IsNullOrEmpty(actionComment) ? "-" : actionComment;
            memo.StatusName = memo.StatusName ?? "-";
            Documentnumber = Documentnumber ?? "-";
            memo.MemoSubject = memo.MemoSubject ?? "-";
            memo.RNameTh = memo.RNameTh ?? "-";

            var sURLToRequest = $"{ConfigurationSettings.AppSettings["TinyUrl"]}Request?MemoID={memo.MemoId}";

            string subject = $"Wolf ISO: {memo.StatusName}{memoDocumentCode} : {memo.MemoSubject}";
            string body = $"<b>From:</b> admin opr&lt;<a href='mailto:{fromEmail}'>{fromEmail}</a>&gt;<br>" +
                          $"<b>Sent:</b> {formattedDate}<br>" +
                          $"<b>To:</b> {toEmailsString}<br>" +
                          $"<b>Cc:</b> {ccEmailsString}<br>" +
                          $"<b>Subject:</b> Wolf ISO: {memo.StatusName}{memoDocumentCode} : {memo.MemoSubject}<br><br>" +
                          $"Dear All<br><br>" +
                          $"Please be informed than document as detail below has been completed:<br><br>" +
                          $"Document No.&emsp;&emsp;&emsp;&emsp;&ensp;: {Documentnumber}<br><br>" +
                          $"Subject&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&ensp;&nbsp;: {memo.MemoSubject}<br><br>" +
                          $"Requested by&emsp;&emsp;&emsp;&emsp;&ensp;&ensp;: {memo.RNameTh}<br><br>" +
                          $"Request Date&emsp;&emsp;&emsp;&emsp;&ensp;&ensp;&nbsp;: {requestDateTimeString}<br><br>" +
                          $"Last Actor by&emsp;&emsp;&emsp;&emsp;&ensp;&ensp;&nbsp;&nbsp;: {nameTh}<br><br>" +
                          $"Action by&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&ensp;&nbsp;&nbsp;:{actionProcess}<br><br>" +
                          $"Action Date&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&nbsp;:{actionDate}<br><br>" +
                          $"Status&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&ensp;&nbsp;:{memo.StatusName}<br><br>" +
                          $"Comment&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&nbsp;:{actionComment}<br><br>" +
                          $"You can action by<a href={sURLToRequest}>Click</a><br><br>" +
                          $"Best Regards;<br>" +
                          $"Wolf Approve<br>";

            foreach (var CheckListName in employees)
            {
                string NameEmail = CheckListName.Email.ToString();
                if (FixMail != string.Empty)
                {
                    NameEmail = FixMail;
                }

                SmtpClient smtpClient = new SmtpClient(smtpServer)
                {
                    Port = smtpPort,
                    Credentials = new NetworkCredential(fromEmail, fromPassword),
                    EnableSsl = true // สำหรับการเชื่อมต่อผ่าน SSL/TLS
                };

                MailMessage mailMessage = new MailMessage(fromEmail, NameEmail, subject, body);
                mailMessage.IsBodyHtml = true;

                try
                {
                    // ส่งอีเมลล์
                    smtpClient.Send(mailMessage);
                    Console.WriteLine("Email sent successfully: " + memo.MemoId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error sending email: {ex.Message}");
                    log.Info($"Error sending email: {ex.Message}");
                }
            }
        }

        private static string FixMail
        {
            get
            {
                var FixMail = ConfigurationManager.AppSettings["FixMail"];
                if (!string.IsNullOrEmpty(FixMail))
                {
                    return FixMail;
                }
                return "";
            }
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
    }
}
