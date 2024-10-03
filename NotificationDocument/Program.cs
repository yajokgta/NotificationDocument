using log4net;
using log4net.Config;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Mail;

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

        private static void Main(string[] args)
        {
            //var memoa = dbContext.TRNMemos.Where(x => dbContext.TRNUsageLogs.Any(a => a.Note01 == "5" && a.Note02 == "JOB_NOTI")).ToList();
            XmlConfigurator.Configure();
            log.Info($"=============================================================================================================");
            var currents = new List<string>()
            {
                currentDate.ToString("dd MMM yyyy", new CultureInfo("en-GB")),
                currentDate.ToString("dd/MMM/yyyy", new CultureInfo("en-GB")),
                currentDate.ToString("dd MM yyyy", new CultureInfo("en-GB")),
                currentDate.ToString("dd/MM/yyyy", new CultureInfo("en-GB")),
                currentDate.AddDays(-1).ToString("dd MMM yyyy", new CultureInfo("en-GB")),
                currentDate.AddDays(-1).ToString("dd/MMM/yyyy", new CultureInfo("en-GB")),
                currentDate.AddDays(-1).ToString("dd MM yyyy", new CultureInfo("en-GB")),
                currentDate.AddDays(-1).ToString("dd/MM/yyyy", new CultureInfo("en-GB"))
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
                        date.ToString("dd MMM yyyy", new CultureInfo("en-GB")),
                        date.ToString("dd/MMM/yyyy", new CultureInfo("en-GB")),
                        date.ToString("dd MM yyyy", new CultureInfo("en-GB")),
                        date.ToString("dd/MM/yyyy", new CultureInfo("en-GB"))
                    };

                    manuals.AddRange(addDays);
                }

                memos = dbContext.TRNMemos.Where(x => x.DocumentNo.Contains("DAR") && x.StatusName == "Completed" &&
                dbContext.TRNMemoForms.Any(a => x.MemoId == a.MemoId && a.obj_label == effectiveLabel && manuals.Contains(a.obj_value))).ToList();
                log.Info($"Date Format: {string.Join(",", manuals)}");
            }
            else
            {
                memos = dbContext.TRNMemos.Where(x => x.DocumentNo.Contains("DAR") && x.StatusName == "Completed" &&
                dbContext.TRNMemoForms.Any(a => x.MemoId == a.MemoId && a.obj_label == effectiveLabel && currents.Contains(a.obj_value)) &&
                dbContext.LogSentEmails.Any(a => x.MemoId != a.MemoId)).ToList();
                log.Info($"Date Format: {string.Join(",", currents)}");
            }

            log.Info($"Send Memo Count : {memos.Count()}");

            var viewEmployeeQuery = dbContext.ViewEmployees.Where(x => x.IsActive == true);

            foreach (var memo in memos)
            {
                var employees = new List<ViewEmployee>();
                var additionalEmployees = new List<ViewEmployee>();
                var tempCcPersonsList = new List<ViewEmployee>();

                var memoLineApproves = dbContext.TRNLineApproves.Where(y => y.MemoId == memo.MemoId).ToList();

                var buGroup = getValueAdvanceForm(memo.MAdvancveForm, "Business Group").Split(' ').ToList().FirstOrDefault();
                var department = getValueAdvanceForm(memo.MAdvancveForm, "Department");
                var documentNumber = getValueAdvanceForm(memo.MAdvancveForm, DocumentNumber);
                var promulgation = getValueAdvanceForm(memo.MAdvancveForm, "การประกาศใช้");
                var departmentInTable = dbContext.TRNMemoForms.Where(x => x.MemoId == memo.MemoId && x.obj_label == "หน่วยงานที่เกี่ยวข้อง" && x.col_label == "หน่วยงาน").Select(s => s.col_value).ToList();
                var employeeNames = dbContext.TRNMemoForms.Where(x => x.MemoId == memo.MemoId && x.obj_label == "กรณีเฉพาะบุคคลที่เกี่ยวข้อง" && x.col_label == "ชื่อผู้เกี่ยวข้อง").Select(s => s.col_value).ToList();

                log.Info("Documentnumber : " + documentNumber);
                log.Info("MemoId : " + memo.MemoId);
                log.Info("promulgation : " + promulgation);

                string AdditionalEmp = "";
                var ccPersonDistinct = new List<string>();
                var ccPersonData = "";

                void AddRequesterAndCreator()
                {
                    additionalEmployees.Add(viewEmployeeQuery.FirstOrDefault(e => e.EmployeeId == memo.RequesterId));
                    additionalEmployees.Add(viewEmployeeQuery.FirstOrDefault(e => e.EmployeeId == memo.CreatorId));
                }

                void ProcessAndAddCcPersons()
                {
                    var ccPersonNames = memo.CcPerson.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();
                    employees.AddRange(viewEmployeeQuery.Where(x => ccPersonNames.Contains(x.NameEn) || ccPersonNames.Contains(x.NameTh)).ToList());
                }

                void FinalizeEmployeeData()
                {
                    AdditionalEmp = string.Join(",", additionalEmployees.Select(nameth => nameth.NameTh));
                    ccPersonDistinct = employees
                        .Where(employee => !additionalEmployees.Select(ae => ae.NameTh).Contains(employee.NameTh))
                        .Select(employee => employee.NameTh).ToList();
                    ccPersonData = string.Join(",", ccPersonDistinct);
                }

                void AddLineApproveEmployees()
                {
                    foreach (var lineapprove in memoLineApproves)
                    {
                        var emp = viewEmployeeQuery.FirstOrDefault(v => v.EmployeeId == lineapprove.EmployeeId);
                        if (emp != null)
                        {
                            employees.Add(emp);
                        }
                    }
                }

                if (promulgation == "ทุกคนทั้งองค์กร")
                {
                    var buGroupId = dbContext.MSTDepartments.FirstOrDefault(x => x.NameEn.Contains(buGroup) || x.NameTh.Contains(buGroup))?.DepartmentId ?? 0;
                    var deptBelows = GetDepartmentBelows(buGroupId);

                    employees.AddRange(deptBelows
                                .Join(viewEmployeeQuery.ToList(), bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());
                }
                else if (promulgation == "เฉพาะหน่วยงาน")
                {
                    if (!string.IsNullOrEmpty(buGroup))
                    {
                        AddLineApproveEmployees();

                        var buGroupId = dbContext.MSTDepartments.FirstOrDefault(x => x.NameEn.Contains(buGroup) || x.NameTh.Contains(buGroup))?.DepartmentId ?? 0;
                        var deptBelows = GetDepartmentBelows(buGroupId);

                        employees.AddRange(deptBelows.Where(x => x.NameEn.Contains(department) || x.NameTh.Contains(department))
                                    .Join(viewEmployeeQuery.ToList(), bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());
                    }
                    else
                    {
                        AddLineApproveEmployees();

                        employees.AddRange(dbContext.MSTDepartments.Where(x => x.NameEn.Contains(department) || x.NameTh.Contains(department))
                            .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());
                    }
                }
                else if (promulgation == "เฉพาะบุคคล")
                {
                    if (!string.IsNullOrEmpty(buGroup))
                    {
                        AddLineApproveEmployees();

                        var buGroupId = dbContext.MSTDepartments.FirstOrDefault(x => x.NameEn.Contains(buGroup) || x.NameTh.Contains(buGroup))?.DepartmentId ?? 0;
                        var deptBelows = GetDepartmentBelows(buGroupId);

                        employees.AddRange(deptBelows.Where(x => x.NameEn.Contains(department) || x.NameTh.Contains(department))
                                   .Join(viewEmployeeQuery.ToList(), bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        employees.AddRange(viewEmployeeQuery.Where(x => employeeNames.Contains(x.NameEn) || employeeNames.Contains(x.NameTh)).ToList());
                    }
                    else
                    {
                        AddLineApproveEmployees();

                        employees.AddRange(dbContext.MSTDepartments.Where(x => x.NameEn.Contains(department) || x.NameTh.Contains(department))
                            .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        employees.AddRange(viewEmployeeQuery.Where(x => employeeNames.Contains(x.NameEn) || employeeNames.Contains(x.NameTh)).ToList());
                    }
                }
                else if (promulgation == "--select--")
                {
                    if (!string.IsNullOrEmpty(buGroup))
                    {
                        AddLineApproveEmployees();

                        var buGroupId = dbContext.MSTDepartments.FirstOrDefault(x => x.NameEn.Contains(buGroup) || x.NameTh.Contains(buGroup))?.DepartmentId ?? 0;
                        var deptBelows = GetDepartmentBelows(buGroupId);

                        employees.AddRange(deptBelows.Where(x => x.NameEn.Contains(department) || x.NameTh.Contains(department))
                                   .Join(viewEmployeeQuery.ToList(), bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        employees.AddRange(dbContext.MSTDepartments.Where(x => departmentInTable.Contains(x.NameEn) || departmentInTable.Contains(x.NameTh))
                                    .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        employees.AddRange(viewEmployeeQuery.Where(x => employeeNames.Contains(x.NameEn) || employeeNames.Contains(x.NameTh)).ToList());
                    }
                    else
                    {
                        AddLineApproveEmployees();

                        employees.AddRange(dbContext.MSTDepartments.Where(x => x.NameEn.Contains(department) || x.NameTh.Contains(department))
                           .Join(viewEmployeeQuery, bu => bu.DepartmentId, emp => emp.DepartmentId, (bu, emp) => emp).ToList());

                        employees.AddRange(viewEmployeeQuery.Where(x => employeeNames.Contains(x.NameEn) || employeeNames.Contains(x.NameTh)).ToList());
                    }
                }
                else
                {
                    log.Info("Invalid promulgation");
                }

                AddRequesterAndCreator();
                ProcessAndAddCcPersons();
                FinalizeEmployeeData();

                SendEmail(employees, memo, documentNumber, ccPersonData, AdditionalEmp);
                AddLogSendMemo(memo);
                log.Info("All Email :" + employees.Count);
                log.Info("--------------------------");
            }

            log.Info($"=============================================================================================================");
        }

        public static void AddLogSendMemo(TRNMemo memo)
        {
            LogSentEmail savetblog = new LogSentEmail();
            savetblog.MemoId = memo.MemoId;
            savetblog.ModifiedDate = memo.ModifiedDate;
            savetblog.StatusName = memo.StatusName;
            savetblog.DocumentNo = memo.DocumentNo;
            savetblog.DocumentCode = memo.DocumentCode;
            savetblog.MemoSubject = memo.MemoSubject;
            savetblog.RNameTh = memo.RNameTh;
            savetblog.RequestDate = memo.RequestDate;
            savetblog.LastActionBy = memo.LastActionBy;
            dbContext.LogSentEmails.InsertOnSubmit(savetblog);
            dbContext.SubmitChanges();
        }

        public static void SendEmail(List<ViewEmployee> employees, TRNMemo memo, String Documentnumber, string ccPersonData, string additionalEmployees)
        {
            employees.RemoveAll(x => excludeRoles.Contains(x.Email));
            employees.Distinct();

            log.Info($"Send : {string.Join("|", employees.Select(s => s.Email))}");
            var AllEmployee = dbContext.ViewEmployees.Where(x => x.IsActive == true);
            DateTime sentDatetime = DateTime.Now;
            string formattedDate = sentDatetime.ToString("dddd, MMMM dd, yyyy h:mm:ss tt");
            string requestDateTimeString = memo.RequestDate.HasValue ? memo.RequestDate.Value.ToString("M/d/yyyy h:mm:ss tt") : "-";
            string toPersons = additionalEmployees;
            var topersonNames = toPersons.Split(',').ToList();
            var ccpersonNames = ccPersonData.Split(',').ToList();

            topersonNames.RemoveAll(x => excludeRoles.Contains(x));
            ccpersonNames.RemoveAll(x => excludeRoles.Contains(x));

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

        public static List<MSTDepartment> GetDepartmentBelows(int departmentId)
        {
            var result = new List<MSTDepartment>();

            // หาข้อมูลของแผนกปัจจุบัน
            var department = dbContext.MSTDepartments.FirstOrDefault(d => d.DepartmentId == departmentId);

            if (department != null)
            {
                // เพิ่มแผนกปัจจุบันใน result
                result.Add(department);

                // ค้นหาแผนกลูก ๆ ที่มี ParentId เป็น departmentId ของปัจจุบัน
                var children = dbContext.MSTDepartments.Where(d => d.ParentId == departmentId).ToList();

                // สำหรับแผนกลูกแต่ละอัน ให้ดึงลูก ๆ ลงไปด้วย
                foreach (var child in children)
                {
                    var childDepartments = GetDepartmentBelows(child.DepartmentId);
                    result.AddRange(childDepartments);
                }
            }

            return result;
        }

        public static List<MSTDepartment> GetDepartmentAboves(int departmentId)
        {
            var result = new List<MSTDepartment>();

            // หาข้อมูลของ Department ตาม ID
            var department = dbContext.MSTDepartments.FirstOrDefault(d => d.DepartmentId == departmentId);

            if (department != null)
            {
                // เพิ่ม Department ปัจจุบันใน result
                result.Add(department);

                // ถ้ามี ParentId ให้เรียกฟังก์ชันตัวเองเพื่อดึง Parent ขึ้นไปเรื่อยๆ
                if (department.ParentId.HasValue)
                {
                    var parentDepartments = GetDepartmentAboves(department.ParentId.Value);
                    result.AddRange(parentDepartments);
                }
            }

            return result;
        }
    }
}