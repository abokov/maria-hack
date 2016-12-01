using ManagerReportConsole.ReportExecution2005;
using Microsoft.SharePoint.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Protocols;
//using File = System.IO.File;

namespace ManagerReportConsole
{
    public class AppcConfig //Системные параметры
    {
        public string FromMail;            //Откого будет приходить письмо       
        public string SmtpClient;          //Имя SMTP клиента
        public string UrlSP;               //Ссылка на сайт SharePoint
        public string FolderTempPath;      //Папка где будут времмено формироватся файлы очетов
        public string FolderTempletPath;   //Папка где хранятся шаблоны писем

        public AppcConfig()
        {
            this.FromMail = "reports@marya.ru";
            this.SmtpClient = "mail.marya.ru";
            this.UrlSP = "https://office.marya.ru/sites/tradebi-beta"; //Путь к шарику
            this.FolderTempPath = @"C:\Users\Public\Documents\";//Папка для временного хранения сформированных файлов отчетов            
            this.FolderTempletPath = @"C:\Users\Public\Documents\templet\";
        }
    }

    public class RC_ParamReports //Параметры для формирования файла
    {
        public string Name { get; set; }//Системное имя параметра
        public string Type { get; set; }//Тип параметра (Константа = 1; Выражение = 2; Справочник = 3)
        public string ValueFix { get; set; }//Фиксированная величина параметра        
        public string ValueCatalog { get; set; }//Значения параметра из справочника
        public string ValueExpression { get; set; }//Выражение для расчета параметра
        public string Format { get; set; } //Формат параметра (int, string, date, datetime)

        public RC_ParamReports(int Code, string CatalogValue)
        {
            using (ManagerReportsEntities db = new ManagerReportsEntities())
            {   //Параметры для формирования файла
                IEnumerable<L_ParamReports> pr_s = db.L_ParamReports.Where(p => p.id == Code);
                foreach (L_ParamReports pr in pr_s)
                {
                    this.Name = pr.name;
                    this.Type = pr.S_ParamSysType.sysname;
                    this.ValueFix = (CatalogValue != null && pr.S_ParamSysType.sysname == "param_directory" ? CatalogValue : pr.value_fix);
                    this.ValueExpression = pr.S_ParamSysValueExp != null ? pr.S_ParamSysValueExp.sysname : null;
                    //int? d = 0;                     d.HasValue 
                    this.Format = pr.S_ParamSysValueFormat.sysname;
                }
            }         
        }
    }

    public class RC_File
    {
        //public int ID { get; set; }                         //Уникальный Код файла в рассылке
        public int Code { get; set; }                       //Код файла из БД
        public string Name { get; set; }                    //Имя файла
        public string Format_file { get; set; }             //Расширение файла
        public string Format_report { get; set; }           //Формат файла с RS
        public string Place_save_local { get; set; }        //Место временного хранения файла локально
        public string PlaceSaveSp_name { get; set; }        //Библиотека для хранения файла на SP
        public string Path_report { get; set; }             //Путь к файлу на RS
        public int ID_subsribes { get; set; }               //Ссылка на рассылку
        public bool f_Sendmail { get; set; }                //Отправляем файл по почте или нет
        public string Email { get; set; }                   //E-mail на который будет отправлен файл
        //public int? TypeParamReport { get; set; }         //Тип параметра
        public string NameParamReport { get; set; }         //Название параметра если он из справочника        
        //public int? ParentParamReport { get; set; }
        public int? ID_Directory { get; set; }              //Значение параметра если он из справочника
        public string Name_Directory { get; set; }          //Видимое название параметра из справочника
        //public string Email_Directory { get; set; }       //E-mail из справочника параметров
        public List<RC_ParamReports> Param { get; set; }    //Параметры для формирования файла        
        public bool Success { get; set; }                   //Данный файл успешно сформирован

        private static string GetParamValue(string type, string paramvalue, string paramexpression, string paramformat)
        {//type = (Константа = 1; Выражение = 2; Справочник = 3)
            //type = (Константа = param_const; Выражение = param_expression; Справочник = param_directory)
            string value = "";

            switch (type)
            {
                case "param_const": value = paramvalue;
                    break;
                case "param_expression":
                    switch (paramexpression)
                    {
                        case "today(0)": value = (DateTime.Today).ToString("d");
                            break;
                        case "today(-1)": value = (DateTime.Today.AddDays(-1)).ToString("d");
                            break;
                        case "today(-2)": value = (DateTime.Today.AddDays(-2)).ToString("d");
                            break;
                        case "enddatemonth(-1)": value = (DateTime.Today.AddDays(-DateTime.Today.Day)).ToString("d");
                            break;
                        case "month(0)": value = (DateTime.Today.Year).ToString("0000") + (DateTime.Today.Month).ToString("00");
                            break;
                        case "month(-1)": value = (DateTime.Today.Year).ToString("0000") + (DateTime.Today.AddMonths(-1).Month).ToString("00");
                            break;
                        default: value = (DateTime.Today).ToString("d");
                            break;
                    }
                    break;
                case "param_directory": value = paramvalue; //Этот момент разрулил раньше  ???
                    break;
                default: value = "";
                    break;
            }
            return value;
        }

        protected static string GetReport(string reportpath, string format, List<RC_ParamReports> Param, string NameFile)
        {
            ReportExecutionService re2005 = new ReportExecutionService();
            re2005.Credentials = CredentialCache.DefaultCredentials;
            // Render arguments
            byte[] result = null;
            string historyID = null;
            string devInfo = @"<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";

            // Prepare report parameter
            ReportExecution2005.ParameterValue[] parameters = new ReportExecution2005.ParameterValue[Param.Count];

            for (int i = 0; i < Param.Count; i++)
            {
                parameters[i] = new ReportExecution2005.ParameterValue();
                parameters[i].Name = Param[i].Name;
                parameters[i].Value = GetParamValue(Param[i].Type, Param[i].ValueFix, Param[i].ValueExpression, Param[i].Format);

                Console.WriteLine("Param[i].Name = {0}; Param[i].ValueExpression = {1}; Param[i].ValueCatalog = {2}; Param[i].ValueFix = {3}", Param[i].Name, Param[i].ValueExpression, Param[i].ValueCatalog, Param[i].ValueFix);

                Console.WriteLine("parameters[i].Name = {0}; parameters[i].Value = {1}", parameters[i].Name, parameters[i].Value);
            }
            //Блок херни
            ReportExecution2005.DataSourceCredentials[] credentials = null;
            string showHideToggle = null;
            string encoding;
            string mimeType;
            string extension = "";
            ReportExecution2005.Warning[] warnings = null;
            ReportExecution2005.ParameterValue[] reportHistoryParameters = null;
            string[] streamIDs = null;
            //            
            ExecutionInfo execInfo = new ExecutionInfo();
            ExecutionHeader execHeader = new ExecutionHeader();
            try
            {//
                re2005.ExecutionHeaderValue = execHeader;
                execInfo = re2005.LoadReport(reportpath, historyID);

                re2005.SetExecutionParameters(parameters, "ru-RU");
                String SessionId = re2005.ExecutionHeaderValue.ExecutionID;
                ///Console.WriteLine("SessionID: {0}", re2005.ExecutionHeaderValue.ExecutionID);

                result = re2005.Render(format, devInfo, out extension, out encoding, out mimeType, out warnings, out streamIDs);
                ///Console.WriteLine("format = {0}; devInfo = {1}; extension = {2}; encoding = {3}; mimeType = {4}; warnings = {5}; streamIDs = {6}", format, devInfo, extension, encoding, mimeType, warnings, streamIDs);
            }
            catch (SoapException e)
            {
                Console.WriteLine(e.Detail.OuterXml);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            ///finally            {                Console.WriteLine(e.Message);            }
            // Write the contents of the report to an MHTML file.
            try
            {
                using (FileStream stream = System.IO.File.Create(NameFile, result.Length))
                {
                    ///Console.WriteLine("File created.");
                    stream.Write(result, 0, result.Length);
                    Console.WriteLine("Result written to the file {0}.", NameFile);
                    //                    stream.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("ОШИБКА " + e.Message);
            }
            return NameFile;
        }

        public void CreateFileReport()//Создание файлов из рассылки на Sp
        {
            
            //Получаем системные параметры
            AppcConfig opt = new AppcConfig();
            string st_ex = this.Name_Directory != null ? " (" + this.Name_Directory + ")" : "";
            string curFile = GetReport(this.Path_report, this.Format_report, this.Param, opt.FolderTempPath + this.Name +st_ex+ "."+ this.Format_file);
            //Если такой файл создался, т.е. существует локально то

            this.Success = System.IO.File.Exists(curFile);

        }

        /*
        Примеры
        http://stackoverflow.com/questions/17057074/how-to-download-upload-files-from-to-sharepoint-2013-using-csom
        http://msdn.microsoft.com/en-us/library/office/ee956524(v=office.14).aspx
        http://blogs.msdn.com/b/sridhara/archive/2010/03/12/uploading-files-using-client-object-model-in-sharepoint-2010.aspx          
        */
        public void UploadFileSp()
        {
            AppcConfig opt = new AppcConfig();
            string url = opt.UrlSP;

            //string fullNameFile = this.Place_save_local + this.Name + "." + this.Format_file;
            string st_ex = this.Name_Directory != null ? " (" + this.Name_Directory + ")" : ""; //Задублированная логика А если полное наименование засунуть в объект
            string fullNameFile = this.Place_save_local + this.Name + st_ex + "." + this.Format_file;
            //opt.FolderTempPath + this.Name +st_ex+ "."+ this.Format_file
            using (var clientContext = new ClientContext(url))
            {
                clientContext.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //Вываливается если файл не был создан
                using (var fs = new FileStream(fullNameFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var fi = new FileInfo(fullNameFile);
                    var list = clientContext.Web.Lists.GetByTitle(this.PlaceSaveSp_name);
                    clientContext.Load(list.RootFolder);
                    clientContext.ExecuteQuery();
                    var fileUrl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, fi.Name);
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(clientContext, fileUrl, fs, true);
                    Console.WriteLine("Кладём файл {0} в библиотеку {1} на портал {2}", this.Name, this.PlaceSaveSp_name, opt.UrlSP);
                }
            }
        }
  
    }

    class RC_Mail
    {
        public string Code { get; set; }//уникальный код письма если Стат то Code связи Рассылка-Получатели, а если динамически
        public string Subject { get; set; }//Тема письма
        public int CodeSubscribes { get; set; }//Код рассылки
        public string From { get; set; }//E-mail отправителя
        public string ToList { get; set; }//Получатели письма
        public string Body { get; set; }//Тело письма     
        public List<RC_File> FilesMail { get; set; }//Файлы письма
        public bool Success { get; set; }//Данное письмо успешно отправлено                
    }

    class RC_Subsribes
    {
        public int Code { get; set; }                       //Код рассылки
        public string Name { get; set; }                    //Имя рассылки
        public List<RC_File> FilesAttach { get; set; }        //Файлы письма
        public List<RC_Mail> Mails { get; set; }            //Письма входящие в рассылку        
        public bool Success { get; set; }                   //Данная рассылка успешно

        public RC_Subsribes(int Code)
        {
            using (ManagerReportsEntities db = new ManagerReportsEntities())
            {
                IEnumerable<L_Subscribes> sb_s = db.L_Subscribes.Where(p => p.id == Code);
                foreach (L_Subscribes sb in sb_s)
                {
                    this.Code = sb.id;
                    this.Name = sb.name;
                }
                //Формируем список всех файлов
                IEnumerable<ViewFileReport_Result> fr_s = db.ViewFileReport(Code);
                FilesAttach = new List<RC_File>();
                foreach (ViewFileReport_Result fr in fr_s)
                {   
                    string[] words = fr.email.Split(';');            
                    foreach (string s in words) 
                    { //Каждый файл на каждый email
                        //Формируем список файлов
                        RC_File FilesRep = new RC_File();
                        FilesRep.Code = fr.id;
                        FilesRep.Name = fr.name;
                        FilesRep.Format_file = fr.Format_name;
                        FilesRep.Format_report = fr.Format_sysname;
                        FilesRep.Place_save_local = fr.place_save_local;
                        FilesRep.PlaceSaveSp_name = fr.PlaceSaveSp_name;
                        FilesRep.Path_report = fr.path_report;
                        FilesRep.ID_subsribes = fr.id_subsribes;
                        FilesRep.f_Sendmail = fr.sendmail;
                        FilesRep.Email = s;//fr.email;
                        FilesRep.NameParamReport = fr.NameParamReport;
                        FilesRep.ID_Directory = fr.ID_Directory;
                        FilesRep.Name_Directory = fr.Name_Directory;
                        //Для каждого файла список параметров если они не из справочника
                        IEnumerable<L_ParamReports> pr_s = db.L_ParamReports.Where(p => p.id_filereport == FilesRep.Code);
                        //Список параметров по всем параметрам для этого файла
                        FilesRep.Param = new List<RC_ParamReports>();
                        foreach (L_ParamReports pr in pr_s)
                        {
                            FilesRep.Param.Add(new RC_ParamReports(pr.id, FilesRep.ID_Directory.ToString()));
                        }
                        if (FilesRep.NameParamReport != null)
                        {}
                       ///Console.WriteLine("------------------------------------");
                       ///Console.WriteLine("{0}; {1}; {2}; {3}", fr.name, fr.Name_Directory, fr.email, s);                                    
                        FilesAttach.Add(FilesRep);
                    }                 
                }                
            }
            

            //Если файлы не сформировались ставим Success = False

            //Формируем список всех писем

            //Если письма ,не сформировались ставим Success = False
        }

        public void CreateFileReport()//Создание файлов из рассылки на Sp
        {
            //Получаем системные параметры
            AppcConfig opt = new AppcConfig();
            this.Success = true;
            for (int j = 0; j < FilesAttach.Count; j++)
            {
                Console.WriteLine("Пытаемся создать " + FilesAttach[j].Name);
                RC_File p = FilesAttach[j];
                p.CreateFileReport();
                Console.WriteLine((p.Success ? "Успешное " : " неуспешное ") + "создание файла {0}", FilesAttach[j].Name);
                this.Success = this.Success && p.Success;                
            }
            Console.WriteLine((this.Success ? "Успешное " : " неуспешное ") + "формирование файлов рассылки {0}", this.Name);
        }

        public void UploadFileSp()
        {
            //Получаем системные параметры
            AppcConfig opt = new AppcConfig();
            this.Success = true;
            for (int j = 0; j < FilesAttach.Count; j++)
            {
                Console.WriteLine("Пытаемся создать " + FilesAttach[j].Name);
                RC_File p = FilesAttach[j];
                p.UploadFileSp();
                Console.WriteLine((p.Success ? "Успешное " : " неуспешное ") + "создание файла SP {0}", FilesAttach[j].Name);
                this.Success = this.Success && p.Success;
            }
            Console.WriteLine((this.Success ? "Успешное " : " неуспешное ") + "формирование файлов SP рассылки {0}", this.Name);
        }

        public void CreateEmail()
        {
            Console.WriteLine("-----------------------------------------------------------------------");
            /*
            for (int i = 0; i < FilesAttach.Count; i++)
                Console.WriteLine("Файл {0}; адрес {1}", FilesAttach[i].Name, FilesAttach[i].Email);
            */
            string str = "";
            string fl = "";
            foreach (var f in FilesAttach.OrderBy(f=> f.Email)) //Списки проходим так отказываемся от for
            {
                if (f.f_Sendmail)
                {
                    if (str != f.Email)
                    {                    
                        str= f.Email;
                        Console.WriteLine("Файл {0}; адрес {1}", fl, f.Email);
                        fl = f.Name;
                    }
                    else 
                    {
                        fl = fl + f.Name;
                    }
                }
            }


        }

        public void SendEmail()
        { 
        
        }

    }


    class ManagerReport
    {
        static void Main(string[] args)
        {
            int code = 1;
            //Создаем рассылку
            RC_Subsribes sb = new RC_Subsribes(code);
            //Создаем файлы рассылки
            ////sb.CreateFileReport();
            //Кладем файлы рассылки на портал
            ////sb.UploadFileSp();
            //Рассылаем файлы рассылки с портала
            sb.CreateEmail();
            sb.SendEmail();

            Console.ReadKey();
        }
    }
}
