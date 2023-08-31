using System;
using System.Data;
using System.ServiceProcess;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Threading;

namespace ChemwellFileRunner_Service
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }


        #region settings

        public static string user = "PSMExchangeUser"; // логин для базы обмена файлами и для базы CGM Analytix
        public static string password = "PSM_123456";  // пароль для базы обмена файлами и для базы CGM Analytix

        public static bool ServiceIsActive;            // флаг для запуска и остановки потока

        static object FileExportLogLocker = new object();    // локер для логов обмена
        static object ServiceLogLocker = new object();       // локер для логов драйвера

        //public static string AnalyzerResultPath = AppDomain.CurrentDomain.BaseDirectory + "\\ChemwellResults"; // папка для файлов с результатами

        #endregion

        #region фунции логов

        // Лог драйвера
        static void ServiceLog(string Message)
        {
            lock (ServiceLogLocker)
            {
                try
                {
                    string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\Service";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    string filename = path + "\\ServiceThread_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
                    if (!System.IO.File.Exists(filename))
                    {
                        using (StreamWriter sw = System.IO.File.CreateText(filename))
                        {
                            sw.WriteLine(DateTime.Now + ": " + Message);
                        }
                    }
                    else
                    {
                        using (StreamWriter sw = System.IO.File.AppendText(filename))
                        {
                            sw.WriteLine(DateTime.Now + ": " + Message);
                        }
                    }
                }
                catch
                {

                }
            }
        }

        // лог записи файла в FNC2
        static void FileExportLog(string Message)
        {
            lock (FileExportLogLocker)
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\FileExport";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string filename = path + "\\ExchangeThread_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
                if (!System.IO.File.Exists(filename))
                {
                    using (StreamWriter sw = System.IO.File.CreateText(filename))
                    {
                        sw.WriteLine(DateTime.Now + ": " + Message);
                    }
                }
                else
                {
                    using (StreamWriter sw = System.IO.File.AppendText(filename))
                    {
                        sw.WriteLine(DateTime.Now + ": " + Message);
                    }
                }
            }
        }

        // проверка старых логов
        public void CheckOldLog()
        {
            while (ServiceIsActive)
            {
                try
                {
                    ServiceLog("Проверка логов");
                    //сегодня минус 28 дней
                    //все что старше удаляем
                    DateTime now = DateTime.Now.AddDays(-28);
                    string LogsDirectory = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\Service";

                    if (Directory.Exists(LogsDirectory))
                    {
                        string[] files = Directory.GetFiles(LogsDirectory);
                        foreach (string s in files)
                        {
                            DateTime FileDate = File.GetCreationTime(s);
                            int comparison = FileDate.CompareTo(now);
                            if (comparison == -1)
                            {
                                File.Delete(s);
                            }

                        }
                    }

                }
                catch (Exception LogError)
                {
                    ServiceLog("Ошибка: " + LogError);
                }
                Thread.Sleep(4320000);
            }

        }

        #endregion

        #region вспомогательные функции
        //дописываем к номеру месяца ноль если нужно
        public static string CheckZero(int CheckPar)
        {
            string BackPar = "";
            if (CheckPar < 10)
            {
                BackPar = $"0{CheckPar}";
            }
            else
            {
                BackPar = $"{CheckPar}";
            }
            return BackPar;
        }
        #endregion

        #region функция отправки файлов в базу FNC2, таблица FileExchangeIn
        public static void WorkingWithOutFiles(string file)
        {
            string ServiceId = ConfigurationManager.AppSettings["ServiceCode"];
            string OutFolder = ConfigurationManager.AppSettings["FolderOut"];
            // папка для ошибок
            string ErrorPath = OutFolder + @"\errors";
            if (!Directory.Exists(ErrorPath))
            {
                Directory.CreateDirectory(ErrorPath);
            }

            // обрезаем только имя текущего файла
            string FileName = file.Substring(OutFolder.Length + 1);

            // Меняем название файла
            //DateTime now = DateTime.Now;
            //string FileName = "Results_" + now.Year + CheckZero(now.Month) + CheckZero(now.Day) + CheckZero(now.Hour) + CheckZero(now.Minute) + CheckZero(now.Second) + CheckZero(now.Millisecond) + ".res";

            string CGMConnectionString = ConfigurationManager.ConnectionStrings["SQLConnection"].ConnectionString;
            CGMConnectionString = String.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");

            try
            {
                using (SqlConnection CGMconnection = new SqlConnection(CGMConnectionString))
                {
                    CGMconnection.Open();

                    //SqlCommand SqlInsertCommand = new SqlCommand();
                    SqlCommand SqlInsertCommand = CGMconnection.CreateCommand();
                    SqlInsertCommand.CommandText = "INSERT INTO FileExchangeIn VALUES (@filename, @file, @changeuser, @changedate, @serviceid, 0 )";

                    // создаем параметры для инсерта
                    SqlParameter filenameParam = new SqlParameter("@filename", SqlDbType.NVarChar, 200);
                    SqlParameter fileParam = new SqlParameter("@file", SqlDbType.Image, 1000000);
                    SqlParameter changeuserParam = new SqlParameter("@changeuser", SqlDbType.NVarChar, 50);
                    SqlParameter changedateParam = new SqlParameter("@changedate", SqlDbType.NVarChar, 50);
                    SqlParameter serviceidParam = new SqlParameter("@serviceid", SqlDbType.NVarChar, 50);

                    // добавляем параметры к команде
                    SqlInsertCommand.Parameters.Add(filenameParam);
                    SqlInsertCommand.Parameters.Add(fileParam);
                    SqlInsertCommand.Parameters.Add(changeuserParam);
                    SqlInsertCommand.Parameters.Add(changedateParam);
                    SqlInsertCommand.Parameters.Add(serviceidParam);

                    // массив для хранения бинарных данных файла
                    byte[] imageData;
                    using (System.IO.FileStream fs = new System.IO.FileStream(file, FileMode.Open))
                    {
                        imageData = new byte[fs.Length];
                        fs.Read(imageData, 0, imageData.Length);
                    }

                    string NowString = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " " + DateTime.Now.ToShortTimeString() + ":" + DateTime.Now.Second + "." + DateTime.Now.Millisecond;

                    //string FileName = file.Substring(OutFolder.Length + 1);

                    // передаем данные в команду через параметры
                    SqlInsertCommand.Parameters["@filename"].Value = FileName;
                    SqlInsertCommand.Parameters["@file"].Value = imageData;
                    SqlInsertCommand.Parameters["@changeuser"].Value = ServiceId;
                    SqlInsertCommand.Parameters["@changedate"].Value = NowString;
                    SqlInsertCommand.Parameters["@serviceid"].Value = ServiceId;

                    SqlInsertCommand.ExecuteNonQuery();
                }

                File.Delete(file);
                FileExportLog($"Файл {file} обработан и записан в базу FNC2.");
                FileExportLog("");
            }
            catch (Exception ex)
            {
                FileExportLog(ex.Message);
                // помещение файла в папку с ошибками
                if (File.Exists(ErrorPath + @"\" + FileName))
                {
                    File.Delete(ErrorPath + @"\" + FileName);
                }
                File.Move(file, ErrorPath + @"\" + FileName);

                FileExportLog($"Ошибка обработки файла {file}. Файл перемещен в папку Error");
                FileExportLog("");
            }
        }
        #endregion

        // поток мониторинга и отправки файлов с результатами в базу FNC2, таблица FileExchangeIn
        public static void GetOut()
        {
            int ServerCount = 0; // счетчик

            while (ServiceIsActive)
            {
                //int ServerCount = 0; // счетчик

                ServerCount++;
                if (ServerCount == 100)
                {
                    ServerCount = 0;
                    ServiceLog("waiting...");
                }

                try
                {
                    string OutFolder = ConfigurationManager.AppSettings["FolderOut"];

                    if (Directory.Exists(OutFolder))
                    {
                        string[] FileFormat = ConfigurationManager.AppSettings["FileType"].Split(';');
                        foreach (string format in FileFormat)
                        {
                            string[] Files = Directory.GetFiles(OutFolder, format);
                            foreach (string file in Files)
                            {
                                //FileExportLog($"Файл: {file}");
                                ServiceLog($"Файл: {file}");
                                WorkingWithOutFiles(file);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    FileExportLog(ex.Message);
                }

                Thread.Sleep(1000);
            }
        }

        protected override void OnStart(string[] args)
        {
            ServiceIsActive = true;

            //поток для отправки файлов в базу
            Thread GetOutThread = new Thread(new ThreadStart(GetOut));
            GetOutThread.Name = "GetOutFiles Thread";
            GetOutThread.Start();
            ServiceLog("Сервис начал работу.");
        }

        protected override void OnStop()
        {
            ServiceIsActive = false;
            ServiceLog("Сервис остановлен");
        }
    }
}
