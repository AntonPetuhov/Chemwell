using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.Net;
using System.Threading;
using System.Configuration;
using System.Data.SqlClient;
using System.Globalization;
using System.ServiceProcess;

namespace ChemwellService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        #region settings

        public static string AnalyzerCode = "907";                   // код из аналайзер конфигурейшн, который связывает прибор в PSMV2
        public static string AnalyzerConfigurationCode = "CHEMWELL"; // код прибора из аналайзер конфигурейшн

        public static string user = "PSMExchangeUser"; // логин для базы обмена файлами и для базы CGM Analytix
        public static string password = "PSM_123456";  // пароль для базы обмена файлами и для базы CGM Analytix 

        public static bool ServiceIsActive;                            // флаг для запуска и остановки потока
        public static List<Thread> ListOfThreads = new List<Thread>(); // список работающих потоков

        public static string AnalyzerResultPath = AppDomain.CurrentDomain.BaseDirectory + "AnalyzerResults"; // папка для файлов с результатами
        // буфферная папка для файлов с результатами, файлы из базы сначала будут помещаться в эту папку
        public static string bufResultPath = AnalyzerResultPath + "\\buf";

        static object ExchangeLogLocker = new object();    // локер для логов обмена
        static object FileResultLogLocker = new object();  // локер для логов обмена
        static object BufFileResultLogLocker = new object();  // локер для логов обмена
        static object ServiceLogLocker = new object();     // локер для логов драйвера

        #endregion

        #region функции логов

        // Лог драйвера
        static void ServiceLog(string Message)
        {
            lock (ServiceLogLocker)
            {
                try
                {
                    string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\Service";
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

        // Лог обработки файлов результатов для CGM
        static void FileResultLog(string Message)
        {
            try
            {
                lock (FileResultLogLocker)
                {
                    //string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\FileResult" + "\\" + DateTime.Now.Year + "\\" + DateTime.Now.Month;
                    string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\FileResult";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    //string filename = path + $"\\{FileName}" + ".txt";
                    string filename = path + $"\\ResultLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";

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
            catch
            {

            }
        }

        // Лог обработки файлов прибора .lis в буфферной папке buf
        static void bufResultLog(string Message)
        {
            try
            {
                lock (BufFileResultLogLocker)
                {
                    //string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\FileResult" + "\\" + DateTime.Now.Year + "\\" + DateTime.Now.Month;
                    string path = AppDomain.CurrentDomain.BaseDirectory + "\\Log\\BufferFileResult";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    //string filename = path + $"\\{FileName}" + ".txt";
                    string filename = path + $"\\BufLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";

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
            catch
            {

            }
        }

        #endregion

        #region вспомогательные функции

        //апдейтим статус
        public static void SQLUpdateStatus(int IDParam, SqlConnection DBconnection)
        {
            SqlCommand UpdateStatusCommand = new SqlCommand(@"UPDATE FileExchangeIn  SET [Status]=1 where id = @id", DBconnection);
            UpdateStatusCommand.Parameters.AddWithValue("@id", IDParam);
            UpdateStatusCommand.ExecuteNonQuery();
            ServiceLog("Статус записи изменен");

        }

        //удаляем из очереди
        public static void SQLDelete(int IDParam, SqlConnection DBconnection)
        {
            SqlCommand DeleteCommand = new SqlCommand(@"Delete FROM FileExchangeIn where id = @id", DBconnection);
            DeleteCommand.Parameters.AddWithValue("@id", IDParam);
            DeleteCommand.ExecuteNonQuery();
            ServiceLog("Запись удалена из таблицы FileExchangeIn");
            ServiceLog("");
        }

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

        // Функция преобразования кода теста прибора в код теста PSMV2 в CGM
        public static string TranslateToPSMCodes(string AnalyzerTestCodesPar)
        {
            string BackTestCode = "";
            try
            {
                string CGMConnectionString = ConfigurationManager.ConnectionStrings["CGMConnection"].ConnectionString;
                //string CGMConnectionString = @"Data Source=CGM-APP11\SQLCGMAPP11;Initial Catalog=KDLPROD; Integrated Security=True;";
                CGMConnectionString = String.Concat(CGMConnectionString, $"User Id = {user}; Password = {password}");
                using (SqlConnection Connection = new SqlConnection(CGMConnectionString))
                {
                    Connection.Open();
                    // Ищем только тесты, которые настроены для прибора exias и настроены для PSMV2
                    SqlCommand TestCodeCommand = new SqlCommand(
                       "SELECT k1.amt_analyskod  FROM konvana k " +
                       "LEFT JOIN konvana k1 ON k1.met_kod = k.met_kod AND k1.ins_maskin = 'PSMV2' " +
                       $"WHERE k.ins_maskin = '{AnalyzerConfigurationCode}' AND k.amt_analyskod = '{AnalyzerTestCodesPar}' ", Connection);
                    SqlDataReader Reader = TestCodeCommand.ExecuteReader();

                    if (Reader.HasRows) // если есть данные
                    {
                        while (Reader.Read())
                        {
                            if (!Reader.IsDBNull(0)) { BackTestCode = Reader.GetString(0); };
                        }
                    }
                    Reader.Close();
                    Connection.Close();
                }
            }
            catch (Exception error)
            {
                FileResultLog($"Error: {error}");
            }

            return BackTestCode;
        }

        // функция, которая формирует результирующий файл и отправляет в папку для обработки FileGetterService
        public static void SendFileToCGM(string message, string file)
        {
            #region папки архива, результатов и ошибок

            string OutFolder = ConfigurationManager.AppSettings["FolderOut"];
            // архивная папка
            string ArchivePath = AnalyzerResultPath + @"\Archive";
            // папка для ошибок
            string ErrorPath = AnalyzerResultPath + @"\Error";
            // папка для файлов с результатами для CGM
            string CGMPath = AnalyzerResultPath + @"\CGM";

            if (!Directory.Exists(ArchivePath))
            {
                Directory.CreateDirectory(ArchivePath);
            }

            if (!Directory.Exists(ErrorPath))
            {
                Directory.CreateDirectory(ErrorPath);
            }

            if (!Directory.Exists(CGMPath))
            {
                Directory.CreateDirectory(CGMPath);
            }
            #endregion

            DateTime now = DateTime.Now;
            string filename = "Results_" + now.Year + CheckZero(now.Month) + CheckZero(now.Day) + CheckZero(now.Hour) + CheckZero(now.Minute) + CheckZero(now.Second) + CheckZero(now.Millisecond) + ".res";
            // название файла .ок, который должен создаваться вместе с результирующим для обработки службой FileGetterService
            string OkFileName = "";
            // получаем название файла .ок на основании файла с результатом
            if (filename.IndexOf(".") != -1)
            {
                OkFileName = filename.Split('.')[0] + ".ok";
            }

            try
            {
                // собираем полное сообщение с результатом
                //AllMessage = MessageHead + "\r" + MessageTest;
                //FileResultLog(AllMessage);


                // создаем файл для записи результата в папке для рез-тов
                if (!File.Exists(CGMPath + @"\" + filename))
                //if (!File.Exists(OutFolder + @"\" + filename))
                {
                    using (StreamWriter sw = File.CreateText(CGMPath + @"\" + filename))
                    //using (StreamWriter sw = File.CreateText(OutFolder + @"\" + filename))
                    {
                        foreach (string msg in message.Split('\r'))
                        {
                            sw.WriteLine(msg);
                        }
                    }
                }
                else
                {
                    File.Delete(CGMPath + @"\" + filename);
                    using (StreamWriter sw = File.CreateText(CGMPath + @"\" + filename))
                    //File.Delete(OutFolder + @"\" + filename);
                    //using (StreamWriter sw = File.CreateText(OutFolder + @"\" + filename))
                    {
                        foreach (string msg in message.Split('\r'))
                        {
                            sw.WriteLine(msg);
                        }
                    }
                }

                // создаем .ok файл в папке для рез-тов
                if (OkFileName != "")
                {
                    if (!File.Exists(CGMPath + @"\" + OkFileName))
                    //if (!File.Exists(OutFolder + @"\" + OkFileName))
                    {
                        using (StreamWriter sw = File.CreateText(CGMPath + @"\" + OkFileName))
                        //using (StreamWriter sw = File.CreateText(OutFolder + @"\" + OkFileName))
                        {
                            sw.WriteLine("ok");
                        }
                    }
                    else
                    {
                        File.Delete(CGMPath + OkFileName);
                        using (StreamWriter sw = File.CreateText(CGMPath + @"\" + OkFileName))
                        //File.Delete(OutFolder + OkFileName);
                        //using (StreamWriter sw = File.CreateText(OutFolder + @"\" + OkFileName))
                        {
                            sw.WriteLine("ok");
                        }
                    }
                }

                /*
                // помещение файла в архивную папку
                if (File.Exists(ArchivePath + @"\" + filename))
                {
                    File.Delete(ArchivePath + @"\" + filename);
                }

                File.Move(file, ArchivePath + @"\" + filename);

                FileResultLog("Файл обработан и перемещен в папку Archive");
                FileResultLog("");
                */
            }
            catch (Exception e)
            {
                /*
                FileResultLog(e.ToString());
                // помещение файла в папку с ошибками
                if (File.Exists(ErrorPath + @"\" + filename))
                {
                    File.Delete(ErrorPath + @"\" + filename);
                }
                File.Move(file, ErrorPath + @"\" + filename);

                FileResultLog("Ошибка обработки файла. Файл перемещен в папку Error");
                FileResultLog("");
                */
            }
        }

        // Создание отдельного файла с результатом для RID в папке AnalyzerResults
        public static void CreateResultFile(string message)
        {
            DateTime now = DateTime.Now;
            string filename = "Results_" + now.Year + CheckZero(now.Month) + CheckZero(now.Day) + CheckZero(now.Hour) + CheckZero(now.Minute) + CheckZero(now.Second) + CheckZero(now.Millisecond) + ".res";

            try
            {
                // создаем файл для записи результата в папке для рез-тов
                if (!File.Exists(AnalyzerResultPath + @"\" + filename))
                {
                    using (StreamWriter sw = File.CreateText(AnalyzerResultPath + @"\" + filename))
                    {
                        foreach (string msg in message.Split('\r'))
                        {
                            sw.WriteLine(msg);
                        }
                    }
                }
                else
                {
                    File.Delete(AnalyzerResultPath + @"\" + filename);
                    using (StreamWriter sw = File.CreateText(AnalyzerResultPath + @"\" + filename))
                    {
                        foreach (string msg in message.Split('\r'))
                        {
                            sw.WriteLine(msg);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FileResultLog($"Exception: {ex.Message}");
            }

        }

        // помещение исходного файла от прибора в архивную папку после обработки
        public static void OriginalFileToArchive(string file, string filename)
        {
            string bufArchivePath = bufResultPath + @"\archive";

            if (!Directory.Exists(bufArchivePath))
            {
                Directory.CreateDirectory(bufArchivePath);
            }

            // помещение файла в папку с ошибками
            if (File.Exists(bufArchivePath + @"\" + filename))
            {
                File.Delete(bufArchivePath + @"\" + filename);
            }
            File.Move(file, bufArchivePath + @"\" + filename);

        }

        // менеджер потоков
        public static void CheckThreads()
        {
            while (ServiceIsActive)
            {
                List<Thread> ListOfThreadsSearch = new List<Thread>();
                foreach (Thread th in ListOfThreads)
                {
                    ListOfThreadsSearch.Add(th);
                }

                foreach (Thread th in ListOfThreadsSearch)
                {
                    if (!th.IsAlive)
                    {
                        ServiceLog($"Поток {th.Name} не работает. IsAlive: {th.IsAlive}, Состояние потока: {th.ThreadState}");

                        try
                        {
                            if (th.Name == "Result Getter")
                            {
                                ListOfThreads.Remove(th);
                                Thread NewThread = new Thread(ResultGetterFunction);
                                NewThread.Name = th.Name;
                                ListOfThreads.Add(NewThread);
                                NewThread.Start();
                            }
                            if (th.Name == "Orinial analyzer files handler")
                            {
                                ListOfThreads.Remove(th);
                                Thread NewThread = new Thread(bufFilesProcessing);
                                NewThread.Name = th.Name;
                                ListOfThreads.Add(NewThread);
                                NewThread.Start();
                            }
                            if (th.Name == "ResultsProcessing")
                            {
                                ListOfThreads.Remove(th);
                                Thread NewThread = new Thread(ResultsProcessing);
                                NewThread.Name = th.Name;
                                ListOfThreads.Add(NewThread);
                                NewThread.Start();
                            }
                        }
                        catch (Exception ex)
                        {
                            ServiceLog($"Не получилось запустить поток {th.Name}: {ex}");
                        }

                    }
                    else
                    {
                        ServiceLog($"Поток {th.Name} работает. IsAlive: {th.IsAlive}, Состояние потока: {th.ThreadState}");
                    }
                }
                ListOfThreadsSearch.Clear();
                Thread.Sleep(360000);
            }
        }

        #endregion

        #region функция, которая сохраняет в папку файлы с результатами из базы FNC2
        public static void ResultGetterFunction()
        {
            while (ServiceIsActive)
            {
                try
                {
                    ServiceLog("Ожидание файлов с результатами от Chemwell.");

                    // если нет папки для результатов, создаем её
                    if (!Directory.Exists(AnalyzerResultPath))
                    {
                        Directory.CreateDirectory(AnalyzerResultPath);
                    }

                    // если нет буферной папки, создаём
                    if (!Directory.Exists(bufResultPath))
                    {
                        Directory.CreateDirectory(bufResultPath);
                    }


                    string ExchangeDBConnectionString = ConfigurationManager.ConnectionStrings["ExchangeDBConnection"].ConnectionString;
                    ExchangeDBConnectionString = String.Concat(ExchangeDBConnectionString, $"User Id = {user}; Password = {password}");

                    try
                    {
                        using (SqlConnection FNC2connection = new SqlConnection(ExchangeDBConnectionString))
                        {
                            FNC2connection.Open();
                            // запрос получения файлов от Chemwell
                            // SqlCommand GetResultFiles = new SqlCommand($"SELECT TOP 100 fi.id, fi.FileName, [File] FROM FileExchangeIn fi " +
                            //                                            $"WHERE fi.ServiceID like '{AnalyzerConfigurationCode}' and fi.Status = 0 ORDER BY ChangeDate", FNC2connection);

                            SqlCommand GetResultFiles = new SqlCommand($"SELECT TOP 1 fi.id, fi.FileName, [File] FROM FileExchangeIn fi " +
                                                                       $"WHERE fi.ServiceID like '{AnalyzerConfigurationCode}' and fi.Status = 0 ORDER BY ChangeDate", FNC2connection);
                            SqlDataReader Reader = GetResultFiles.ExecuteReader();

                            int id = 0;
                            string FileName = "";
                            byte[] resData = { };
                            //int ind = 0;

                            if (Reader.HasRows)
                            {
                                while (Reader.Read())
                                {
                                    if (!Reader.IsDBNull(0)) { id = Reader.GetInt32(0); }
                                    if (!Reader.IsDBNull(1)) { FileName = Reader.GetString(1); };
                                    if (!Reader.IsDBNull(2)) { resData = (byte[])Reader.GetValue(2); }

                                    string DirFileName = "";

                                    DateTime now = DateTime.Now;
                                    //FileName = FileName + now.Year + CheckZero(now.Month) + CheckZero(now.Day) + CheckZero(now.Hour) + CheckZero(now.Minute) + CheckZero(now.Second) + CheckZero(now.Millisecond);
                                    FileName = CheckZero(now.Hour) + CheckZero(now.Minute) + CheckZero(now.Second) + CheckZero(now.Millisecond) + "_" + FileName;
                                    //DirFileName = $"{AnalyzerResultPath}\\" + FileName;
                                    DirFileName = $"{bufResultPath}\\" + FileName;
                                    using (System.IO.FileStream fs = new System.IO.FileStream(DirFileName, FileMode.OpenOrCreate))
                                    {
                                        fs.Write(resData, 0, resData.Length);
                                    }

                                    ServiceLog($"Файл с результатами: {FileName}");
                                    FileResultLog($"");

                                }

                                Reader.Close();
                                SQLUpdateStatus(id, FNC2connection);
                                SQLDelete(id, FNC2connection);

                            }

                            FNC2connection.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        ServiceLog(ex.Message);
                        ServiceLog(ex.ToString());
                    }

                }
                catch (Exception ex)
                {
                    ServiceLog(ex.Message);
                }

                Thread.Sleep(30000);
            }
        }
        #endregion

        #region функция, котрая обрабатывает файлы в буферной папке и создает отдельные файлы с результатами для каждой заявки
        public static void bufFilesProcessing()
        {
            while (ServiceIsActive)
            {
                // получаем список всех файлов в буферной папке
                string[] Files = Directory.GetFiles(bufResultPath, "*.lis");

                // пробегаем по файлам
                foreach (string file in Files)
                {
                    bufResultLog("Обработка файла прибора: " + "\r" + file);

                    string RequestMessage = "";

                    bool firstRID = true; // Шк первый в результирующем файле
                    // обрезаем только имя текущего файла
                    string FileName = file.Substring(bufResultPath.Length + 1);

                    //string[] lines = System.IO.File.ReadAllLines(file);
                    Encoding encoding = Encoding.GetEncoding("windows-1251");
                    string[] lines = System.IO.File.ReadAllLines(file, encoding);


                    // проходим по строкам в файле
                    foreach (string line in lines)
                    {
                        if (line.StartsWith("P|"))
                        {
                            // и если это первый ШК в результирующем файле
                            if (firstRID)
                            {
                                // следующий ШК будет уже не первым в файле
                                firstRID = false;
                            }
                            else
                            {
                                // при обнаружении следующего номера шк, нужно сформировать файл с результатами для предыдушего
                                bufResultLog(RequestMessage);
                                bufResultLog("Формируем файл с результатом и переходим к следующем ШК в результирующем файле");
                                bufResultLog("");
                                CreateResultFile(RequestMessage);
                                Thread.Sleep(100);
                                RequestMessage = "";
                            }
                        }

                        if (line.StartsWith("P|") || line.StartsWith("OBR|") || line.StartsWith("OBX|"))
                        {
                            //FileResultLog("Собираем строки для файла");
                            RequestMessage = RequestMessage + line + "\r";

                        }

                        // Если достигли конца файла с результатом, нужно сформировать результирующий файл для последнего ШК в результатах
                        if (line.StartsWith("L||"))
                        {
                            bufResultLog(RequestMessage);
                            bufResultLog("Формируем файл с результатом для последнего ШК");
                            //FileResultLog(RequestMessage);
                            bufResultLog("");
                            CreateResultFile(RequestMessage);

                            // Исходный большой файл перемещаем в папку archive
                            OriginalFileToArchive(file, FileName);
                        }

                    }
                }
                Thread.Sleep(1000);
            }

        }
        #endregion

        #region Функция обработки файлов с результатами и создания файлов для службы, которая разберет файл и запишет данные в CGM
        static void ResultsProcessing()
        {
            while (ServiceIsActive)
            {
                #region папки архива, результатов и ошибок

                string OutFolder = ConfigurationManager.AppSettings["FolderOut"];
                // архивная папка
                string ArchivePath = AnalyzerResultPath + @"\Archive";
                // папка для ошибок
                string ErrorPath = AnalyzerResultPath + @"\Error";
                // папка для файлов с результатами для CGM
                string CGMPath = AnalyzerResultPath + @"\CGM";

                if (!Directory.Exists(ArchivePath))
                {
                    Directory.CreateDirectory(ArchivePath);
                }

                if (!Directory.Exists(ErrorPath))
                {
                    Directory.CreateDirectory(ErrorPath);
                }

                //if (!Directory.Exists(CGMPath))
                //{
                    //Directory.CreateDirectory(CGMPath);
                //}
                #endregion

                // строки для формирования файла (psm файла) с результатами для службы,
                // которая разбирает файлы и записывает результаты в CGM
                //string MessageHead = "";
                //string MessageTest = "";
                //string AllMessage = "";

                // получаем список всех файлов в текущей папке
                string[] Files = Directory.GetFiles(AnalyzerResultPath, "*.res");

                // шаблоны регулярных выражений для поиска данных
                string RIDPattern = @"P[|]\d+[|]R-(?<RID>\d+)[|]\S*";
                string TestPattern = @"OBR[|]\d+[|]{3}(?<Test>.+).*";
                string ResultPattern = @"OBX[|]\d+[|]ST[|]{2}(?<Result>\d+[.]?\d*)[|]+";

                Regex RIDRegex = new Regex(RIDPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                Regex TestRegex = new Regex(TestPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));
                Regex ResultRegex = new Regex(ResultPattern, RegexOptions.None, TimeSpan.FromMilliseconds(150));

                // пробегаем по файлам
                foreach (string file in Files)
                {
                    // строки для формирования файла (psm файла) с результатами для службы,
                    // которая разбирает файлы и записывает результаты в CGM
                    string MessageHead = "";
                    string MessageTest = "";
                    string AllMessage = "";

                    FileResultLog("Обработка файлов с результатами и формирование файлов для CGM");
                    FileResultLog(file);

                    string[] lines = System.IO.File.ReadAllLines(file);
                    string RID = "";
                    string Test = "";
                    string Result = "";

                    string PSMTestCode = "";

                    bool isInterpretationOK = false;

                    // обрезаем только имя текущего файла
                    string FileName = file.Substring(AnalyzerResultPath.Length + 1);
                    // название файла .ок, который должен создаваться вместе с результирующим для обработки службой FileGetterService
                    string OkFileName = "";

                    // проходим по строкам в файле
                    foreach (string line in lines)
                    {
                        //FileResultLog(line);
                        string line_ = line.Replace("^", "@");
                        Match RIDMatch = RIDRegex.Match(line_);
                        Match TestMatch = TestRegex.Match(line_);
                        Match ResultMatch = ResultRegex.Match(line_);

                        // поиск RID в строке
                        if (RIDMatch.Success)
                        {
                            RID = RIDMatch.Result("${RID}");
                            FileResultLog($"Заявка № {RID}");
                            MessageHead = $"O|1|{RID}||ALL|R|20230101000100|||||X||||ALL||||||||||F";
                        }

                        // поиск теста в строке
                        if (TestMatch.Success)
                        {
                            Test = TestMatch.Result("${Test}");
                            // преобразуем тест в код теста PSM
                            FileResultLog($"Тест: {Test}");
                            PSMTestCode = "";
                            PSMTestCode = TranslateToPSMCodes(Test);
                            //FileResultLog($"{Test} преобразован в код CGM (PSMV2): {PSMTestCode}");

                            if (PSMTestCode == "NOD" || PSMTestCode == "")
                            {
                                FileResultLog($"{Test} не получилось интерпретировать в код CGM (PSMV2): {PSMTestCode}");
                                isInterpretationOK = false;
                            }
                            else
                            {
                                isInterpretationOK = true;
                                FileResultLog($"{Test} преобразован в код CGM (PSMV2): {PSMTestCode}");
                            }
                        }

                        // поиск результата в строке
                        if (ResultMatch.Success)
                        {
                            Result = ResultMatch.Result("${Result}");
                            //FileResultLog($"PSMV2 код: {PSMTestCode}");
                            FileResultLog($"Результат: {Result}");

                            // формируем строку с ответом для результирующего файла
                            //MessageTest = MessageTest + $"R|1|^^^{PSMTestCode}^^^^{AnalyzerCode}|{Result}|||N||F||Chemwell^||20230101000001|{AnalyzerCode}" + "\r";

                            // Если тест, который посылает прибор, был интерпретирован в PSMV2 код
                            if (isInterpretationOK)
                            {
                                //FileResultLog($"{PSMTestCode} услвоие выполнено");
                                // формируем строку с ответом для результирующего файла
                                MessageTest = MessageTest + $"R|1|^^^{PSMTestCode}^^^^{AnalyzerCode}|{Result}|||N||F||Chemwell^||20230101000001|{AnalyzerCode}" + "\r";
                            }

                            // обнулим переменные теста и результата
                            Test = "";
                            Result = "";
                        }
                    }

                    // получаем название файла .ок на основании файла с результатом
                    if (FileName.IndexOf(".") != -1)
                    {
                        OkFileName = FileName.Split('.')[0] + ".ok";
                    }

                    // если строки с результатами и с ШК не пустые, значит формируем результирующий файл
                    if (MessageHead != "" && MessageTest != "")
                    {
                        try
                        {
                            // собираем полное сообщение с результатом
                            AllMessage = MessageHead + "\r" + MessageTest;
                            FileResultLog(AllMessage);

                            // создаем файл для записи результата в папке для рез-тов
                            //if (!File.Exists(CGMPath + @"\" + FileName))
                            if (!File.Exists(OutFolder + @"\" + FileName))
                            {
                                //using (StreamWriter sw = File.CreateText(CGMPath + @"\" + FileName))
                                using (StreamWriter sw = File.CreateText(OutFolder + @"\" + FileName))
                                {
                                    foreach (string msg in AllMessage.Split('\r'))
                                    {
                                        sw.WriteLine(msg);
                                    }
                                }
                            }
                            else
                            {
                                //File.Delete(CGMPath + @"\" + FileName);
                                //using (StreamWriter sw = File.CreateText(CGMPath + @"\" + FileName))
                                File.Delete(OutFolder + @"\" + FileName);
                                using (StreamWriter sw = File.CreateText(OutFolder + @"\" + FileName))
                                {
                                    foreach (string msg in AllMessage.Split('\r'))
                                    {
                                        sw.WriteLine(msg);
                                    }
                                }
                            }

                            // создаем .ok файл в папке для рез-тов
                            if (OkFileName != "")
                            {
                                //if (!File.Exists(CGMPath + @"\" + OkFileName))
                                if (!File.Exists(OutFolder + @"\" + OkFileName))
                                {
                                    //using (StreamWriter sw = File.CreateText(CGMPath + @"\" + OkFileName))
                                    using (StreamWriter sw = File.CreateText(OutFolder + @"\" + OkFileName))
                                    {
                                        sw.WriteLine("ok");
                                    }
                                }
                                else
                                {
                                    //File.Delete(CGMPath + OkFileName);
                                    //using (StreamWriter sw = File.CreateText(CGMPath + @"\" + OkFileName))
                                    File.Delete(OutFolder + OkFileName);
                                    using (StreamWriter sw = File.CreateText(OutFolder + @"\" + OkFileName))
                                    {
                                        sw.WriteLine("ok");
                                    }
                                }
                            }

                            // помещение файла в архивную папку
                            if (File.Exists(ArchivePath + @"\" + FileName))
                            {
                                File.Delete(ArchivePath + @"\" + FileName);
                            }
                            File.Move(file, ArchivePath + @"\" + FileName);

                            FileResultLog("Файл обработан и перемещен в папку Archive");
                            FileResultLog("");
                        }
                        catch (Exception ex)
                        {
                            FileResultLog(ex.ToString());
                            // помещение файла в папку с ошибками
                            if (File.Exists(ErrorPath + @"\" + FileName))
                            {
                                File.Delete(ErrorPath + @"\" + FileName);
                            }
                            File.Move(file, ErrorPath + @"\" + FileName);

                            FileResultLog("Ошибка обработки файла. Файл перемещен в папку Error");
                            FileResultLog("");
                        }
                    }
                    else
                    {
                        // помещение файла в папку с ошибками
                        if (File.Exists(ErrorPath + @"\" + FileName))
                        {
                            File.Delete(ErrorPath + @"\" + FileName);
                        }
                        File.Move(file, ErrorPath + @"\" + FileName);

                        FileResultLog("Ошибка обработки файла. Файл перемещен в папку Error");
                        FileResultLog("");
                    }
                }
                Thread.Sleep(2000);
            }
        }

        #endregion


        protected override void OnStart(string[] args)
        {
            ServiceIsActive = true;
            ServiceLog("Сервис начал работу.");

            //Поток, который следит за другими потоками
            Thread ManagerThread = new Thread(new ThreadStart(CheckThreads));
            ManagerThread.Name = "Thread Manager";
            ManagerThread.Start();

            //Поток, который сохраняет файлы в папку из базы
            Thread ResultGetterThread = new Thread(new ThreadStart(ResultGetterFunction));
            ResultGetterThread.Name = "Result Getter";
            ListOfThreads.Add(ResultGetterThread);
            ResultGetterThread.Start();

            //Поток, который обрабатывает исходные файлы от прибора и разделяет их на разные файлы, в зависимости от количества заявок в исходном файле
            Thread OriginalFileProcessingThread = new Thread(new ThreadStart(bufFilesProcessing));
            OriginalFileProcessingThread.Name = "Orinial analyzer files handler";
            ListOfThreads.Add(OriginalFileProcessingThread);
            OriginalFileProcessingThread.Start();


            // Поток обработки результатов
            Thread ResultProcessingThread = new Thread(ResultsProcessing);
            ResultProcessingThread.Name = "ResultsProcessing";
            ListOfThreads.Add(ResultProcessingThread);
            ResultProcessingThread.Start();
        }

        protected override void OnStop()
        {
            ServiceIsActive = false;
            ServiceLog("Сервис остановлен");
        }
    }
}
