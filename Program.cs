using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;
using System.Net;
using System.Globalization;
using static System.Environment;
using IWshRuntimeLibrary;
using File = IWshRuntimeLibrary.File;


// нужно добавить ссылку


namespace SDP_Diagnostics
{
    class SDP
    {

        /// <summary>
        /// Информации о установленной ОС и железе.
        /// </summary>
        /// <param name="system_Info"></param>
        /// <param name="Client"></param>
        public static void ShowSystemInfo(string system_Info)
        {


            FileStream fs = new FileStream(system_Info, FileMode.Append, FileAccess.Write); //Открывает и записывает в файл информацию
            StreamWriter w = new StreamWriter(fs, Encoding.Default); // кодирует информацию из потока fs
            {
                w.WriteLine("Операционная система (номер версии): {0}", OSVersion.Version);
                w.WriteLine("Разрядность процессора:  {0}", arg0: GetEnvironmentVariable("PROCESSOR_ARCHITECTURE"));
                w.WriteLine("Модель процессора:  {0}", GetEnvironmentVariable("PROCESSOR_IDENTIFIER"));
                w.WriteLine("Путь к системному каталогу:  {0}", SystemDirectory);
                w.WriteLine("Число процессоров:  {0}", ProcessorCount);
                w.WriteLine("Имя пользователя: {0}", UserName);





               
                DriveInfo[] allDrives = DriveInfo.GetDrives();

                w.WriteLine($"Количество логических дисков на компьютере = {allDrives.Count()}");
                foreach (DriveInfo currDrvInf in allDrives)
                {
                    w.WriteLine($"Имя = {currDrvInf.Name}");
                    w.WriteLine($" Тип диска = { currDrvInf.DriveType.ToString()}");
                    if (currDrvInf.IsReady)
                    {
                        w.WriteLine(" Формат файловой системы          = " + currDrvInf.DriveFormat);
                        w.WriteLine(" Общий размер                     = " + ((currDrvInf.TotalSize / Math.Pow(1024d, 3d))).ToString(CultureInfo.InvariantCulture) + " Гб.");
                        w.WriteLine(" Свободное место                  = " + ((currDrvInf.TotalFreeSpace / Math.Pow(1024d, 3d))).ToString(CultureInfo.InvariantCulture) + " Гб.");
                        w.WriteLine(" Доступное свободное пространство = " + (currDrvInf.AvailableFreeSpace / Math.Pow(1024d, 3d)).ToString(CultureInfo.InvariantCulture) + " Гб.");
                        w.WriteLine(" Метка тома                       = " + currDrvInf.VolumeLabel);
                        w.WriteLine(" Корневой каталог                 = " + currDrvInf.RootDirectory.FullName);
                    }
                    else
                    {
                        w.WriteLine(" Диск не готов! Другая информация не доступна!");
                    }


                }


                
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_VideoController");

                foreach (ManagementBaseObject o in searcher.Get())

                {
                    var query_obj = (ManagementObject)o;
                    w.WriteLine($"Название:  {query_obj["Caption"]}");
                    w.WriteLine(value: $"Семейство:  {query_obj["VideoProcessor"]}");
                    w.WriteLine(value: $"Обьем: {query_obj["AdapterRAM"]}");
                }

                
                ManagementObjectSearcher ram_monitor = //запрос к WMI для получения памяти ПК
                    new ManagementObjectSearcher("SELECT TotalVisibleMemorySize,FreePhysicalMemory FROM Win32_OperatingSystem");

                foreach (var o in ram_monitor.Get())
                {
                    var objram = (ManagementObject)o;
                    UInt64 total_ram = Convert.ToUInt64(objram["TotalVisibleMemorySize"]); //общая память ОЗУ
                    UInt64 busy_ram = total_ram - Convert.ToUInt64(objram["FreePhysicalMemory"]);

                    w.WriteLine($"Количество ОП:  {total_ram / 1024}");
                    w.WriteLine($"Процентное выражение занятой ОП:  {(busy_ram * 100) / total_ram}");

                }

                w.Close(); // закрывает поток и сам файлс!
            }
        }

        /// <summary>
        /// Запрос ключа для  фреймворка
        /// </summary>
        public static void Key_Framework(string System_Info)
        {
            const String sub_key = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\";

            FileStream fs = new FileStream(System_Info, FileMode.Append, FileAccess.Write);
            StreamWriter w = new StreamWriter(fs, Encoding.Default);
            RegistryKey ndp_key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(sub_key);

            if (ndp_key != null && ndp_key.GetValue("Release") != null)
            {
                String res = Ver_Framework(Convert.ToInt32(ndp_key.GetValue("Release"))); // Вызов метода
                w.WriteLine($"Версия фреймворка: .NET Framework Version{res}.");
            }
            else
            {
                w.Write(".NET Framework Version 4.5 or later is not detected.");
            }

            w.Close();


        }

        /// <summary>
        /// На основании полученного ключа выдает версию фреймворка.
        /// </summary>
        public static string Ver_Framework(int release_Key)
        {


            if (release_Key >= 528040)
                return "4.8 or later";
            if (release_Key >= 461808)
                return "4.7.2";
            if (release_Key >= 461308)
                return "4.7.1";
            if (release_Key >= 460798)
                return "4.7";
            if (release_Key >= 394802)
                return "4.6.2";
            if (release_Key >= 394254)
                return "4.6.1";
            if (release_Key >= 393295)
                return "4.6";
            if (release_Key >= 379893)
                return "4.5.2";
            if (release_Key >= 378675)
                return "4.5.1";
            if (release_Key >= 378389)
                return "4.5";
            return "No 4.5 or later version detected";
        }

        /// <summary>
        /// Записись логов и фул логов в файл.
        /// </summary>
        public static void Save_log(string Log_Error_Users, string Error_Log, string Log_Full_Users, string Full_Log)
        {
            try
            {
                System.IO.File.Copy(Full_Log, Log_Full_Users, true);

            }
            catch (System.IndexOutOfRangeException)
            {
                FileStream fs = new FileStream(Log_Full_Users, FileMode.Append, FileAccess.Write);
                StreamWriter w = new StreamWriter(fs, Encoding.Default);
                w.WriteLine("Файл FulLog не обнаружен!");
                throw;
            }

            try
            {
                System.IO.File.Copy(Error_Log, Log_Error_Users, true);
            }
            catch (System.IndexOutOfRangeException)
            {
                FileStream fs = new FileStream(Log_Full_Users, FileMode.Append, FileAccess.Write);
                StreamWriter w = new StreamWriter(fs, Encoding.Default);
                w.WriteLine("Файл ErorLog не обнаружен!");
                throw;
            }
           

        }

        /// <summary>
        /// Проверяет наличие докана и возвращает его версию.
        /// </summary>
        public static void Dokan(string System_Info)
        {
            FileStream f = new FileStream(System_Info, FileMode.Append, FileAccess.Write);
            StreamWriter w = new StreamWriter(f, Encoding.Default);
            string registryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey);
            if (key != null)
            {
                Int32 count = 0;
                foreach (String a in key.GetSubKeyNames())
                {
                    RegistryKey subkey = key.OpenSubKey(name: a);

                    if (subkey?.GetValue(name: "DisplayName") != null)
                    {

                        if (subkey.GetValue(name: "DisplayName").ToString().IndexOf(value: "Dokan", comparisonType: StringComparison.Ordinal) == 0)
                        {
                            w.WriteLine(subkey.GetValue(name: "DisplayName"));

                        }

                        if (key.GetSubKeyNames().Length == count)
                        {
                            w.Write(value: "Dokan не установлен");
                            break;
                        }
                    }


                    count++;
                }
                w.Close();
            }

        }

        /// <summary>
        /// Поиск антивируса.
        /// </summary>
        /// <param name="System_info"></param>
        public static void Antivirus(string System_Info)
        {
            FileStream f = new FileStream(System_Info, FileMode.Append, FileAccess.Write);
            StreamWriter w = new StreamWriter(f, Encoding.Default);
            String name = (String)(new ManagementObjectSearcher("root\\SecurityCenter2", "SELECT * FROM AntiVirusProduct").Get().Cast<ManagementObject>().LastOrDefault()?["displayName"]);
            w.WriteLine("Антивирус: " + name);
            w.Close();


        }

        /// <summary>
        /// IP_Config
        /// </summary>
        /// <param name="System_info"></param>
        public static void Ip_config(string System_Info)
        {
            FileStream f = new FileStream(System_Info, FileMode.Append, FileAccess.Write);
            StreamWriter w = new StreamWriter(f, Encoding.Default);
            string command1 = "ipconfig";
            string parametr = "/all";


            // создается ProcessStartInfo с использованием "CMD" в качестве программы для запуска
            // и "/c " в качестве параметров.
            // /c говорит CMD, что далее будет следовать команда для запуска
            ProcessStartInfo procStartInfo = new ProcessStartInfo("cmd", "/c " + command1 + parametr);
            // Следующая команды означает, что нужно перенаправить стандартынй вывод
            // на Process.StandardOutput StreamReader.
            procStartInfo.RedirectStandardOutput = true;
            procStartInfo.UseShellExecute = false;
            // не создавать окно CMD
            procStartInfo.CreateNoWindow = true;

            Process proc = new Process();
            // Получение текста в виде кодировки 866 win
            procStartInfo.StandardOutputEncoding = Encoding.GetEncoding(866);
            //запуск CMD
            proc.StartInfo = procStartInfo;
            proc.Start();
            //чтение результата
            string result = proc.StandardOutput.ReadToEnd();

            w.WriteLine(result);
            w.Close();




        }

        /// <summary>
        /// Трассировка и ip адрес.
        /// </summary>
        /// <param name="System_info"></param>
        
        public static void Tracing_and_IP(string System_Info)
        {
            FileStream f = new FileStream(System_Info, FileMode.Append, FileAccess.Write);
            StreamWriter w = new StreamWriter(f, Encoding.Default);
            String host = Dns.GetHostName();
            //Получение ip-адреса.
            IPAddress ip = Dns.GetHostByName(host).AddressList[0];
            //Показ адреса в label'е.
            w.WriteLine("\r\n ip adress:" + ip);
            string command = "tracert SDPSS.ru";

            // создается ProcessStartInfo с использованием "CMD" в качестве программы для запуска
            // и "/c " в качестве параметров.
            // /c говорит CMD, что далее будет следовать команда для запуска
            ProcessStartInfo procStartInfo =
                new ProcessStartInfo("cmd", "/c " + command);
            // Следующая команды означает, что нужно перенаправить стандартынй вывод
            // на Process.StandardOutput StreamReader.
            procStartInfo.RedirectStandardOutput = true;
            procStartInfo.UseShellExecute = false;
            // не создавать окно CMD
            procStartInfo.CreateNoWindow = true;

            Process proc = new Process();
            // Получение текста в виде кодировки 866 win
            procStartInfo.StandardOutputEncoding = Encoding.GetEncoding(866);
            //запуск CMD
            proc.StartInfo = procStartInfo;
            proc.Start();
            //чтение результата
            string result = proc.StandardOutput.ReadToEnd();

            w.WriteLine(result);
            w.Close();

        }

        /// <summary>
        /// Путь к SDP
        /// </summary>
        /// <param name="System_info"></param>
        public static void Path_SDP_and_Ver(string System_Info)
        {
            FileStream fs = new FileStream(System_Info, FileMode.Append, FileAccess.Write); 
            StreamWriter w = new StreamWriter(fs, Encoding.Default);
            string dir = GetFolderPath(SpecialFolder.Desktop);
            var di = new DirectoryInfo(dir);
            FileInfo[] fis = di.GetFiles();
            w.Write("Путь к SDPClient.exe: ");
            if (fis.Length > 0)
            {
                
                foreach (FileInfo fi in fis)
                {
                    if (fi.FullName.EndsWith("lnk"))
                    {
                        IWshShell shell = new WshShell();
                        var lnk = shell.CreateShortcut(fi.FullName) as IWshShortcut;
                        if (lnk != null)
                        {
                            String ww = lnk.TargetPath;
                            string[] ff = ww.Split(' ', '\\');

                            
                            for (int j = 0; j < ff.Length; j++)
                            {
                                if (ff[j] == "SDPClient.exe")
                                {

                                    string value = String.Concat<string>(ff);
                                    var SDPFileInfo = FileVersionInfo.GetVersionInfo(ww);
                                    w.WriteLine(SDPFileInfo);
                                    w.Close();
                                    return;
                                   
                                }
                                
                            }



                        }

                    }
                }



            }

           

            

        }

        /// <summary>
        /// Размер кэша
        /// </summary>
        /// <param name="Cash"></param>
        /// <param name="System_info"></param>
        public static void Cache_size(string Cash, string System_Info)
        {
            FileStream fs = new FileStream(System_Info, FileMode.Append, FileAccess.Write); 
            StreamWriter w = new StreamWriter(fs, Encoding.Default);
            FileInfo ww = new FileInfo(Cash);
            w.WriteLine("Размер кэша: " + (double)ww.Length / 1048576 + "МБ");
            w.Close();
            



        }

        /// <summary>
        /// Сравнение на целостность с помощью хэш-суммы.
        /// </summary>
        /// <param name="Hash_act"></param>

        public static void Hash(string Hash_act, string System_Info)
        {
            FileStream f = new FileStream(System_Info, FileMode.Append, FileAccess.Write);
            StreamWriter w = new StreamWriter(f, Encoding.Default);
            byte[] buffer = new byte[4096];
            using (var md5 = new MD5CryptoServiceProvider())
            {
                foreach (var file in Directory.EnumerateFiles(@"C:\Program Files (x86)\SDP3.0", "*", SearchOption.AllDirectories))
                {
                    int length;
                    using (var fs = System.IO.File.OpenRead(file))
                    {
                        do
                        {
                            length = fs.Read(buffer, 0, buffer.Length);
                            md5.TransformBlock(buffer, 0, length, buffer, 0);
                        } while (length > 0);
                    }
                }
                md5.TransformFinalBlock(buffer, 0, 0);
                string ww = BitConverter.ToString(md5.Hash);
                string[] Get_Hash = ww.Split(' ', '-');
                

                string[] Hash = Hash_act.Split(' ', '-');
                bool isEqual = Hash.SequenceEqual(Get_Hash);
                w.Write(isEqual ? "Хэш корневой папки целый." : "Имеются повреждения файла!");
                w.Close();
            }

        }






        
        static void Main()
        {
            String Log_Full_Users = Path.Combine(Path.GetTempPath(), "SDPdiagnostics", "Сбор данных (FullLog).txt");
            String Log_Error_Users = Path.Combine(Path.GetTempPath(), "SDPdiagnostics", "Сбор данных (ErorLog).txt");
            String System_Info = Path.Combine(Path.GetTempPath(), "SDPdiagnostics", "Сбор данных(сведения о системе).txt");
            String Full_Log = Path.Combine(GetFolderPath(folder: SpecialFolder.ApplicationData), "SDP3", "ClientLog", "full", "full_log.log");
            String Error_Log = Path.Combine(GetFolderPath(folder: SpecialFolder.ApplicationData), "SDP3", "ClientLog", "error", "error_log.log");
            String Cash = (@"C:\Program Files\Common Files\SDP3\FileCache\sdp_file_cache.db");
            String Hash_act = "4B-BF-1C-C3-AB-EB-51-CD-23-B4-F0-AD-98-20-35-62";
            String Direct = Path.Combine(Path.GetTempPath(), "SDPdiagnostics");
           
            DirectoryInfo F = Directory.CreateDirectory(Direct);

            ShowSystemInfo(System_Info);
            Save_log(Log_Error_Users, Error_Log, Log_Full_Users, Full_Log);
            Key_Framework(System_Info);
            Dokan(System_Info);
            Antivirus(System_Info);
            Ip_config(System_Info);
            Tracing_and_IP(System_Info);
            Path_SDP_and_Ver(System_Info);
            Cache_size(Cash, System_Info);
            Hash(Hash_act, System_Info);

            
            



        }














    }


}






