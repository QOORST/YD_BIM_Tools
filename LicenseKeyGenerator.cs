// LicenseKeyGenerator.cs
using System;
using System.Xml;
using Newtonsoft.Json;
using YD_RevitTools.LicenseManager;

namespace YD_RevitTools.KeyGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("╔════════════════════════════════════════════╗");
            Console.WriteLine("║  YD BIM 工具 - 授權金鑰生成器           ║");
            Console.WriteLine("╚════════════════════════════════════════════╝\n");

            while (true)
            {
                try
                {
                    // 選擇授權類型
                    Console.WriteLine("\n請選擇授權類型:");
                    Console.WriteLine("1. 試用版 (30天)");
                    Console.WriteLine("2. 標準版 (365天)");
                    Console.WriteLine("3. 專業版 (365天)");
                    Console.WriteLine("0. 退出");
                    Console.Write("\n請輸入選項 (0-3): ");

                    string choice = Console.ReadLine();

                    if (choice == "0")
                        break;

                    LicenseType licenseType;
                    int days;

                    switch (choice)
                    {
                        case "1":
                            licenseType = LicenseType.Trial;
                            days = 30;
                            break;
                        case "2":
                            licenseType = LicenseType.Standard;
                            days = 365;
                            break;
                        case "3":
                            licenseType = LicenseType.Professional;
                            days = 365;
                            break;
                        default:
                            Console.WriteLine("無效的選項，請重新選擇。");
                            continue;
                    }

                    // 輸入使用者資訊
                    Console.Write("\n使用者名稱: ");
                    string userName = Console.ReadLine();

                    if (string.IsNullOrWhiteSpace(userName))
                    {
                        Console.WriteLine("使用者名稱不能為空！");
                        continue;
                    }

                    Console.Write("公司名稱: ");
                    string company = Console.ReadLine();

                    if (string.IsNullOrWhiteSpace(company))
                    {
                        Console.WriteLine("公司名稱不能為空！");
                        continue;
                    }

                    Console.Write("機器碼 (選填，按Enter跳過): ");
                    string machineCode = Console.ReadLine();

                    // 設定起始日期
                    Console.Write("起始日期 (格式: yyyy-MM-dd，按Enter使用今天): ");
                    string startDateStr = Console.ReadLine();
                    DateTime startDate = string.IsNullOrWhiteSpace(startDateStr)
                        ? DateTime.Now.Date
                        : DateTime.Parse(startDateStr);

                    // 創建授權資訊
                    var license = new LicenseInfo
                    {
                        IsEnabled = true,
                        LicenseType = licenseType,
                        UserName = userName,
                        Company = company,
                        StartDate = startDate,
                        ExpiryDate = startDate.AddDays(days),
                        LicenseKey = Guid.NewGuid().ToString(),
                        MachineCode = machineCode
                    };

                    // 生成授權金鑰
                    string json = JsonConvert.SerializeObject(license, Newtonsoft.Json.Formatting.None);
                    string licenseKey = Convert.ToBase64String(
                        System.Text.Encoding.UTF8.GetBytes(json));

                    // 顯示結果
                    Console.WriteLine("\n╔════════════════════════════════════════════╗");
                    Console.WriteLine("║            授權金鑰生成成功              ║");
                    Console.WriteLine("╚════════════════════════════════════════════╝");
                    Console.WriteLine($"\n授權類型: {license.GetLicenseTypeName()}");
                    Console.WriteLine($"使用者: {userName}");
                    Console.WriteLine($"公司: {company}");
                    Console.WriteLine($"起始日期: {startDate:yyyy-MM-dd}");
                    Console.WriteLine($"到期日期: {license.ExpiryDate:yyyy-MM-dd}");
                    Console.WriteLine($"有效期限: {days} 天");
                    if (!string.IsNullOrWhiteSpace(machineCode))
                        Console.WriteLine($"綁定機器碼: {machineCode}");

                    Console.WriteLine("\n授權金鑰:");
                    Console.WriteLine("─────────────────────────────────────────");
                    Console.WriteLine(licenseKey);
                    Console.WriteLine("─────────────────────────────────────────");

                    Console.WriteLine("\n請將此金鑰提供給用戶。");

                    // 選擇是否儲存到文件
                    Console.Write("\n是否儲存到文件？ (Y/N): ");
                    if (Console.ReadLine()?.ToUpper() == "Y")
                    {
                        string fileName = $"License_{userName}_{company}_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                        System.IO.File.WriteAllText(fileName,
                            $"YD BIM 工具授權金鑰\n" +
                            $"生成時間: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n\n" +
                            $"授權類型: {license.GetLicenseTypeName()}\n" +
                            $"使用者: {userName}\n" +
                            $"公司: {company}\n" +
                            $"起始日期: {startDate:yyyy-MM-dd}\n" +
                            $"到期日期: {license.ExpiryDate:yyyy-MM-dd}\n" +
                            $"有效期限: {days} 天\n" +
                            (!string.IsNullOrWhiteSpace(machineCode) ? $"綁定機器碼: {machineCode}\n" : "") +
                            $"\n授權金鑰:\n{licenseKey}");

                        Console.WriteLine($"已儲存到: {fileName}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\n錯誤: {ex.Message}");
                }

                Console.WriteLine("\n按任意鍵繼續...");
                Console.ReadKey();
                Console.Clear();
                Console.WriteLine("╔════════════════════════════════════════════╗");
                Console.WriteLine("║  YD BIM 工具 - 授權金鑰生成器           ║");
                Console.WriteLine("╚════════════════════════════════════════════╝\n");
            }
        }
    }
}