// MachineCodeHelper.cs
using System;
using System.Management;
using System.Security.Cryptography;
using System.Text;

namespace YD_RevitTools.LicenseManager
{
    public static class MachineCodeHelper
    {
        // 生成機器碼（基於硬體資訊）
        public static string GetMachineCode()
        {
            try
            {
                string cpuId = GetCpuId();
                string motherboardId = GetMotherboardId();

                string combined = $"{cpuId}-{motherboardId}";

                using (MD5 md5 = MD5.Create())
                {
                    byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(combined));
                    return BitConverter.ToString(hash).Replace("-", "").Substring(0, 16);
                }
            }
            catch
            {
                return "UNKNOWN";
            }
        }

        private static string GetCpuId()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor");
                foreach (ManagementObject obj in searcher.Get())
                {
                    return obj["ProcessorId"]?.ToString() ?? "UNKNOWN";
                }
            }
            catch { }
            return "UNKNOWN";
        }

        private static string GetMotherboardId()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard");
                foreach (ManagementObject obj in searcher.Get())
                {
                    return obj["SerialNumber"]?.ToString() ?? "UNKNOWN";
                }
            }
            catch { }
            return "UNKNOWN";
        }
    }
}