using System;
using System.Management;
using System.Security.Cryptography;
using System.Text;

namespace 课件帮PPT助手
{
    public static class HardwareInfoHelper
    {
        public static string GetHardwareId()
        {
            string cpuId = GetCpuId();
            string diskId = GetDiskId();

            // 使用 CPU ID 和硬盘 ID 生成唯一的硬件 ID
            return cpuId + diskId;
        }

        private static string GetCpuId()
        {
            string cpuId = string.Empty;

            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("select ProcessorId from Win32_Processor"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        cpuId = obj["ProcessorId"].ToString();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // 处理获取 CPU ID 时的异常
                Console.WriteLine("获取 CPU ID 失败: " + ex.Message);
            }

            return cpuId;
        }

        private static string GetDiskId()
        {
            string diskId = string.Empty;

            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_PhysicalMedia"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        diskId = obj["SerialNumber"].ToString();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // 处理获取硬盘序列号时的异常
                Console.WriteLine("获取硬盘序列号失败: " + ex.Message);
            }

            return diskId;
        }

        public static string GenerateActivationCode(string hardwareId)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(hardwareId));
                StringBuilder stringBuilder = new StringBuilder();

                // 将哈希值转换为十六进制字符串
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    stringBuilder.Append(hashBytes[i].ToString("X2"));
                }

                // 截取前16位作为激活码
                return stringBuilder.ToString().Substring(0, 16);
            }
        }
    }
}
