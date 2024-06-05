using Microsoft.Win32;
using System;

namespace 课件帮PPT助手
{
    public static class SettingsHelper
    {
        private const string RegistryKeyPath = @"Software\课件帮PPT助手";
        private const string AlignmentValueName = "DefaultAlignment";

        public static void SaveAlignmentSetting(Alignment alignment)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                key.SetValue(AlignmentValueName, alignment.ToString());
            }
        }

        public static Alignment LoadAlignmentSetting()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
            {
                if (key != null)
                {
                    string alignmentValue = key.GetValue(AlignmentValueName)?.ToString();
                    if (Enum.TryParse(alignmentValue, out Alignment alignment))
                    {
                        return alignment;
                    }
                }
            }
            return Alignment.Center; // 默认值
        }
    }
}
