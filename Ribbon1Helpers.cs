using System;
using System.IO;
using 课件帮PPT助手;

internal static class Ribbon1Helpers
{

    // 唯一的 GetResourceText 方法
    private static string GetResourceText(string resourceName)
    {
        var asm = typeof(Ribbon1).Assembly;
        using (var stream = asm.GetManifestResourceStream(resourceName))
        {
            if (stream == null)
            {
                return null;
            }
            using (var reader = new System.IO.StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }


    private static string getResourceText(string resourceName)
    {
        var asm = System.Reflection.Assembly.GetExecutingAssembly();
        string[] resourceNames = asm.GetManifestResourceNames();
        foreach (string resource in resourceNames)
        {
            if (string.Compare(resourceName, resource, StringComparison.OrdinalIgnoreCase) == 0)
            {
                using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resource)))
                {
                    if (resourceReader != null)
                    {
                        return resourceReader.ReadToEnd();
                    }
                }
            }
        }
        return null;
    }
}