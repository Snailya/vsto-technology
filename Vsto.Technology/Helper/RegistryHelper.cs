using Microsoft.Win32;

namespace Vsto.Technology.Helper;

public static class RegistryHelper
{
    public static string GetRegistryValue(RegistryKey hive, string subKey, string valueName)
    {
        using var key = hive.OpenSubKey(subKey);
        if (key == null)
            throw new RegistryValueNotFoundException($"Subkey '{subKey}' not found in '{hive.Name}'.");

        var value = key.GetValue(valueName);
        if (value != null) return value.ToString();

        throw new RegistryValueNotFoundException($"Value '{valueName}' not found in '{subKey}'");
    }
}