using Microsoft.Win32;
using System.IO;

namespace MSOfficeManager.API
{
    internal class Static
    {
        public static bool IsRegistred(string name)
        {
            RegistryKey view32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            RegistryKey view64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey k32 = view32.OpenSubKey("SOFTWARE")?.OpenSubKey("Microsoft")?.OpenSubKey("Office");
            RegistryKey k64 = view64.OpenSubKey("SOFTWARE")?.OpenSubKey("Microsoft")?.OpenSubKey("Office");
            string p32 = RegistrySearch(k32, "Office") as string;
            string p64 = RegistrySearch(k64, "Office") as string;
            if (p32 != null) p32 = Path.Combine(p32, "WINWORD.EXE");
            if (p64 != null) p64 = Path.Combine(p64, "WINWORD.EXE");
            if (File.Exists(p32))
                return true;
            else if (File.Exists(p64))
                return true;
            return false;
        }

        public static object RegistrySearch(RegistryKey key, string name)
        {
            if (key != null && name == "InstallRoot")
                return key.GetValue("Path");
            string[] child = key.GetSubKeyNames();
            if (child != null)
                for (int i = 0; i < child.Length; i++)
                {
                    RegistryKey k = key.OpenSubKey(child[i]);
                    object r = RegistrySearch(k, child[i]);
                    if (r != null)
                        return r;
                }
            return null;
        }
    }
}
