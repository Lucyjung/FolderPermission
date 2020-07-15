using System;
using System.Configuration;
using System.Windows.Forms;

namespace FolderPermission.Utilities
{
    public class Config
    {
        public static string fileServer;
        public static string[] ignoreList;
        public static string domain;
        public static void GetConfigurationValue()
        {
            try
            {
                fileServer = ConfigurationManager.AppSettings["fileServer"];
                domain = ConfigurationManager.AppSettings["domain"];
                ignoreList = ConfigurationManager.AppSettings["ignoreList"].Split(',');
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        public static bool isIgnore(string inStr)
        {
            bool isIgnore = false;
            foreach (string ignore in Config.ignoreList)
            {
                if (inStr.Contains(ignore))
                {
                    isIgnore = true;
                    break;
                }
            }
            return isIgnore;
        }
    }
}
