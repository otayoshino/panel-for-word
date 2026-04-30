using System;
using Microsoft.Win32;

namespace panel_for_word_installer
{
    public partial class ThisAddIn
    {
        private const string ManifestUrl =
            "https://otayoshino.github.io/panel-for-word/manifest.xml";
        private const string AddInId = "com.otayoshino.panel-for-word";

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            RegisterWebAddIn();
        }

        private void RegisterWebAddIn()
        {
            try
            {
                string regPath =
                    @"Software\Microsoft\Office\16.0\WEF\Developer";

                using (RegistryKey key = Registry.CurrentUser
                    .CreateSubKey(regPath, true))
                {
                    key.SetValue(AddInId, ManifestUrl);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"登録エラー: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}