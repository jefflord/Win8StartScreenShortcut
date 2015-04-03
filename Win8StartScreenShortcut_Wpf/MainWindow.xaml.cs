using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using MS.WindowsAPICodePack.Internal;
using System;
using System.Diagnostics;
using System.Windows;
using Win8StartScreenShortcut_Wpf.ShellHelpers;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;


namespace Win8StartScreenShortcut_Wpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const String APP_ID = "JL.Win8StartScreenShortcut";
        private const String X1 = "JL.Win8StartScreenShortcut";
        private const String X2 = "JL.jefflord";


        public MainWindow()
        {
            InitializeComponent();

            TryCreateShortcut();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ShowToast();
        }

        private void ShowToast()
        {

            ////string toastXmlString = "<toast>"
            ////                   + "<visual version='1'>"
            ////                   + "<binding template='ToastText04'>"
            ////                   + "<text id='1'>Heading text</text>"
            ////                   + "<text id='2'>First body text</text>"
            ////                   + "<text id='3'>Second body text</text>"
            ////                   + "</binding>"
            ////                   + "</visual>"
            ////                   + "</toast>";


            XmlDocument toastXml = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastText01);

            XmlNodeList stringElements = toastXml.GetElementsByTagName("text");
            for (int i = 0; i < stringElements.Length; i++)
            {
                stringElements[i].AppendChild(toastXml.CreateTextNode("Line " + i));
            }


            //Windows.Data.Xml.Dom.XmlDocument toastDOM = new Windows.Data.Xml.Dom.XmlDocument();
            //toastDOM.LoadXml(toastXmlString);

            //// Create a toast, then create a ToastNotifier object to show
            //// the toast


            //string x = @"<toast><visual version=""1""><binding template=""ToastText01""><text id=""1"">Body text that wraps over three lines</text></binding></visual></toast>";
            //Windows.Data.Xml.Dom.XmlDocument toastXml = new Windows.Data.Xml.Dom.XmlDocument();
            //toastXml.LoadXml(x);


            ToastNotification toast = new ToastNotification(toastXml);
            toast.Activated += ToastActivated;
            toast.Dismissed += ToastDismissed;
            toast.Failed += ToastFailed;


            // If you have other applications in your package, you can specify the AppId of
            // the app to create a ToastNotifier for that application

            try
            {
                ToastNotificationManager.CreateToastNotifier(APP_ID).Show(toast);
            }
            catch (Exception e)
            {
                MessageBox.Show("EX:" + e.Message);
                //Application.Exit();
            }

        }

        private void ToastActivated(ToastNotification sender, object e)
        {
            Dispatcher.Invoke(() =>
            {
                Activate();
                MessageBox.Show("The user activated the toast.");
            });
        }

        private void ToastDismissed(ToastNotification sender, ToastDismissedEventArgs e)
        {
            String outputText = "";
            switch (e.Reason)
            {
                case ToastDismissalReason.ApplicationHidden:
                    outputText = "The app hid the toast using ToastNotifier.Hide";
                    break;
                case ToastDismissalReason.UserCanceled:
                    outputText = "The user dismissed the toast";
                    break;
                case ToastDismissalReason.TimedOut:
                    outputText = "The toast has timed out";
                    break;
            }

            Dispatcher.Invoke(() =>
            {
                MessageBox.Show(outputText);
            });
        }

        private void ToastFailed(ToastNotification sender, ToastFailedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show("The toast encountered an error.");
            });
        }

        private bool TryCreateShortcut()
        {
            String shortcutPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Windows\\Start Menu\\Programs\\Win8StartScreenShortcut_Wpf.lnk";
            if (!System.IO.File.Exists(shortcutPath))
            {
                InstallShortcut(shortcutPath);
                return true;
            }
            return false;
        }

        private void InstallShortcut(String shortcutPath)
        {
            // Find the path to the current executable
            String exePath = Process.GetCurrentProcess().MainModule.FileName;
            IShellLinkW newShortcut = (IShellLinkW)new CShellLink();

            // Create a shortcut to the exe            
            ShellHelpers.ErrorHelper.VerifySucceeded(newShortcut.SetPath(exePath));
            ShellHelpers.ErrorHelper.VerifySucceeded(newShortcut.SetArguments(""));

            // Open the shortcut property store, set the AppUserModelId property
            IPropertyStore newShortcutProperties = (IPropertyStore)newShortcut;

            using (PropVariant appId = new PropVariant(APP_ID))
            {
                ShellHelpers.ErrorHelper.VerifySucceeded(newShortcutProperties.SetValue(SystemProperties.System.AppUserModel.ID, appId));
                ShellHelpers.ErrorHelper.VerifySucceeded(newShortcutProperties.Commit());
            }

            // Commit the shortcut to disk
            IPersistFile newShortcutSave = (IPersistFile)newShortcut;

            ShellHelpers.ErrorHelper.VerifySucceeded(newShortcutSave.Save(shortcutPath, true));
        }

    }
}
