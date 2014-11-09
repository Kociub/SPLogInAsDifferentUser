using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SHDocVw;
using mshtml;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Web;

namespace SPLogInAsDifferentUser
{
    [
    ComVisible(true),
    Guid("FD3AF5F8-CFC3-4A50-972E-9E9DC937157B"),
    ClassInterface(ClassInterfaceType.None)
    ]
    public class BHO : IObjectWithSite
    {
        public static string BhoKeyName =
  "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";
        public static string AcceessDeniedUrlPart = "AccessDenied.aspx?Source=";
        public static string RedirectButton = @"<button onclick='location.href = 'www.yoursite.com';' id='myButton' class='float-left submit-button'>TEST</button>";

        WebBrowser webBrowser;
        HTMLDocument document;

        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            document = (HTMLDocument)webBrowser.Document;

            System.Windows.Forms.MessageBox.Show(document.url);

            int indexOfAccessDenied = document.url.IndexOf(AcceessDeniedUrlPart);

            if(indexOfAccessDenied > 0)
            {
                System.Windows.Forms.MessageBox.Show(String.Format("Index of access denied={0}", indexOfAccessDenied));

                string sourceUrl = HttpUtility.UrlDecode(document.url.Substring(indexOfAccessDenied));
            }

            IHTMLElement element = document.getElementById("ms-error-header");

            element.insertAdjacentHTML("beforeBegin", RedirectButton);

            //AccessDenied.aspx?Source=http%3A%2F%2Fwitisdev%2Ewint%2Ech

            /*
            foreach (
                IHTMLInputElement tempElement in document.getElementsByTagName("INPUT"))
            {
                
                System.Windows.Forms.MessageBox.Show(
                    tempElement.name != null ? tempElement.name : "it sucks, no name, try id" +
                        ((IHTMLElement)tempElement).id
                    );
            }
            */
        }
 
        public int SetSite(object site)
        {
            if (site != null)
            {
                webBrowser = (WebBrowser)site;
                webBrowser.DocumentComplete +=
                new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                webBrowser.BeforeNavigate2 +=
                   new DWebBrowserEvents2_BeforeNavigate2EventHandler(this.OnBeforeNavigate2);
            }
            else
            {
                webBrowser.DocumentComplete -=
                  new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                webBrowser.BeforeNavigate2 -=
                  new DWebBrowserEvents2_BeforeNavigate2EventHandler(this.OnBeforeNavigate2);
                webBrowser = null;
            }

            return 0;
        }

        public void OnBeforeNavigate2(object pDisp, ref object URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancel)
        {
            document = (HTMLDocument)webBrowser.Document;

            foreach (IHTMLInputElement tempElement in document.getElementsByTagName("INPUT"))
            {
                if (tempElement.type.ToLower() == "password")
                {
                    System.Windows.Forms.MessageBox.Show(tempElement.value);
                }
            }
        }

        public int GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);

            return hr;
        }

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BhoKeyName, true);

            if (registryKey == null)
                registryKey = Registry.LocalMachine.CreateSubKey(BhoKeyName);

            string guid = type.GUID.ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid);

            if (ourKey == null)
                ourKey = registryKey.CreateSubKey(guid);

            ourKey.SetValue("Alright", 1);
            registryKey.Close();
            ourKey.Close();
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BhoKeyName, true);
            string guid = type.GUID.ToString("B");

            if (registryKey != null)
                registryKey.DeleteSubKey(guid, false);
        }
    }
}
