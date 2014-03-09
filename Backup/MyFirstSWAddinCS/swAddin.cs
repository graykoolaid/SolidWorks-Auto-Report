using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using SldWorks;
using SWPublished;
using SwConst;
using SwCommands;

using SolidWorksTools;
using SolidWorksTools.File;

namespace MyFirstSWAddinCS
{
    [Guid("883d6c36-100b-4036-a888-27c88b394881"), ComVisible(true)]
    [SwAddin(Description = "My addin description", Title = "My First Addin", LoadAtStartup = true)]
    public class swAddin : ISwAddin
    {
        ISldWorks iSwApp;

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            iSwApp.SendMsgToUser2("Hello!", (int)swMessageBoxIcon_e.swMbInformation, (int)swMessageBoxBtn_e.swMbOk);
            return true;
        }

        public bool DisconnectFromSW()
        {
            iSwApp.SendMsgToUser2("Bye!", (int)swMessageBoxIcon_e.swMbInformation, (int)swMessageBoxBtn_e.swMbOk);
            iSwApp = null;
            GC.Collect();
            return true;
        }

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

            string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
            Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
            addinkey.SetValue(null, 0);
            addinkey.SetValue("Description", "Sample Addin");
            addinkey.SetValue("Title", "MyFirstAddin");

            keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
            addinkey = hkcu.CreateSubKey(keyname);
            addinkey.SetValue(null, 1);
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            //Insert code here.
            Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

            string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
            hklm.DeleteSubKey(keyname);

            keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
            hkcu.DeleteSubKey(keyname);
        }
    }
}
