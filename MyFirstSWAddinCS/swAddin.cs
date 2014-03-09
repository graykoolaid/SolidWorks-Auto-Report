using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace MyFirstSWAddinCS
{
    public class values
    {
        // GUI related fields
        public String customer;
        public String date;
        public String name;
        public String part;
        public String serial;
        public String job;

        public List<String> p_file1;
        public List<String> p_file2;
        public List<String> v_ptol;

        public List<String> d_file1;
        public List<String> d_file2;
        public List<String> v_dtol;

        public List<String> r_file1;
        public List<String> r_file2;
        public List<String> v_rtol;

        public List<String> p_ball;
        public List<String> d_ball;
        public List<String> r_ball;

        public List<String> p_letter;
        public List<String> d_letter;
        public List<String> r_letter;

        public List<String> rd;


        // Other Data
        public List<String> target_vals;
        public List<String> deviat_vals;

        public values()
        {
            p_file1 = new List<String>();
            p_file2 = new List<String>();
            d_file1 = new List<String>();
            d_file2 = new List<String>();
            r_file1 = new List<String>();
            r_file2 = new List<String>();
            v_ptol = new List<String>();
            v_dtol = new List<String>();
            v_rtol = new List<String>();
            p_ball = new List<String>();
            d_ball = new List<String>();
            r_ball = new List<String>();
            p_letter = new List<String>();
            d_letter = new List<String>();
            r_letter = new List<String>();

            target_vals = new List<String>();
            deviat_vals = new List<String>();

            rd = new List<String>();

        }
    }

    [Guid("883d6c36-100b-4036-a888-27c88b394881"), ComVisible(true)]
    public class swAddin : ISwAddin
    {
        ISldWorks iSwApp;
        ICommandManager iCmdMgr;

        String path = "c:\\TempNinja";

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            iSwApp = (ISldWorks)ThisSW;

            // Check and create our directory
            //iSwApp.SendMsgToUser2(path, (int)swMessageBoxIcon_e.swMbInformation, (int)swMessageBoxBtn_e.swMbOk);
            if( !System.IO.Directory.Exists(path) )
                System.IO.Directory.CreateDirectory(path);

            if (!System.IO.Directory.Exists(path + "\\p1"))
                System.IO.Directory.CreateDirectory(path + "\\p1");

            if (!System.IO.Directory.Exists(path + "\\p2"))
                System.IO.Directory.CreateDirectory(path + "\\p2");

            if (!System.IO.Directory.Exists(path + "\\d1"))
                System.IO.Directory.CreateDirectory(path + "\\d1");

            if (!System.IO.Directory.Exists(path + "\\d2"))
                System.IO.Directory.CreateDirectory(path + "\\d2");

            if (!System.IO.Directory.Exists(path + "\\r1"))
                System.IO.Directory.CreateDirectory(path + "\\r1");

            if (!System.IO.Directory.Exists(path + "\\r2"))
                System.IO.Directory.CreateDirectory(path + "\\r2");

            if (!System.IO.Directory.Exists(@"C:\admin\"))
                System.IO.Directory.CreateDirectory(@"C:\admin\");

            if (!System.IO.Directory.Exists(@"C:\admin\ninjafiles\"))
                System.IO.Directory.CreateDirectory(@"C:\admin\ninjafiles\");

            // Need this for the menu item
            iSwApp.SetAddinCallbackInfo(0, this, Cookie);
            iCmdMgr = iSwApp.GetCommandManager(Cookie);
            AddCommandMgr();
            return true;
        }

        private void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            cmdGroup = iCmdMgr.CreateCommandGroup(1, "Ninja", "Efficient Ninja Applications", "", 3);
            cmdGroup.AddCommandItem2("Ninja Report", 0, "Creates a Test Report", "", -1, "_ninjaSoftware", "", 0,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();
        }

        public void _ninjaSoftware()
        {
            values data = new values();
            
            Form topForm = new Form1(iSwApp, data);
            topForm.TopMost = true;
            Application.Run(topForm);
        }

        public bool DisconnectFromSW()
        {
            iCmdMgr.RemoveCommandGroup(1);
            iSwApp = null;
            GC.Collect();

            //if (System.IO.Directory.Exists(path))
            //    System.IO.Directory.Delete(path, true);

            return true;
        }

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

            string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
            Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
            addinkey.SetValue(null, 1);
            addinkey.SetValue("Description", "Reporting a Ninja is a punishable offense");
            addinkey.SetValue("Title", "Ninja Report");

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
