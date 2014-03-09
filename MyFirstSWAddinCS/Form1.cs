using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

using System.Diagnostics;
using System.Threading;
using System.IO;

namespace MyFirstSWAddinCS
{

    public partial class Form1 : Form
    {
        ISldWorks iSwApp;
        String path;
        values data;

        public Form1()
        {
            InitializeComponent();
            Form1_Load(null, null);
        }

        public Form1(ISldWorks SwApp, values data_passed)
        {
            data = data_passed;
            InitializeComponent();
            iSwApp = SwApp;
            path = @"C:\Program Files (x86)\Pattern Ninja\HereLiesStuff\";
            this.FormClosing += Form1_FormClosing;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.IO.DirectoryInfo dir;

            dir = new System.IO.DirectoryInfo(@"c:\TempNinja\p1");
            foreach (System.IO.FileInfo file in dir.GetFiles()) file.Delete();

            dir = new System.IO.DirectoryInfo(@"c:\TempNinja\p2");
            foreach (System.IO.FileInfo file in dir.GetFiles()) file.Delete();

            dir = new System.IO.DirectoryInfo(@"c:\TempNinja\d1");
            foreach (System.IO.FileInfo file in dir.GetFiles()) file.Delete();

            dir = new System.IO.DirectoryInfo(@"c:\TempNinja\d2");
            foreach (System.IO.FileInfo file in dir.GetFiles()) file.Delete();

            dir = new System.IO.DirectoryInfo(@"c:\TempNinja\r1");
            foreach (System.IO.FileInfo file in dir.GetFiles()) file.Delete();

            dir = new System.IO.DirectoryInfo(@"c:\TempNinja\r2");
            foreach (System.IO.FileInfo file in dir.GetFiles()) file.Delete();

            foreach (FileInfo f in new DirectoryInfo(@"c:\TempNinja\").GetFiles("*.txt")) f.Delete();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            date.Text = DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year;

            p1.DataSource = null;
            p1.DataSource = data.p_file1;
            p2.DataSource = null;
            p2.DataSource = data.p_file2;
            d1.DataSource = null;
            d1.DataSource = data.d_file1;
            d2.DataSource = null;
            d2.DataSource = data.d_file2;
            r1.DataSource = null;
            r1.DataSource = data.r_file1;
            r2.DataSource = null;
            r2.DataSource = data.r_file2;
        }

        private void load_config()
        {
            Form1_FormClosing(null, null);

            if (config.Text == "")
            {
                MessageBox.Show("No file chosen");
                return;
            }

            SolidWorks.Interop.sldworks.ModelDoc2 swModelDoc;
            SolidWorks.Interop.sldworks.SelectionMgr selMgr;
            //SolidWorks.Interop.sldworks.Feature swFeature;
            //SolidWorks.Interop.sldworks.Feature name;
            //SolidWorks.Interop.sldworks.Feature copyName;
            //SolidWorks.Interop.sldworks.Feature scopyName;
            //SolidWorks.Interop.sldworks.Feature swSketchFeature;

            swModelDoc = iSwApp.ActiveDoc;
            //if (swModelDoc == null)
            //{
            //    MessageBox.Show("No active document found");
            //    return;
            //}

            selMgr = swModelDoc.SelectionManager;

            //if (selMgr == null)
            //{
            //    MessageBox.Show("An error has occurred. Not Michael's fault");
            //    return;
            //}


            //if (selMgr.GetSelectedObject2(1) == null)
            //{
            //    MessageBox.Show("Nothing is selected");
            //    return;
            //}

            //swFeature = selMgr.GetSelectedObject5(1);
            //swSketchFeature = swFeature.GetSpecificFeature2();

            // Clear Everything
            data.p_ball.Clear();
            data.p_file1.Clear();
            data.p_file2.Clear();
            data.p_letter.Clear();
            data.v_ptol.Clear();

            data.d_ball.Clear();
            data.d_file1.Clear();
            data.d_file2.Clear();
            data.d_letter.Clear();
            data.v_dtol.Clear();

            data.r_ball.Clear();
            data.r_file1.Clear();
            data.r_file2.Clear();
            data.r_letter.Clear();
            data.v_rtol.Clear();
            data.rd.Clear();


            String[] f = File.ReadAllLines(config.Text);
            String[] list;
            foreach (String li in f)
            {
                list = li.Split(';');

                if (list[0] != "i")
                {
                    if (list[0] == "r")
                    {
                        swModelDoc.SelectByName(0, list[1]);
                        if (selMgr.GetSelectedObject2(1) == null)
                        {
                            MessageBox.Show("Cannot find sketch: " + list[1] + "\n\nOther Information:\nActual: " + list[2] + "\nR\\D: " + list[6] + "\nTolerance: " + " " + list[3] + "\nLetter: " + list[4] + "\nBall Rad: " + list[5]);
                            continue;
                        }
                        iSwApp.RunMacro(path + "Macro1.swp", "", "");


                        swModelDoc.SelectByName(0, list[2]);
                        if (selMgr.GetSelectedObject2(1) == null)
                        {
                            MessageBox.Show("Cannot find: " + list[1] + "\n\nOther Information:\nTarget:" + list[2] + "\nR\\D: " + list[6] + "\nTolerance: " + " " + list[3] + "\nLetter: " + list[4] + "\nBall Rad: " + list[5]);
                            continue;
                        }
                        iSwApp.RunMacro(path + "Macro1.swp", "", "");
                    }
                    else
                    {
                        swModelDoc.SelectByName(0, list[1]);
                        if (selMgr.GetSelectedObject2(1) == null)
                        {
                            MessageBox.Show("Cannot find sketch: " + list[1] + "\n\nOther Information:\nActual: " + list[2] + "\nTolerance: " + " " + list[3] + "\nLetter: " + list[4] + "\nBall Rad: " + list[5]);
                            continue;
                        }
                        iSwApp.RunMacro(path + "Macro1.swp", "", "");


                        swModelDoc.SelectByName(0, list[2]);
                        if (selMgr.GetSelectedObject2(1) == null)
                        {
                            MessageBox.Show("Cannot find: " + list[1] + "\n\nOther Information:\nTarget:" + list[2] + "\nTolerance: " + " " + list[3] + "\nLetter: " + list[4] + "\nBall Rad: " + list[5]);
                            continue;
                        }
                        iSwApp.RunMacro(path + "Macro1.swp", "", "");
                    }
                }

                System.Threading.Thread.Sleep(250);
                if (list[0] == "p")
                {
                    System.IO.File.Move(@"C:\TempNinja\" + list[1] + ".txt", @"C:\TempNinja\p1\" + list[1] + ".txt");
                    System.IO.File.Move(@"C:\TempNinja\" + list[2] + ".txt", @"C:\TempNinja\p2\" + list[2] + ".txt");
                    data.p_file1.Add(list[1] + ".txt");
                    data.p_file2.Add(list[2] + ".txt");
                    data.v_ptol.Add(list[3]);
                    data.p_letter.Add(list[4]);
                    data.p_ball.Add(list[5]);
                }
                else if (list[0] == "d")
                {
                    System.IO.File.Move(@"C:\TempNinja\" + list[1] + ".txt", @"C:\TempNinja\d1\" + list[1] + ".txt");
                    System.IO.File.Move(@"C:\TempNinja\" + list[2] + ".txt", @"C:\TempNinja\d2\" + list[2] + ".txt");
                    data.d_file1.Add(list[1] + ".txt");
                    data.d_file2.Add(list[2] + ".txt");
                    data.v_dtol.Add(list[3]);
                    data.d_letter.Add(list[4]);
                    data.d_ball.Add(list[5]);
                }
                else if (list[0] == "r")
                {
                    data.r_file1.Add(list[1] + ".txt");
                    data.r_file2.Add(list[2] + ".txt");
                    System.IO.File.Move(@"C:\TempNinja\" + list[1] + ".txt", @"C:\TempNinja\r1\" + list[1] + ".txt");
                    System.IO.File.Move(@"C:\TempNinja\" + list[2] + ".txt", @"C:\TempNinja\r2\" + list[2] + ".txt");
                    data.v_rtol.Add(list[3]);
                    data.r_letter.Add(list[4]);
                    data.r_ball.Add(list[5]);
                    data.rd.Add(list[6]);
                }
                else if (list[0] == "i")
                {
                    customer.Text = list[1];
                    part.Text = list[2];
                    name.Text = list[3];
                    //date.Text       = list[4];
                    serial.Text = list[5];
                    job.Text = list[6];
                }
            }

            // Update things
            p_ball.DataSource = null;
            d_ball.DataSource = null;
            r_ball.DataSource = null;

            p_ball.DataSource = data.p_ball;
            d_ball.DataSource = data.d_ball;
            r_ball.DataSource = data.r_ball;

            ptol.DataSource = null;
            dtol.DataSource = null;
            rtol.DataSource = null;

            ptol.DataSource = data.v_ptol;
            dtol.DataSource = data.v_dtol;
            rtol.DataSource = data.v_rtol;

            pletter.DataSource = null;
            dletter.DataSource = null;
            rletter.DataSource = null;

            pletter.DataSource = data.p_letter;
            dletter.DataSource = data.d_letter;
            rletter.DataSource = data.r_letter;

            p1.DataSource = null;
            p2.DataSource = null;
            d1.DataSource = null;
            d2.DataSource = null;
            r1.DataSource = null;
            r2.DataSource = null;

            p1.DataSource = data.p_file1;
            p2.DataSource = data.p_file2;
            d1.DataSource = data.d_file1;
            d2.DataSource = data.d_file2;
            r1.DataSource = data.r_file1;
            r2.DataSource = data.r_file2;

            d_r_cb.DataSource = null;
            d_r_cb.DataSource = data.rd;

            return;
        }


        private void generate_Click(object sender, EventArgs e)
        {
            if (data.p_file1.Count == 0 && data.d_file1.Count == 0 && data.r_file1.Count == 0)
            {
                MessageBox.Show("No files selected");
                return;
            }
            if (data.p_letter.Count == 0 && data.p_file1.Count > 0)
                data.p_letter.Add("");
            if (data.d_letter.Count == 0 && data.d_file1.Count > 0)
                data.d_letter.Add("");
            if (data.r_letter.Count == 0 && data.r_file1.Count > 0)
                data.r_letter.Add("");

            if (data.p_file1.Count != data.p_file2.Count)
            {
                MessageBox.Show("Please enter the same number of Point files in Target and Actual");
                return;
            }
            if (data.d_file1.Count != data.d_file2.Count)
            {
                MessageBox.Show("Please enter the same number of Lineal files in Target and Actual");
                return;
            }
            if (data.r_file1.Count != data.r_file2.Count)
            {
                MessageBox.Show("Please enter the same number of Radial files in Target and Actual");
                return;
            }

            if (data.v_ptol.Count == 0 && data.p_file1.Count > 0)
            {
                MessageBox.Show("Please enter Point Dimension Tolerances");
                return;
            }
            if (data.v_dtol.Count == 0 && data.d_file1.Count > 0)
            {
                MessageBox.Show("Please enter Lineal Dimension Tolerances");
                return;
            }
            if (data.v_rtol.Count == 0 && data.r_file1.Count > 0)
            {
                MessageBox.Show("Please enter Radial Dimension Tolerances");
                return;
            }

            if (data.p_ball.Count == 0 && data.p_file1.Count > 0)
            {
                MessageBox.Show("Please enter Point ball dimensions");
                return;
            }
            if (data.d_ball.Count == 0 && data.d_file1.Count > 0)
            {
                MessageBox.Show("Please enter Lineal ball dimensions");
                return;
            }
            if (data.r_ball.Count == 0 && data.r_file1.Count > 0)
            {
                MessageBox.Show("Please enter Radial ball dimensions");
                return;
            }

            if (data.rd.Count == 0 && data.r_file1.Count > 0)
            {
                MessageBox.Show("Please enter 'R' or 'D' for radius or diameter");
                return;
            }


            data.customer = customer.Text;
            data.serial = serial.Text;
            data.part = part.Text;
            data.name = name.Text;
            data.job = job.Text;
            data.date = date.Text;

            // Auto fill-in
            while (data.v_ptol.Count < data.p_file1.Count)
                data.v_ptol.Add(data.v_ptol[data.v_ptol.Count - 1]);
            while (data.v_dtol.Count < data.d_file1.Count)
                data.v_dtol.Add(data.v_dtol[data.v_dtol.Count - 1]);
            while (data.v_rtol.Count < data.r_file1.Count)
                data.v_rtol.Add(data.v_rtol[data.v_rtol.Count - 1]);
            while (data.rd.Count < data.r_file1.Count)
                data.rd.Add(data.rd[data.rd.Count - 1]);

            ptol.DataSource = null;
            dtol.DataSource = null;
            rtol.DataSource = null;

            ptol.DataSource = data.v_ptol;
            dtol.DataSource = data.v_dtol;
            rtol.DataSource = data.v_rtol;

            while (data.p_ball.Count < data.p_file1.Count)
                data.p_ball.Add(data.p_ball[data.p_ball.Count - 1]);
            while (data.d_ball.Count < data.d_file1.Count)
                data.d_ball.Add(data.d_ball[data.d_ball.Count - 1]);
            while (data.r_ball.Count < data.r_file1.Count)
                data.r_ball.Add(data.r_ball[data.r_ball.Count - 1]);

            p_ball.DataSource = null;
            d_ball.DataSource = null;
            r_ball.DataSource = null;

            p_ball.DataSource = data.p_ball;
            d_ball.DataSource = data.d_ball;
            r_ball.DataSource = data.r_ball;

            while (data.p_letter.Count < data.p_file1.Count)
                data.p_letter.Add(data.p_letter[data.p_letter.Count - 1]);
            while (data.d_letter.Count < data.d_file1.Count)
                data.d_letter.Add(data.d_letter[data.d_letter.Count - 1]);
            while (data.r_letter.Count < data.r_file1.Count)
                data.r_letter.Add(data.r_letter[data.r_letter.Count - 1]);

            pletter.DataSource = null;
            dletter.DataSource = null;
            rletter.DataSource = null;

            pletter.DataSource = data.p_letter;
            dletter.DataSource = data.d_letter;
            rletter.DataSource = data.r_letter;

            if (part.Text == "" || serial.Text == "")
            {
                MessageBox.Show("Part Number and Serial Number must be entered");
                return;
            }

            // Save Config
            String lines = "";
            // P files
            lines = "i;" + data.customer + ";" + data.part + ";" + data.name + ";" + data.date + ";" + data.serial + ";" + data.job + System.Environment.NewLine;
            for (int i = 0; i < data.p_file1.Count; i++)
            {
                lines = lines + "p;" + data.p_file1[i].Substring(0, data.p_file1[i].Length - 4) + ";" + data.p_file2[i].Substring(0, data.p_file2[i].Length - 4) + ";" + data.v_ptol[i] + ";" + data.p_letter[i] + ";" + data.p_ball[i] + System.Environment.NewLine;
            }
            // D files
            for (int i = 0; i < data.d_file1.Count; i++)
            {
                lines = lines + "d;" + data.d_file1[i].Substring(0, data.d_file1[i].Length - 4) + ";" + data.d_file2[i].Substring(0, data.d_file2[i].Length - 4) + ";" + data.v_dtol[i] + ";" + data.d_letter[i] + ";" + data.d_ball[i] + System.Environment.NewLine;
            }
            // R files
            for (int i = 0; i < data.r_file1.Count; i++)
            {
                lines = lines + "r;" + data.r_file1[i].Substring(0, data.r_file1[i].Length - 4) + ";" + data.r_file2[i].Substring(0, data.r_file2[i].Length - 4) + ";" + data.v_rtol[i] + ";" + data.r_letter[i] + ";" + data.r_ball[i] + ";" + data.rd[i] + System.Environment.NewLine;
            }
            //lines = lines + "\b";

            

            if( File.Exists(savepath.Text) )
            {
                DialogResult dialogResult = MessageBox.Show("Do you want to overwrite existing file?\n"+savepath.Text, "Overwrite File?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                    return;
            }

            if (savepath.Text != "" && config.Text == "")
            {
                System.IO.File.WriteAllText(savepath.Text, lines);
            }
            else if (savepath.Text != "" && config.Text != "")
                System.IO.File.WriteAllText(savepath.Text, lines);
            else
                System.IO.File.WriteAllText(@"C:\admin\ninjafiles\" + serial.Text + "_" + part.Text + ".config", lines);

            // Run Report
            ReportGen.report rep = new ReportGen.report();
            rep.generate(data);
        }


        ListBox lb;
        List<String> ls;
        String d;
        private void add_to_box(object sender)
        {
            Button button = sender as Button;

            iSwApp.RunMacro(path + "Macro1.swp", "", "");

            var directory = new System.IO.DirectoryInfo("c:\\TempNinja");

            switch (button.Name)
            {
                case "a_p1":
                    lb = p1;
                    ls = data.p_file1;
                    d = "c:\\TempNinja\\p1";
                    break;
                case "a_p2":
                    lb = p2;
                    ls = data.p_file2;
                    d = "c:\\TempNinja\\p2";
                    break;
                case "a_d1":
                    lb = d1;
                    ls = data.d_file1;
                    d = "c:\\TempNinja\\d1";
                    break;
                case "a_d2":
                    lb = d2;
                    ls = data.d_file2;
                    d = "c:\\TempNinja\\d2";
                    break;
                case "a_r1":
                    lb = r1;
                    ls = data.r_file1;
                    d = "c:\\TempNinja\\r1";
                    break;
                case "a_r2":
                    lb = r2;
                    ls = data.r_file2;
                    d = "c:\\TempNinja\\r2";
                    break;
            }

            var file = directory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();
            try
            {
                System.IO.File.Move(@"C:\TempNinja\" + file.Name, d + "\\" + file.Name);
                var directory2 = new System.IO.DirectoryInfo(d);

                file = directory2.GetFiles().OrderByDescending(f => f.LastWriteTime).First();

                // Only add the file name if it doesnt already exist
                if (ls.Count == 0)
                    ls.Add(file.Name);

                else if (ls.Last() != file.Name)
                    ls.Add(file.Name);

                lb.DataSource = null;
                lb.DataSource = ls;

                // Some automated form filling in
                if (button.Name == "a_p1")
                {
                    if (data.p_file1.Count > 2)
                    {
                        if (data.v_ptol.Count == data.p_file1.Count - 2)
                            data.v_ptol.Add(data.v_ptol[data.v_ptol.Count - 1]);
                        if (data.p_letter.Count == data.p_file1.Count - 2)
                            data.p_letter.Add(data.p_letter[data.p_letter.Count - 1]);
                        if (data.p_ball.Count == data.p_file1.Count - 2)
                            data.p_ball.Add(data.p_ball[data.p_ball.Count - 1]);

                        p_ball.DataSource = null;
                        pletter.DataSource = null;
                        ptol.DataSource = null;

                        p_ball.DataSource = data.p_ball;
                        pletter.DataSource = data.p_letter;
                        ptol.DataSource = data.v_ptol;
                    }
                }
                if (button.Name == "a_d1")
                {
                    if (data.d_file1.Count > 2)
                    {
                        if (data.v_dtol.Count == data.d_file1.Count - 2)
                            data.v_dtol.Add(data.v_dtol[data.v_dtol.Count - 1]);
                        if (data.d_letter.Count == data.d_file1.Count - 2)
                            data.d_letter.Add(data.d_letter[data.d_letter.Count - 1]);
                        if (data.d_ball.Count == data.d_file1.Count - 2)
                            data.d_ball.Add(data.d_ball[data.d_ball.Count - 1]);

                        d_ball.DataSource = null;
                        dletter.DataSource = null;
                        dtol.DataSource = null;

                        d_ball.DataSource = data.d_ball;
                        dletter.DataSource = data.d_letter;
                        dtol.DataSource = data.v_dtol;
                    }
                }
                if (button.Name == "a_r1")
                {
                    if (data.r_file1.Count > 2)
                    {
                        if (data.v_rtol.Count == data.r_file1.Count - 2)
                            data.v_rtol.Add(data.v_rtol[data.v_rtol.Count - 1]);
                        if (data.r_letter.Count == data.r_file1.Count - 2)
                            data.r_letter.Add(data.r_letter[data.r_letter.Count - 1]);
                        if (data.r_ball.Count == data.r_file1.Count - 2)
                            data.r_ball.Add(data.r_ball[data.r_ball.Count - 1]);
                        if (data.rd.Count == data.r_file1.Count - 2)
                            data.rd.Add(data.rd[data.rd.Count - 1]);

                        r_ball.DataSource = null;
                        rletter.DataSource = null;
                        rtol.DataSource = null;

                        r_ball.DataSource = data.r_ball;
                        rletter.DataSource = data.r_letter;
                        rtol.DataSource = data.v_rtol;
                    }
                }
            }
            catch { }
        }

        void remove_from_box(object sender)
        {
            Button button = sender as Button;

            switch (button.Name)
            {
                case "x_p1":
                    lb = p1;
                    ls = data.p_file1;
                    d = "c:\\TempNinja\\p1";
                    break;
                case "x_p2":
                    lb = p2;
                    ls = data.p_file2;
                    d = "c:\\TempNinja\\p2";
                    break;
                case "x_d1":
                    lb = d1;
                    ls = data.d_file1;
                    d = "c:\\TempNinja\\d1";
                    break;
                case "x_d2":
                    lb = d2;
                    ls = data.d_file2;
                    d = "c:\\TempNinja\\d2";
                    break;
                case "x_r1":
                    lb = r1;
                    ls = data.r_file1;
                    d = "c:\\TempNinja\\r1";
                    break;
                case "x_r2":
                    lb = r2;
                    ls = data.r_file2;
                    d = "c:\\TempNinja\\r2";
                    break;
            }

            int selectedIndex = lb.SelectedIndex;
            try
            {
                // Remove the item in the List.
                System.IO.FileInfo fi = new System.IO.FileInfo(d + "\\" + ls[selectedIndex]);
                try
                {
                    fi.Delete();
                }
                catch
                {
                }
                ls.RemoveAt(selectedIndex);

            }
            catch
            {
            }

            lb.DataSource = null;
            lb.DataSource = ls;
        }

        private void add_tol(object sender)
        {
            Button button = sender as Button;
            List<String> f = data.v_ptol;
            TextBox b = v_ptol;
            ListBox lb = ptol;

            double t;
            String[] t_tol;
            switch (button.Name)
            {
                case "a_ptol":
                    f = data.v_ptol;
                    b = v_ptol;
                    lb = ptol;
                    t_tol = b.Text.Split(' ');
                    if (b.Text.Split(' ').Length != 2)
                    {
                        MessageBox.Show("Tolerance Format: + space -. Example: .002 -.300");
                        return;
                    }
                    if (!double.TryParse(t_tol[0], out t) || !double.TryParse(t_tol[1], out t))
                    {
                        MessageBox.Show("Tolerances must be numbers");
                        return;
                    }
                    break;
                case "a_dtol":
                    f = data.v_dtol;
                    b = v_dtol;
                    lb = dtol;
                    t_tol = b.Text.Split(' ');
                    if (b.Text.Split(' ').Length != 2)
                    {
                        MessageBox.Show("Tolerance Format: + space -. Example: .002 -.300");
                        return;
                    }
                    if (!double.TryParse(t_tol[0], out t) || !double.TryParse(t_tol[1], out t))
                    {
                        MessageBox.Show("Tolerances must be numbers");
                        return;
                    }
                    break;
                case "a_rtol":
                    f = data.v_rtol;
                    b = v_rtol;
                    lb = rtol;
                    t_tol = b.Text.Split(' ');
                    if (b.Text.Split(' ').Length != 2)
                    {
                        MessageBox.Show("Tolerance Format: + space -. Example: .002 -.300");
                        return;
                    }
                    if (!double.TryParse(t_tol[0], out t) || !double.TryParse(t_tol[1], out t))
                    {
                        MessageBox.Show("Tolerances must be numbers");
                        return;
                    }
                    break;
                case "a_pball":
                    f = data.p_ball;
                    b = p_ballText;
                    lb = p_ball;
                    if (!double.TryParse(b.Text, out t))
                    {
                        MessageBox.Show("Numbers only for ball diameter please");
                        return;
                    }
                    break;
                case "a_dball":
                    f = data.d_ball;
                    b = d_ballText;
                    lb = d_ball;
                    if (!double.TryParse(b.Text, out t))
                    {
                        MessageBox.Show("Numbers only for ball diameter please");
                        return;
                    }
                    break;
                case "a_rball":
                    f = data.r_ball;
                    b = r_ballText;
                    lb = r_ball;
                    if (!double.TryParse(b.Text, out t))
                    {
                        MessageBox.Show("Numbers only for ball diameter please");
                        return;
                    }
                    break;
                case "a_pletter":
                    f = data.p_letter;
                    b = p_letter;
                    lb = pletter;
                    break;
                case "a_dletter":
                    f = data.d_letter;
                    b = d_letter;
                    lb = dletter;
                    break;
                case "a_rletter":
                    f = data.r_letter;
                    b = r_letter;
                    lb = rletter;
                    break;
                case "a_rd":
                    f = data.rd;
                    b = rd;
                    lb = d_r_cb;
                    if (rd.Text.ToUpper() != "R" && rd.Text.ToUpper() != "D")
                    {
                        MessageBox.Show("Please enter an R or D for Radius or Diameter");
                        return;
                    }
                    b.Text = b.Text.ToUpper();
                    break;
            }
            f.Add(b.Text);
            b.Text = "";
            lb.DataSource = null;
            lb.DataSource = f;
        }

        private void x_tol(object sender)
        {
            Button button = sender as Button;
            List<String> f = data.v_ptol;
            TextBox b = v_ptol;
            ListBox lb = ptol;
            switch (button.Name)
            {
                case "x_ptol":
                    f = data.v_ptol;
                    b = v_ptol;
                    lb = ptol;
                    break;
                case "x_dtol":
                    f = data.v_dtol;
                    b = v_dtol;
                    lb = dtol;
                    break;
                case "x_rtol":
                    f = data.v_rtol;
                    b = v_rtol;
                    lb = rtol;
                    break;
                case "x_pletter":
                    f = data.p_letter;
                    b = p_letter;
                    lb = pletter;
                    break;
                case "x_dletter":
                    f = data.d_letter;
                    b = d_letter;
                    lb = dletter;
                    break;
                case "x_rletter":
                    f = data.r_letter;
                    b = r_letter;
                    lb = rletter;
                    break;
                case "x_pball":
                    f = data.p_ball;
                    b = p_ballText;
                    lb = p_ball;
                    break;
                case "x_dball":
                    f = data.d_ball;
                    b = d_ballText;
                    lb = d_ball;
                    break;
                case "x_rball":
                    f = data.r_ball;
                    b = r_ballText;
                    lb = r_ball;
                    break;
                case "x_rd":
                    f = data.rd;
                    b = rd;
                    lb = d_r_cb;
                    break;
            }
            int index = lb.SelectedIndex;
            f.RemoveAt(index);
            lb.DataSource = null;
            lb.DataSource = f;
        }

        private void x_p1_Click(object sender, EventArgs e)
        {
            remove_from_box(sender);
        }

        private void x_p2_Click(object sender, EventArgs e)
        {
            remove_from_box(sender);
        }

        private void x_d1_Click(object sender, EventArgs e)
        {
            remove_from_box(sender);
        }

        private void x_d2_Click(object sender, EventArgs e)
        {
            remove_from_box(sender);
        }

        private void x_r1_Click(object sender, EventArgs e)
        {
            remove_from_box(sender);
        }

        private void x_r2_Click(object sender, EventArgs e)
        {
            remove_from_box(sender);
        }

        private void a_p1_Click(object sender, EventArgs e)
        {
            add_to_box(sender);
        }

        private void a_p2_Click(object sender, EventArgs e)
        {
            add_to_box(sender);
        }

        private void a_d1_Click(object sender, EventArgs e)
        {
            add_to_box(sender);
        }

        private void a_d2_Click(object sender, EventArgs e)
        {
            add_to_box(sender);
        }

        private void a_r1_Click(object sender, EventArgs e)
        {
            add_to_box(sender);
        }

        private void a_r2_Click(object sender, EventArgs e)
        {
            add_to_box(sender);
        }

        private void a_ptol_Click(object sender, EventArgs e)
        {
            add_tol(sender);
        }

        private void a_dtol_Click(object sender, EventArgs e)
        {
            add_tol(sender);
        }

        private void a_rtol_Click(object sender, EventArgs e)
        {
            add_tol(sender);
        }

        private void x_ptol_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void x_dtol_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void x_rtol_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void a_pletter_Click(object sender, EventArgs e)
        {
            add_tol(sender);
        }

        private void a_dletter_Click(object sender, EventArgs e)
        {
            add_tol(sender);

        }

        private void a_rletter_Click(object sender, EventArgs e)
        {
            add_tol(sender);

        }

        private void a_pball_Click(object sender, EventArgs e)
        {
            add_tol(sender);
        }

        private void a_dball_Click(object sender, EventArgs e)
        {
            add_tol(sender);
        }

        private void a_rball_Click(object sender, EventArgs e)
        {
            add_tol(sender);
        }

        private void x_pletter_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void x_dletter_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void x_rletter_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void x_pball_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void x_dball_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void x_rball_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void x_rd_Click(object sender, EventArgs e)
        {
            x_tol(sender);
        }

        private void a_rd_Click(object sender, EventArgs e)
        {
            add_tol(sender);
        }

        private void configload_Click(object sender, EventArgs e)
        {
            load_config();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var FD = new System.Windows.Forms.OpenFileDialog();
            FD.InitialDirectory = @"C:\admin\ninjafiles\";
            if (FD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                config.Text = FD.FileName;
                savepath.Text = FD.FileName;
            }
        }

        private void savepath_b_Click(object sender, EventArgs e)
        {
            //var FD = new System.Windows.Forms.OpenFileDialog();
            //var FD = new System.Windows.Forms.FolderBrowserDialog();
            //FD.SelectedPath = @"C:\admin\ninjafiles\";
            //DialogResult result = FD.ShowDialog();

            var FD = new System.Windows.Forms.OpenFileDialog();
            FD.InitialDirectory = @"C:\admin\ninjafiles\";
            FD.ValidateNames = false;
            FD.CheckFileExists = false;
            FD.CheckPathExists = true;
            FD.FileName = serial.Text+"_"+part.Text;
            if (FD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //FD.FileName = serial.Text + "_" + part.Text;
                //if (result == DialogResult.OK)
                savepath.Text = FD.FileName+".config";
            }
        }

    }
}
