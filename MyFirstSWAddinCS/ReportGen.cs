using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ReportGen
{
    public class report
    {
        public report()
        {
        }

        public void tacobell()
        {
        }

        public double distance(String point1, String point2)
        {
            string[] p1_s = point1.Split(',');
            string[] p2_s = point2.Split(',');

            double[] p1 = { Convert.ToDouble(p1_s[0]), 
                            Convert.ToDouble(p1_s[1]), 
                            Convert.ToDouble(p1_s[2]) };

            double[] p2 = { Convert.ToDouble(p2_s[0]), 
                            Convert.ToDouble(p2_s[1]), 
                            Convert.ToDouble(p2_s[2]) };

            double sqrt = Math.Sqrt( Math.Pow( p1[0] - p2[0], 2 ) +
                                     Math.Pow(p1[1] - p2[1], 2)+
                                     Math.Pow(p1[2] - p2[2], 2) );

            return sqrt;
        }

        public void generate(MyFirstSWAddinCS.values data)
        {
            int p_start = 0;
            int d_start = 0;
            int r_start = 0;

            Excel.Application ex;
            Excel.Workbook wb;
            Excel.Worksheet ws;
            Excel.Worksheet ws2;
            Excel.Worksheet ws3;
            ex = new Excel.Application();

            String path = "C:\\Program Files (x86)\\Pattern Ninja\\HereLiesStuff\\";
            String userPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            //System.IO.File.Copy(path + "CMMReportTemplate.xlsx", userPath + "\\CMMReport1.xlsx", true);


            wb = ex.Workbooks.Add();
            //  wb = ex.Workbooks.Open(@"c:\Users\Michael\Documents\CMMReport1.xlsx");
            //wb = ex.Workbooks.Open(userPath + "\\CMMReport1.xlsx");


            // //Excel.Worksheet ws = wb.Sheets.Add();
            ws = wb.Sheets.get_Item(1);
            ws2 = wb.Sheets.Add();
            ws3 = wb.Sheets.Add();

            ws.Name = "Points";
            ws2.Name = "Lineal";
            ws3.Name = "Radial";

            int line_no = 1;

            // POINT DIMENSIONS
            ws.Cells[line_no, "A"].Value2 = "Points";
            ws.Cells[line_no, "B"].Value2 = "X";
            ws.Cells[line_no, "C"].Value2 = "Y";
            ws.Cells[line_no, "D"].Value2 = "Z";
            ws.Cells[line_no, "E"].Value2 = "Deviation";
            ws.Cells[line_no, "F"].Value2 = "+Tol";
            ws.Cells[line_no, "G"].Value2 = "-Tol";
            ws.Cells[line_no, "H"].Value2 = "Exceeds";
            p_start = line_no;
            int point_no = 1;
            line_no++;
            for( int i = 0; i < data.p_file1.Count; i++ )
            {
                int f_start = line_no;

                //if( i == 0 )
                //    point_no = 1;
                //else if (data.p_letter[i - 1] != data.p_letter[i])
                //    point_no = 1;

                // Read in target file dimensions
                String fname = data.p_file1[i];
                string[] f = File.ReadAllLines("c:\\TempNinja\\p1\\" + fname);
                foreach (var fl in f)
                {
                    string[] xyz = fl.Split(',');
                    for (int j = 0; j < 3; j++)
                    {
                        ws.Cells[line_no, "A"].Value2 = point_no+data.p_letter[i];
                        ws.Cells[line_no, "B"].Value2 = xyz[0];
                        ws.Cells[line_no, "C"].Value2 = xyz[1];
                        ws.Cells[line_no, "D"].Value2 = xyz[2];
                        data.target_vals.Add(fl);
                    }
                    line_no++;
                    point_no++;
                }

                // Read in Deviation file dimensions
                fname = data.p_file2[i];
                f = File.ReadAllLines("c:\\TempNinja\\p2\\" + fname);
                foreach (var fl in f)
                    data.deviat_vals.Add(fl);

                //// Compare the distances to make sure theyre the correct points
                int temp_num = f_start;
                foreach (String line in data.target_vals)
                {
                    String[] t_xyz = line.Split(',');
                    if (data.deviat_vals.Count > 0)
                    {
                        int found = 0;
                        for (int k = 0; k < data.deviat_vals.Count; k++)
                        {
                            String[] d_xyz = data.deviat_vals[k].Split(',');

                            ws.Cells[temp_num, "J"].Value2 = d_xyz[0];
                            ws.Cells[temp_num, "K"].Value2 = d_xyz[1];
                            ws.Cells[temp_num, "L"].Value2 = d_xyz[2];
                            ws.Cells[temp_num, "M"].Value2 = "= SQRT( (B" + temp_num + "-J" + temp_num + ")^2 +(C" + temp_num + "-K" + temp_num + ")^2+(D" + temp_num + "-L" + temp_num + ")^2)";
                           
                            if (ws.Cells[temp_num, "M"].Value2 < Convert.ToDouble(data.p_ball[i]) + .05)
                            {
                                found = 1;
                                ws.Cells[temp_num, "N"].Value2 = data.p_ball[i];
                                ws.Cells[temp_num, "E"].Value2 = "=M" + temp_num + "-N" + temp_num;
                                Excel.Range rangeD = ws.get_Range("E"+temp_num);
                                rangeD.NumberFormat = "0.0000";

                               string[] tols = data.v_ptol[i].Split(' ');

                               ws.Cells[temp_num, "F"].Value2 = tols[0];
                               ws.Cells[temp_num, "G"].Value2 = tols[1];
                                int t = temp_num;
                                ws.Cells[temp_num, "H"].Value2 = "=IF(E"+t+">F"+t+", E"+t+"-F"+t+", IF(E"+t+"<G"+t+",E"+t+"-G"+t+",\"\"))";
                                data.deviat_vals.RemoveAt(k);
                                break;
                            }
                        }
                        if (found == 0)
                        {
                            ws.Cells[temp_num, "J"].Value2 = "ERROR";
                            ws.Cells[temp_num, "K"].Value2 = "ERROR";
                            ws.Cells[temp_num, "L"].Value2 = "ERROR";
                        }
                    }
                    temp_num++;
                }
                data.deviat_vals.Clear();
                data.target_vals.Clear();
            }
            int p_end = line_no;

            line_no+=2;

            // DISTANCE DIMENSIONS
            // Assume everything is point or line format. 2 Points
            ws2.Cells[line_no, "B"].Value2 = "Point";
            ws2.Cells[line_no, "C"].Value2 = "Target";
            ws2.Cells[line_no, "D"].Value2 = "Actual";
            ws2.Cells[line_no, "E"].Value2 = "Deviation";
            ws2.Cells[line_no, "F"].Value2 = "+Tol";
            ws2.Cells[line_no, "G"].Value2 = "-Tol";
            ws2.Cells[line_no, "H"].Value2 = "Exceeds";

            d_start = line_no;
            line_no++;
            point_no = 1;

            for (int i = 0; i < data.d_file1.Count; i++)
            {
                //if (i == 0)
                //    point_no = 1;
                //else if (data.d_letter[i - 1] != data.d_letter[i])
                //    point_no = 1;

                List<String> targets = new List<String>();
                List<double> targets_d = new List<double>();
                List<String> deviate = new List<String>();
                List<double> deviate_d = new List<double>();

                int f_start = line_no;
                String fname = data.d_file1[i];
                string[] f = File.ReadAllLines("c:\\TempNinja\\d1\\" + fname);
                foreach (var fl in f)
                    targets.Add(fl);

                for (int j = 0; j < targets.Count; j += 2)
                {
                    double d = distance(targets[j], targets[j + 1]);
                    ws2.Cells[line_no, "B"].Value2 = point_no + data.d_letter[i];
                    ws2.Cells[line_no, "C"].Value2 = d;
                    targets_d.Add( d );
                    line_no++;
                    point_no++;
                }

                // Start check the second file
                int temp = f_start;
                fname = data.d_file2[i];
                f = File.ReadAllLines("c:\\TempNinja\\d2\\" + fname);
                foreach (var fl in f)
                    deviate.Add(fl);

                
                for (int j = 0; j < deviate.Count; j += 2)
                {
                    double d = distance(deviate[j], deviate[j + 1]);
                    deviate_d.Add(d);
                }
                // Compare points
                for (int j = 0; j < targets.Count; j += 2)
                {
                    for (int k = 0; k < deviate.Count; k += 2)
                    {
                        if (distance(targets[j], deviate[k]) < Convert.ToDouble(data.d_ball[i]) + .05)
                        {
                            if (distance(targets[j + 1], deviate[k + 1]) < Convert.ToDouble(data.d_ball[i]) + .05)
                            {
                                ws2.Cells[temp, "D"].Value2 = deviate_d[k / 2];
                                ws2.Cells[temp, "E"].Value2 = "=D"+temp+"-C"+temp+" - "+(2*Convert.ToDouble(data.d_ball[i]));

                                string[] tols = data.v_dtol[i].Split(' ');

                                ws2.Cells[temp, "F"].Value2 = tols[0];
                                ws2.Cells[temp, "G"].Value2 = tols[1];

                                int t = temp;
                                ws2.Cells[temp, "H"].Value2 = "=IF(E" + t + ">F" + t + ", E" + t + "-F" + t + ", IF(E" + t + "<G" + t + ",E" + t + "-G" + t + ",\"\"))";

                                break;
                            }
                        }
                        else if (distance(targets[j], deviate[k+1]) < Convert.ToDouble(data.d_ball[i]) + .05)
                        {
                            if (distance(targets[j + 1], deviate[k]) < Convert.ToDouble(data.d_ball[i]) + .05)
                            {
                                ws2.Cells[temp, "D"].Value2 = deviate_d[k / 2];
                                ws2.Cells[temp, "E"].Value2 = "=D" + temp + "-C" + temp + " - " + (2 * Convert.ToDouble(data.d_ball[i]));

                                string[] tols = data.v_dtol[i].Split(' ');

                                ws2.Cells[temp, "F"].Value2 = tols[0];
                                ws2.Cells[temp, "G"].Value2 = tols[1];

                                int t = temp;
                                ws2.Cells[temp, "H"].Value2 = "=IF(E" + t + ">F" + t + ", E" + t + "-F" + t + ", IF(E" + t + "<G" + t + ",E" + t + "-G" + t + ",\"\"))";

                                break;
                            }
                        }
                    }
                    if ( Convert.ToString( ws2.Cells[temp, "C"].Value2) == null)
                        ws2.Cells[temp, "C"].Value2 = "ERROR";
                    temp++;
                }
            }
            int d_end = line_no;

            line_no++;
            line_no++;

            // RADIAL DIMENSIONS
            ws3.Cells[line_no, "B"].Value2 = "Point";
            ws3.Cells[line_no, "C"].Value2 = "Target";
            ws3.Cells[line_no, "D"].Value2 = "Actual";
            ws3.Cells[line_no, "E"].Value2 = "Deviation";
            ws3.Cells[line_no, "F"].Value2 = "+Tol";
            ws3.Cells[line_no, "G"].Value2 = "-Tol";
            ws3.Cells[line_no, "H"].Value2 = "Exceeds";

            r_start = line_no;
            line_no++;
            for (int i = 0; i < data.r_file1.Count; i++)
            {
                if (i == 0)
                    point_no = 1;
                else if (data.r_letter[i - 1] != data.r_letter[i])
                    point_no = 1;

                List<String> targets = new List<String>();
                List<double> targets_d = new List<double>();
                List<String> deviate = new List<String>();
                List<double> deviate_d = new List<double>();

                int f_start = line_no;
                String fname = data.r_file1[i];
                string[] f = File.ReadAllLines("c:\\TempNinja\\r1\\" + fname);
                foreach (var fl in f)
                    targets.Add(fl);

                for (int j = 0; j < targets.Count; j += 2)
                {
                    double d = distance(targets[j], targets[j + 1]);
                    ws3.Cells[line_no, "B"].Value2 = point_no + data.r_letter[i];

                    if (data.rd[i] == "D")
                        d = 2 * d;

                    ws3.Cells[line_no, "C"].Value2 = data.rd[i] + " " + Math.Round(d, 4, MidpointRounding.AwayFromZero).ToString("0.0000");

                    targets_d.Add(d);
                    line_no++;
                    point_no++;
                }

                // Start check the second file
                int temp = f_start;
                fname = data.r_file2[i];
                f = File.ReadAllLines("c:\\TempNinja\\r2\\" + fname);
                foreach (var fl in f)
                    deviate.Add(fl);

                for (int j = 0; j < deviate.Count; j += 2)
                {
                    double d = distance(deviate[j], deviate[j + 1]);

                    if (data.rd[i] == "D")
                        d = 2 * d;

                    deviate_d.Add(d);
                }
                // Compare points
                for (int j = 0; j < targets.Count; j += 2)
                {
                    for (int k = 0; k < deviate.Count; k += 2)
                    {
                        if (distance(targets[j], deviate[k]) < Convert.ToDouble(data.r_ball[i]) + .05)
                        {
                            if (distance(targets[j + 1], deviate[k + 1]) < Convert.ToDouble(data.r_ball[i]) + .05 )
                            {
                                ws3.Cells[temp, "D"].Value2 = data.rd[i] + " " + Math.Round(Convert.ToDouble(deviate_d[k / 2]), 4, MidpointRounding.AwayFromZero).ToString("0.0000");

                                if( data.rd[i] == "D" )
                                    ws3.Cells[temp, "E"].Value2 = "=" + deviate_d[k / 2] + " - " + targets_d[j / 2] + " - " + (2 * Convert.ToDouble(data.r_ball[i]));
                                else
                                    ws3.Cells[temp, "E"].Value2 = "=" + deviate_d[k / 2] + " - " + targets_d[j / 2] + " - " + Convert.ToDouble(data.r_ball[i]);

                                string[] tols = data.v_rtol[i].Split(' ');

                                ws3.Cells[temp, "F"].Value2 = tols[0];
                                ws3.Cells[temp, "G"].Value2 = tols[1];

                                int t = temp;
                                ws3.Cells[temp, "H"].Value2 = "=IF(E" + t + ">F" + t + ", E" + t + "-F" + t + ", IF(E" + t + "<G" + t + ",E" + t + "-G" + t + ",\"\"))";

                                break;
                            }
                        }
                        if (distance(targets[j], deviate[k+1]) < Convert.ToDouble(data.r_ball[i]) + .05)
                        {
                            if (distance(targets[j + 1], deviate[k]) < Convert.ToDouble(data.r_ball[i]) + .05)
                            {
                                ws3.Cells[temp, "D"].Value2 = data.rd[i] + " " + Math.Round(Convert.ToDouble(deviate_d[k / 2]), 4, MidpointRounding.AwayFromZero).ToString("0.0000");

                                if (data.rd[i] == "D")
                                    ws3.Cells[temp, "E"].Value2 = "=" + deviate_d[k / 2] + " - " + targets_d[j / 2] + " - " + (2 * Convert.ToDouble(data.r_ball[i]));
                                else
                                    ws3.Cells[temp, "E"].Value2 = "=" + deviate_d[k / 2] + " - " + targets_d[j / 2] + " - " + Convert.ToDouble(data.r_ball[i]);

                                string[] tols = data.v_rtol[i].Split(' ');

                                ws3.Cells[temp, "F"].Value2 = tols[0];
                                ws3.Cells[temp, "G"].Value2 = tols[1];

                                int t = temp;
                                ws3.Cells[temp, "H"].Value2 = "=IF(E" + t + ">F" + t + ", E" + t + "-F" + t + ", IF(E" + t + "<G" + t + ",E" + t + "-G" + t + ",\"\"))";

                                break;
                            }
                        }
                    }
                    if (Convert.ToString(ws3.Cells[temp, "C"].Value2) == null)
                        ws3.Cells[temp, "C"].Value2 = "ERROR";
                    temp++;
                }
            }
            int r_end = line_no;
            
            //ex.Visible = true;
            //wb.Close(true);




            // MICROSOFT WORD TIME
            object oMissing = System.Reflection.Missing.Value;
            //object oEoF = "\\endofdoc"; /* \endofdoc is a predefined bookmark */ 
            object oEoF = Word.WdUnits.wdStory;

            Word._Application oWord;
            Word._Document oDoc;

            oWord = new Word.Application();
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            //////////////////////////////////////////////////////////////////////
            //oDoc.PageSetup.DifferentFirstPageHeaderFooter = -1;
            //oDoc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "First Page Header";
            //oDoc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //oDoc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "Part Number: " + data.part + "  Serial Number: " + data.serial;

            Microsoft.Office.Interop.Word.Selection s = oWord.Selection;

            // code for the page numbers
            // move selection to page footer (Use wdSeekCurrentPageHeader for header)
            oWord.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter;
            // Align right
            s.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            // start typing
            oWord.ActiveWindow.Selection.TypeText("\tPart Number: " + data.part + "  Serial Number: " + data.serial + "\tPage ");
            // create the field  for current page number
            object CurrentPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;
            // insert that field into the selection
            oWord.ActiveWindow.Selection.Fields.Add(s.Range, ref CurrentPage, ref oMissing, ref oMissing);
            // write the "of"
            oWord.ActiveWindow.Selection.TypeText(" of ");
            // create the field for total page number.
            object TotalPages = Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages;
            // insert total pages field in the selection.
            oWord.ActiveWindow.Selection.Fields.Add(s.Range, ref TotalPages, ref oMissing, ref oMissing);
            // return to the document main body.
            oWord.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument;

            //oDoc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter;
            //s = oDoc.ActiveWindow.Selection;
            //s.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //oDoc.ActiveWindow.Selection.TypeText("\tPart Number: " + data.part + "  Serial Number: " + data.serial + "\tPage ");
            //CurrentPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;
            //oDoc.ActiveWindow.Selection.Fields.Add(s.Range, ref CurrentPage, ref oMissing, ref oMissing);
            //oDoc.ActiveWindow.Selection.TypeText("of ");
            //TotalPages = Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages;
            //oDoc.ActiveWindow.Selection.Fields.Add(s.Range, ref TotalPages, ref oMissing, ref oMissing);
            //oDoc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument;

            // INSERT HEADER INFORMATION D:\Downloads
            //oDoc.InlineShapes.AddPicture(@"C:\Program Files (x86)\Pattern Ninja\HereLiesStuff\DocumentHeader-4CMMReport.jpg", oMissing, oMissing, oMissing);
            //oDoc.InlineShapes.AddPicture(@"D:\Downloads\DocumentHeader-4CMMReport.jpg", oMissing, oMissing, oMissing);
            //Word.Paragraphs paragraphs = oDoc.Paragraphs;
            //Word.Paragraph paragraph = paragraphs[1];

            oDoc.Paragraphs.TabStops.Add(oWord.InchesToPoints(1.0f), Word.WdTabAlignment.wdAlignTabLeft);
            oDoc.Paragraphs.TabStops.Add(oWord.InchesToPoints(3.5f), Word.WdTabAlignment.wdAlignTabLeft);
            oDoc.Paragraphs.TabStops.Add(oWord.InchesToPoints(3.25f), Word.WdTabAlignment.wdAlignTabCenter);

            oWord.Selection.Font.Name = "Times New Roman";

            
            
            //oDoc.ActiveWindow.Selection.EndKey(ref oEoF, ref oMissing);
            oDoc.InlineShapes.AddPicture(@"C:\Program Files (x86)\Pattern Ninja\HereLiesStuff\DocumentHeader-4CMMReport.jpg", oMissing, oMissing, oMissing);
            oWord.Selection.InsertBreak(Word.WdBreakType.wdLineBreak);
            //oWord.Selection.Text 

            //oDoc.Range().Text = oDoc.InlineShapes.AddPicture(@"D:\Downloads\DocumentHeader-4CMMReport.jpg", oMissing, oMissing, oMissing) +"\n\n\ntaco bell dog\n";

            oDoc.ActiveWindow.Selection.EndKey(ref oEoF, ref oMissing);

            oWord.Selection.Text = "\tCustomer: " + data.customer + "\t\tInspection Date: " + data.date + "\n" +
                                   "\tPart Number: " + data.part  + "\t\tSerial Number: " + data.serial + "\n" +
                                   "\tName: " + data.name         + "\t\tJob Number: " + data.job + "\n";

            oDoc.ActiveWindow.Selection.EndKey(ref oEoF, ref oMissing);

            oWord.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //oDoc.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            Excel.Range range;
            Excel.Range range2;
            Excel.Range range3;

            oWord.Selection.Text = "\n";
            oDoc.ActiveWindow.Selection.EndKey(ref oEoF, ref oMissing);


            // COPY FIRST TABLE
            if (data.p_file1.Count > 0)
            {
                oWord.Selection.Text = "Point Data";
                oDoc.ActiveWindow.Selection.EndKey(ref oEoF, ref oMissing);
                range = ws.get_Range("A" + p_start, "H" + (p_end - 1));
                range.NumberFormat = "0.0000";
                //range.get_Range("A" + p_start, "A" + p_end).NumberFormat = @"";
                range.Columns[1].NumberFormat = @"";
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.Columns[1].AutoFit();
                range.Columns[2].ColumnWidth = 10;
                range.Columns[3].ColumnWidth = 10;
                range.Columns[4].ColumnWidth = 10;
                range.Columns[5].ColumnWidth = 10;
                range.Columns[6].ColumnWidth = 10;
                range.Columns[7].ColumnWidth = 10;
                range.Columns[8].ColumnWidth = 10;
                range.Copy();
                oWord.Selection.PasteExcelTable(true, true, false);
                //oWord.Selection.Text = "\t";
                //oDoc.ActiveWindow.Selection.EndKey(ref oEoF, ref oMissing);
                if (data.r_file1.Count > 0 || data.d_file1.Count > 0)
                    oWord.Selection.InsertBreak(Word.WdBreakType.wdPageBreak);
            }
            
            // COPY SECOND TABLE
            if (data.d_file1.Count > 0)
            {
                oWord.Selection.Text = "Lineal Dimensions";
                oDoc.ActiveWindow.Selection.EndKey(ref oEoF, ref oMissing);
                range2 = ws2.get_Range("B" + (d_start), "H" + (d_end - 1));
                range2.NumberFormat = "0.0000";
                range2.Columns[1].NumberFormat = @"";
                range2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //range2.Columns[1].AutoFit();
                range2.Columns[2].ColumnWidth = 10;
                range2.Columns[3].ColumnWidth = 10;
                range2.Columns[4].ColumnWidth = 10;
                range2.Columns[5].ColumnWidth = 10;
                range2.Columns[6].ColumnWidth = 10;
                range2.Columns[7].ColumnWidth = 10;
                range2.Copy();
                oWord.Selection.PasteExcelTable(true, true, false);

                if (data.r_file1.Count > 0)
                    oWord.Selection.InsertBreak(Word.WdBreakType.wdPageBreak);

            }

            // COPY THIRD TABLE
            if (data.r_file1.Count > 0)
            {
                oWord.Selection.Text = "Radius/Diameter Dimensions";
                oDoc.ActiveWindow.Selection.EndKey(ref oEoF, ref oMissing);
                range3 = ws3.get_Range("B" + (r_start), "H" + (r_end - 1));
                range3.NumberFormat = "0.0000";
                range3.Columns[1].NumberFormat = @"";
                range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range3.Columns[1].ColumnWidth = 5;
                range3.Columns[2].ColumnWidth = 10;
                range3.Columns[3].ColumnWidth = 10;
                range3.Columns[4].ColumnWidth = 10;
                range3.Columns[5].ColumnWidth = 10;
                range3.Columns[6].ColumnWidth = 10;
                range3.Columns[7].ColumnWidth = 10;
                range3.Copy();
                oWord.Selection.PasteExcelTable(true, true, false);
            }

            oWord.Visible = true;
            //ex.Visible = true;
            wb.Close(false);
            
        }
    }
}