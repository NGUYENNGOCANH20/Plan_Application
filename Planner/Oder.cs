using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Exx = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using outlook = Microsoft.Office.Interop.Outlook;
using System.Data;
using System.IO.Compression;

namespace Planner
{
    internal class Oder
    {
        public Oder(string pathdyelot,string pathCOO)
        {
            Exx.Application app = new Exx.Application();
            string Style = File.ReadAllText(Directory.GetCurrentDirectory() + "\\StyleS.txt");
            SortedList<string, string> styles = new SortedList<string, string>();
            foreach(string valuej in Style.Split('\n'))
            {
                if (valuej != null && valuej != ""&& !styles.ContainsKey(valuej.Split('\t')[0]))
                {
                    styles.Add(valuej.Split('\t')[0], valuej.Split('\t')[1]);
                }
            }
            if (File.Exists(Directory.GetCurrentDirectory() + "\\Output Checking.csv"))
            {
                File.Delete(Directory.GetCurrentDirectory() + "\\Output Checking.csv");
            }
            if (File.Exists(Directory.GetCurrentDirectory() + "\\Output.xlsx"))
            {
                File.Delete(Directory.GetCurrentDirectory() + "\\Output.xlsx");
            }
            if (File.Exists(Directory.GetCurrentDirectory() + "\\Input Checking.csv"))
            {
                File.Delete(Directory.GetCurrentDirectory() + "\\Input Checking.csv");
            }
            int itemline = 0;
            int Moqty = 0;
            int dyelotline = 0;
            int Part = 0;
            Exx.Workbook wb = app.Workbooks.Open(pathdyelot);
            Exx.Worksheet ws = wb.Worksheets[1];
            Exx.Range xlRange = ws.UsedRange;
            for (int i = 1; i <= xlRange.Columns.Count; i++)
            {
                if (xlRange.Cells[1, i] != null && xlRange.Cells[1, i].Value2 != null)
                {
                    string checking = string.Join("", xlRange.Cells[1, i].Value2.ToString().Split(' '));
                    if (checking == "WItem")
                    {
                        itemline = i;
                    }
                    if (checking == "MOQty")
                    {
                        Moqty = i;
                    }
                    if (checking == "Schedule")
                    {
                        dyelotline = i;
                    }
                    if (checking == "Part")
                    {
                        Part = i;
                    }
                }
            }
            var Data = new SortedList<string, StyleGroub>();
            if (xlRange.Cells[1, dyelotline].Value != null && xlRange.Cells[1, dyelotline].Value2 != null)
            {
                for (int i = 2; i < xlRange.Rows.Count; i++)
                {
                    Console.WriteLine(i);
                    Console.Clear();
                    if (!_Checkingkey(Data, xlRange.Cells[i, itemline].Value2.ToString()))
                    {
                        StyleGroub style = new StyleGroub(xlRange.Cells[i, itemline].Value2.ToString());
                        string ckck = xlRange.Cells[i, Moqty].Value2.ToString();
                        if (!ckck.Contains('.'))
                        {
                            if (xlRange.Cells[i, Part].Value != null && xlRange.Cells[i, Part].Value2 != null && xlRange.Cells[i, Part].Value2.ToString() != "")
                            {
                                if (_Checkingvalue(style.Queue, xlRange.Cells[1, dyelotline].Value2.ToString()))
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString(), int.Parse(xlRange.Cells[i, Moqty].Value2.ToString()) * int.Parse(xlRange.Cells[i, Part].Value2.ToString()));
                                }
                                else
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString() + " ", int.Parse(xlRange.Cells[i, Moqty].Value2.ToString()) * int.Parse(xlRange.Cells[i, Part].Value2.ToString()));
                                }
                            }
                            else
                            {
                                if (_Checkingvalue(style.Queue, xlRange.Cells[1, dyelotline].Value2.ToString()))
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString(), int.Parse(xlRange.Cells[i, Moqty].Value2.ToString()));
                                }
                                else
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString() + " ", int.Parse(xlRange.Cells[i, Moqty].Value2.ToString()));
                                }
                            }
                        }
                        else
                        {
                            int inputnumbervalue = int.Parse(Math.Round(double.Parse(ckck),0).ToString());
                            if (xlRange.Cells[i, Part].Value != null && xlRange.Cells[i, Part].Value2 != null && xlRange.Cells[i, Part].Value2.ToString() != "")
                            {
                                if (_Checkingvalue(style.Queue, xlRange.Cells[1, dyelotline].Value2.ToString()))
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString(), inputnumbervalue * int.Parse(xlRange.Cells[i, Part].Value2.ToString()));
                                }
                                else
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString() + " ", inputnumbervalue * int.Parse(xlRange.Cells[i, Part].Value2.ToString()));
                                }
                            }
                            else
                            {
                                if (_Checkingvalue(style.Queue, xlRange.Cells[1, dyelotline].Value2.ToString()))
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString(), inputnumbervalue);
                                }
                                else
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString() + " ", inputnumbervalue);
                                }
                            }
                        }
                        Data.Add(style.Codestyle, style);
                    }
                    else
                    {
                        
                        StyleGroub style = Data[xlRange.Cells[i, itemline].Value2.ToString()];
                        string ckck = xlRange.Cells[i, Moqty].Value2.ToString();
                        if (!ckck.Contains('.'))
                        {
                            if (xlRange.Cells[i, Part].Value != null && xlRange.Cells[i, Part].Value2 != null && xlRange.Cells[i, Part].Value2.ToString() != "")
                            {
                                if (_Checkingvalue(style.Queue, xlRange.Cells[1, dyelotline].Value2.ToString()))
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString(), int.Parse(xlRange.Cells[i, Moqty].Value2.ToString()) * int.Parse(xlRange.Cells[i, Part].Value2.ToString()));
                                }
                                else
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString() + " ", int.Parse(xlRange.Cells[i, Moqty].Value2.ToString()) * int.Parse(xlRange.Cells[i, Part].Value2.ToString()));
                                }
                            }
                            else
                            {
                                if (_Checkingvalue(style.Queue, xlRange.Cells[1, dyelotline].Value2.ToString()))
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString(), int.Parse(xlRange.Cells[i, Moqty].Value2.ToString()));
                                }
                                else
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString() + " ", int.Parse(xlRange.Cells[i, Moqty].Value2.ToString()));
                                }
                            }
                        }
                        else
                        {
                            int inputnumbervalue = int.Parse(Math.Round(double.Parse(ckck), 0).ToString());
                            if (xlRange.Cells[i, Part].Value != null && xlRange.Cells[i, Part].Value2 != null && xlRange.Cells[i, Part].Value2.ToString() != "")
                            {
                                if (_Checkingvalue(style.Queue, xlRange.Cells[1, dyelotline].Value2.ToString()))
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString(), inputnumbervalue * int.Parse(xlRange.Cells[i, Part].Value2.ToString()));
                                }
                                else
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString() + " ", inputnumbervalue * int.Parse(xlRange.Cells[i, Part].Value2.ToString()));
                                }
                            }
                            else
                            {
                                if (_Checkingvalue(style.Queue, xlRange.Cells[1, dyelotline].Value2.ToString()))
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString(), inputnumbervalue);
                                }
                                else
                                {
                                    style.Adding(style.Queue, xlRange.Cells[i, dyelotline].Value2.ToString() + " ", inputnumbervalue);
                                }
                            }
                        }
                        Data[xlRange.Cells[i, itemline].Value2.ToString()] = style;
                    }
                }
            }
            string valuek = "";
            foreach (var keyvalue in Data)
            {
                StyleGroub style = keyvalue.Value;
                valuek = valuek + style.Codestyle + "\t";
                foreach (var item in style.Queue)
                {
                    valuek = valuek + item.Key + ":" + item.Value.ToString() + ",";
                }
                valuek = valuek + "\n";
            }
            File.WriteAllText(Directory.GetCurrentDirectory() + "\\Input Checking.csv", String.Join(",",String.Join(",",valuek.Split('\t')).Split(':')));
            Exx.Workbook wb2 = app.Workbooks.Open(pathCOO);
            Exx.Worksheet ws2 = wb2.Worksheets[1];
            Exx.Range xlRange2 = ws2.UsedRange;
            int styleU = 0;
            int clocode = 0;
            int Deltasize = 0;
            int Odernumber = 0;
            int Mot = 0;
            string VKl = "";
            for (int i = 1; i <= xlRange2.Columns.Count; i++)
            {
                if (xlRange2.Cells[1, i] != null && xlRange2.Cells[1, i].Value2 != null)
                {
                    string checking = string.Join("", xlRange2.Cells[1, i].Value2.ToString().Split(' '));
                    if (checking == "WItem")
                    {
                        styleU = i;
                    }
                    if (checking == "DeltaColorCode")
                    {
                        clocode = i;
                    }
                    if (checking == "DeltaSize")
                    {
                        Deltasize = i;
                    }
                    if (checking == "MOT")
                    {
                        Mot = i;
                    }
                    if (checking == "OrderNumber")
                    {
                        Odernumber = i;
                    }
                }
            }
            xlRange2.Cells[1, xlRange2.Columns.Count + 1].Value2 = "Dyelot number "+DateTime.Now;
            xlRange2.Cells[1, xlRange2.Columns.Count + 2].Value2 = "S Style" + DateTime.Now;
            xlRange2.Cells[1, xlRange2.Columns.Count + 3].Value2 = "S Item" + DateTime.Now;
            xlRange2.Cells[1, xlRange2.Columns.Count + 4].Value2 = "S Color" + DateTime.Now;
            for (int i = 3; i <= xlRange2.Rows.Count; i++)
            {
                if (xlRange2.Cells[i, Odernumber].Value != null && xlRange2.Cells[i, Odernumber].Value2 != null && xlRange2.Cells[i, Odernumber].Value2.ToString()!="0")
                {
                    string colocr = xlRange2.Cells[i, clocode].Value2.ToString();
                    string stylerun = "";
                    string lvaje = "2BT/2DT/2QT/3BT/3BT/3CT/3QT/4QT/6QT/9QT/3DT/4BT/4DT/6BT/6CT/6DT/8BT/8CT/8DT/8QT/9BT/9CT/9DT";
                    List<string> lines = new List<string>();
                    foreach (string ovki in lvaje.Split('/'))
                    {
                        lines.Add(ovki);
                    }
                    if(lines.Contains(xlRange2.Cells[i, Deltasize].Value2.ToString()))
                    {
                        string suz = xlRange2.Cells[i, Deltasize].Value2.ToString();
                        stylerun = xlRange2.Cells[i, styleU].Value2.ToString() + "-" + colocr.Substring(1, colocr.Length - 1) + suz.Substring(0,suz.Length-1);
                    }
                    else
                    {
                        string suz = xlRange2.Cells[i, Deltasize].Value2.ToString();
                        stylerun = xlRange2.Cells[i, styleU].Value2.ToString() + "-" + colocr.Substring(1, colocr.Length - 1) + suz;
                    }
                    string COOO = xlRange2.Cells[i, Odernumber].Value2.ToString();
                    if (styles.ContainsKey(COOO + "-" + stylerun))
                    {
                        xlRange2.Cells[i, xlRange2.Columns.Count + 2].Value = styles[xlRange2.Cells[i, Odernumber].Value2.ToString() + "-" + stylerun].Split('.')[0];
                        xlRange2.Cells[i, xlRange2.Columns.Count + 3].Value = styles[xlRange2.Cells[i, Odernumber].Value2.ToString() + "-" + stylerun].Split('.')[1];
                        xlRange2.Cells[i, xlRange2.Columns.Count + 4].Value = styles[xlRange2.Cells[i, Odernumber].Value2.ToString() + "-" + stylerun].Split('.')[2];
                    }
                    Console.Write(i);
                    Console.Clear();
                    double oderqty = 2 * double.Parse(xlRange2.Cells[i, Mot].Value2.ToString());
                    if (oderqty != 0)
                    {
                        if (_Checkingkey(Data, stylerun))
                        {
                            StyleGroub style = Data[stylerun];
                            var datakeyvalue = style.Queue;
                            if (datakeyvalue.Count > 0)
                            {
                                string Schedule = "";
                                Queue<KeyValuePair<string, int>> Queue = new Queue<KeyValuePair<string, int>>();
                                foreach (KeyValuePair<string, int> pair in datakeyvalue)
                                {
                                    if (oderqty == 0)
                                    {
                                        Queue.Enqueue(new KeyValuePair<string, int>(pair.Key, pair.Value));
                                    }
                                    else
                                    {
                                        Queue.Enqueue(new KeyValuePair<string, int>(pair.Key, pair.Value));
                                        if (pair.Value > oderqty)
                                        {
                                            Queue.Dequeue();
                                            Queue.Enqueue(new KeyValuePair<string, int>(pair.Key, pair.Value - int.Parse(oderqty.ToString())));
                                            Schedule = Schedule + pair.Key + "-" + (oderqty).ToString() + "/";
                                            oderqty = 0;
                                        }
                                        else
                                        {
                                            if (pair.Value < oderqty)
                                            {
                                                Queue.Dequeue();
                                                oderqty = oderqty - pair.Value;
                                                Schedule = Schedule + pair.Key + "-" + pair.Value.ToString() + "/";
                                            }
                                            else
                                            {
                                                Queue.Dequeue();
                                                Schedule = Schedule + pair.Key + "-" + (oderqty).ToString() + "/";
                                                oderqty = 0;
                                            }
                                        }
                                    }
                                }
                                if (oderqty > 0)
                                {
                                  Schedule = Schedule + $" Missing {oderqty}";
                                }
                                else
                                {
                                }
                                VKl = VKl + $"{stylerun},{xlRange2.Cells[i, Odernumber].Value2.ToString()},{Schedule}\n";
                                style.Queue = Queue;
                                Data[stylerun] = style;
                                xlRange2.Cells[i, xlRange2.Columns.Count + 1].Value = Schedule;
                            }
                            else
                            {
                                xlRange2.Cells[i, xlRange2.Columns.Count + 1].Value = "Out of stock";
                            }
                        }
                        else
                        {
                            xlRange2.Cells[i, xlRange2.Columns.Count + 1].Value = "Can't find style in Dyelot";
                        }
                    }
                    else
                    {
                        xlRange2.Cells[i, xlRange2.Columns.Count + 1].Value = "Request Zero";
                        xlRange2.Cells[i, xlRange2.Columns.Count + 1].Value = "Request Zero";
                    }
                }
            }
            File.WriteAllText(Directory.GetCurrentDirectory() + "\\Output Checking.csv", String.Join(",", VKl.Split('\t')));
            ws2.SaveAs(Directory.GetCurrentDirectory() + "\\Output.xlsx");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(ws);
            wb.Close();
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(xlRange2);
            Marshal.ReleaseComObject(ws2);
            wb2.Close();
            Marshal.ReleaseComObject(wb2);
            app.Quit();
            Marshal.ReleaseComObject(app);
            outlook.Application appOut = new outlook.Application();
            outlook.MailItem mail = (outlook.MailItem)appOut.CreateItem(outlook.OlItemType.olMailItem);
            mail.To =File.ReadAllText(Directory.GetCurrentDirectory()+ "\\txt_to.txt");
            mail.Subject = "CO_Production_Balance Repost";
            mail.Body = "Dear team, Pls find file in the attach\nDon't relly this mail";
            mail.Importance = outlook.OlImportance.olImportanceHigh;
            Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\Zipfile\\");
            File.Copy(Directory.GetCurrentDirectory() + "\\Output Checking.csv", Directory.GetCurrentDirectory() + "\\Zipfile\\Output Checking.csv");
            File.Copy(Directory.GetCurrentDirectory() + "\\Output.xlsx", Directory.GetCurrentDirectory() + "\\Zipfile\\Output.xlsx");
            File.Copy(Directory.GetCurrentDirectory() + "\\Input Checking.csv", Directory.GetCurrentDirectory() + "\\Zipfile\\Input Checking.csv");
            File.Copy(pathdyelot, Directory.GetCurrentDirectory() + "\\Zipfile\\"+Path.GetFileName(pathdyelot));
            ZipFile.CreateFromDirectory(Directory.GetCurrentDirectory() + "\\Zipfile\\", Directory.GetCurrentDirectory() + "\\zipoutput_data.zip");
            mail.Attachments.Add(Directory.GetCurrentDirectory() + "\\zipoutput_data.zip");
            ((outlook.MailItem)mail).Send();
            File.Delete(Directory.GetCurrentDirectory() + "\\zipoutput_data.zip");
            Directory.Delete(Directory.GetCurrentDirectory() + "\\Zipfile\\");
        }
        public bool _Checkingvalue(Queue<KeyValuePair<string, int>> valuePair,string checking)
        {
            foreach (var gj in valuePair)
            {
                if (gj.Key == checking)
                {
                    return true;
                }
            }
            return false;
        }
        public bool _Checkingkey(SortedList<string, StyleGroub> keydd,string keyname)
        {
            foreach (var keyv in keydd)
            {
                if (keyname == keyv.Key)
                {
                    return true;
                }
            }
            return false;
        }
        public class StyleGroub
        {
            public StyleGroub(string style)
            {
                this.Codestyle = style;
            }
            public string Codestyle;
            public Queue<KeyValuePair<string, int>> Queue = new Queue<KeyValuePair<string, int>>();
            public void Adding(Queue<KeyValuePair<string, int>> Queue,string Schedule,int Qty)
            {
                Queue.Enqueue(new KeyValuePair<string, int>(Schedule, Qty));
            }
        }
        public static SortedList<string, string> Stylevalue(string pathname)
        {
            Exx.Application app = new Exx.Application();
            SortedList<string, string> styles = new SortedList<string, string>();
            Exx.Workbook wb2 = app.Workbooks.Open(pathname);
            Exx.Worksheet ws2 = wb2.Worksheets[1];
            Exx.Range xlRange2 = ws2.UsedRange;
            int styleU = 0;
            int Style = 0;
            int Sitem = 0;
            int Odernumber = 0;
            int Sclor = 0;
            for (int i = 1; i <= xlRange2.Columns.Count; i++)
            {
                if (xlRange2.Cells[1, i] != null && xlRange2.Cells[1, i].Value2 != null)
                {
                    string checking = string.Join("", xlRange2.Cells[1, i].Value2.ToString().Split(' '));
                    if (checking == "WItem")
                    {
                        styleU = i;
                    }
                    if (checking == "SStyle")
                    {
                        Style = i;
                    }
                    if (checking == "SItem")
                    {
                        Sitem = i;
                    }
                    if (checking == "SColor")
                    {
                        Sclor = i;
                    }
                    if (checking == "OrderNumber")
                    {
                        Odernumber = i;
                    }
                }
            }
            for (int i = 2; i <= xlRange2.Rows.Count; i++)
            {
                Console.WriteLine(i);
                if (xlRange2.Cells[i, Odernumber].Value != null && xlRange2.Cells[i, Odernumber].Value2 != null && xlRange2.Cells[i, Odernumber].Value2.ToString() != "0")
                {
                    if(!styles.ContainsKey(xlRange2.Cells[i, Odernumber].Value2.ToString() + "-" + xlRange2.Cells[i, styleU].Value2.ToString()))
                    {
                        styles.Add(xlRange2.Cells[i, Odernumber].Value2.ToString() + "-" + xlRange2.Cells[i, styleU].Value2.ToString(), xlRange2.Cells[i, Style].Value2.ToString() + "." + xlRange2.Cells[i, Sitem].Value2.ToString() + "." + xlRange2.Cells[i, Sclor].Value2.ToString());
                    }
                }
            }
            GC.Collect();
            Marshal.ReleaseComObject(wb2);
            Marshal.ReleaseComObject(xlRange2);
            Marshal.ReleaseComObject(ws2);
            Marshal.ReleaseComObject(wb2);
            Marshal.ReleaseComObject(app);
            return styles;
        }
    }
}
