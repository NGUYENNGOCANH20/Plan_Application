using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;

namespace Planner
{
    internal class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                var t = Task.Run(() =>
                {
                    Console.Clear();
                    string inputfile1 = File.ReadAllText(Directory.GetCurrentDirectory() + "\\InDyelot.txt");
                    string inputfile2 = File.ReadAllText(Directory.GetCurrentDirectory() + "\\InCO.txt");
                    string Style = File.ReadAllText(Directory.GetCurrentDirectory() + "\\Sstyle.txt");
                    string Sstyj = "";
                    if(Directory.GetFiles(Style).GetLength(0) == 1)
                    {
                        foreach (var Keyva in Oder.Stylevalue(Directory.GetFiles(Style)[0]))
                        {
                            Sstyj = Sstyj +Keyva.Key +"\t"+Keyva.Value+"\n";
                        }
                        File.WriteAllText(Directory.GetCurrentDirectory()+"\\StyleS.txt", Sstyj);
                        Process process = new Process();
                        process.StartInfo.FileName = "cmd.exe";
                        process.StartInfo.Arguments = @"taskkill \f \im EXCEL.EXE";
                        process.Start();
                        process.Kill();
                        File.Delete(Directory.GetFiles(Style)[0]);
                    }
                    if (Directory.GetFiles(inputfile1).GetLength(0) == 1 && Directory.GetFiles(inputfile2).GetLength(0) == 1)
                    {
                        string namefile1 = Directory.GetFiles(inputfile1)[0];
                        string namefile2 = Directory.GetFiles(inputfile2)[0];
                        Oder oder = new Oder(namefile1, namefile2);
                        File.Delete(namefile1);
                        File.Delete(namefile2);
                        oder = null;
                        GC.Collect();
                    }
                    else
                    {
                        Console.WriteLine("Wrong File Count");
                    }
                });
                t.Wait();
                Thread.Sleep(TimeSpan.FromSeconds(120));
            }
            
        }
    }
}
