using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using ClosedXML.Excel;
using static ClosedXML.Excel.XLPredefinedFormat;
using DateTime = System.DateTime;

namespace MergeFilesWss
{
    class Program
    {
        static void Main(string[] args)
        {
            //ReWrite();             
            Cleaning();
            //CleaningWork();
            //MergeFilesWss(); 
            //KillManager();
            //KillClientDubbed();
            //KillLastDateTo();
            //SetStatus();
        }

        static void SetStatus()
        {
            var files = Directory.GetFiles("Clean");
            XLWorkbook wb = new XLWorkbook(files.First());


            Console.WriteLine("Введите дату закрытия");

            string date = Console.ReadLine();

            DateTime dateDT = new DateTime();

            dateDT = Convert.ToDateTime(date);

            foreach (var page in wb.Worksheets)
            {
                List<int> tagged = new List<int>();

                if (page.Name != "Архив")
                {
                    var Rng = page.RangeUsed();

                    for (int i = 2; i <= Rng.LastRow().RowNumber(); i++)
                    {
                        for (int j = 1; j <= Rng.LastColumn().ColumnNumber() - 1; j++)
                        {
                            if (Regex.Match(page.Cell(1, j).GetString(), "Дата", RegexOptions.IgnoreCase).Success == true && j == 7)
                            {
                                if (page.Cell(i, 7).GetString() == "")
                                {
                                    page.Cell(i, 6).SetValue<string>("Закрыт");
                                    page.Cell(i, 7).SetValue<string>(dateDT.ToString("dd.MM.yyyy"));
                                }
                            }
                        }
                    }
                }
            }
            wb.SaveAs(@"Result\Аналитика " + DateTime.Now.ToString("dd.MM") + ".xlsx");
        }

        static void KillLastDateTo()
        {
            var files = Directory.GetFiles("Clean");
            XLWorkbook wb = new XLWorkbook(files.First());

            Console.WriteLine("Введите дату перед которой удалить");

            string date = Console.ReadLine();

            DateTime dateDT = new DateTime();

            dateDT = Convert.ToDateTime(date);

            foreach (var page in wb.Worksheets)
            {
                List<int> tagged = new List<int>();

                if (page.Name != "Архив")
                {
                    var Rng = page.RangeUsed();

                    for (int i = 2; i <= Rng.LastRow().RowNumber(); i++)
                    {
                        for (int j = 1; j <= Rng.LastColumn().ColumnNumber() - 1; j++)
                        {
                            if (Regex.Match(page.Cell(1, j).GetString(), "Дата", RegexOptions.IgnoreCase).Success == true)
                            {
                                string buf = "";
                                date = page.Cell(i, j).GetString();

                                for(int k = 0, q = 0; k < date.Length; k++)
                                {
                                    if (q < 3)
                                    {
                                        if (date[k] == '.')
                                        {
                                            q++;
                                            buf += '/';
                                        }
                                        else buf += date[k];
                                    }
                                    else
                                    {
                                        // текущий год
                                        buf += "2021";
                                    }
                                }

                                if (buf == "") continue;
                                else if (Convert.ToDateTime(buf) <= dateDT)
                                {
                                    page.Row(i).Delete();

                                    i--;
                                }
                            }
                        }
                    }
                }
            }
            wb.SaveAs(@"Result\Аналитика " + DateTime.Now.ToString("dd.MM") + ".xlsx");
        }

        static void KillClientDubbed()
        {
            string buf = "";

            var files = Directory.GetFiles("Clean");
            XLWorkbook wb = new XLWorkbook(files.First());

            foreach (var page in wb.Worksheets)
            {
                List<int> tagged = new List<int>();

                if (page.Name != "Архив")
                {
                    var Rng = page.RangeUsed();

                    for (int i = 2; i <= Rng.LastRow().RowNumber(); i++)
                    {
                        for (int j = 1; j <= Rng.LastColumn().ColumnNumber() - 1; j++)
                        {
                            if (Regex.Match(page.Cell(1, j).GetString(), "Клиент", RegexOptions.IgnoreCase).Success == true)
                            {
                                buf = page.Cell(i, j).GetString();

                                if (page.Cell(i+1, j).GetString() == buf)
                                {
                                    page.Row(i).Delete();
                                    page.Row(i).Delete();
                                }
                                else 
                                {
                                    buf = "";
                                    j = i + 2;
                                }
                            }
                        }
                    }
                }
            }
            wb.SaveAs(@"Result\Аналитика " + DateTime.Now.ToString("dd.MM") + ".xlsx");
        }

        static void KillManager()
        {
            List<string> nameManagers = new List<string>();
            int qM;

            Console.WriteLine("Введите кол-во удаляемых менеджеров");

            qM = Convert.ToInt32(Console.ReadLine());

            for(int i = 0; i < qM; i++)
            {
                Console.WriteLine("Введите имя менеджера");

                nameManagers.Add(Console.ReadLine());
            }

            var files = Directory.GetFiles("Clean");
            XLWorkbook wb = new XLWorkbook(files.First());

            foreach (var page in wb.Worksheets)
            {
                if (page.Name != "Архив")
                {
                    var Rng = page.RangeUsed();

                    for (int i = 1; i <= Rng.LastRow().RowNumber(); i++)
                    {
                        for (int j = 1; j <= Rng.LastColumn().ColumnNumber() - 1; j++)
                        {
                            if ((Regex.Match(page.Cell(1, j).GetString(), "Ответственный", RegexOptions.IgnoreCase).Success) == true)
                            {
                                for (int k = 0; k < nameManagers.Count; k++)
                                {
                                    if (Regex.Match(page.Cell(i, j).GetString(), (string.Format(@"{0}$", nameManagers[k])), RegexOptions.IgnoreCase).Success)
                                    {
                                        page.Row(i).Delete();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            wb.SaveAs(@"Result\Аналитика " + DateTime.Now.ToString("dd.MM") + ".xlsx");
        }

        static void MergeFilesWss()
        {
            var files = Directory.GetFiles("ToMerge");
            var oldfile = files.Where(f => Regex.Match(f, "старая", RegexOptions.IgnoreCase).Success).First();
            var newfile = files.Where(f => Regex.Match(f, "новая", RegexOptions.IgnoreCase).Success).First();
            XLWorkbook oldwb = new XLWorkbook(oldfile);
            XLWorkbook newwb = new XLWorkbook(newfile);
            foreach (var page in newwb.Worksheets)
            {
                if (page.Name != "Архив")
                {
                    var newRange = page.RangeUsed();
                    try
                    { 
                        var oldRange = oldwb.Worksheet(page.Name).RangeUsed();
                        for (int i = 2; i <= newRange.LastRow().RowNumber(); i++)
                        {

                            for (int j = 2; j < oldRange.LastRow().RowNumber(); j++)
                            {
                                var CellClientOld = oldwb.Worksheet(page.Name).Cell(j, 1);
                                var CellClientNew = page.Cell(i, 1);
                                if (
                                    CellClientNew.GetString() == CellClientOld.GetString())
                                //CellClientNew.HasHyperlink && CellClientOld.HasHyperlink
                                //&& 
                                //CellClientOld.GetHyperlink().ExternalAddress.AbsoluteUri == CellClientNew.GetHyperlink().ExternalAddress.AbsoluteUri
                                //&& page.Cell(i, newRange.LastColumn().ColumnNumber() - 3).GetString() == oldwb.Worksheet(page.Name).Cell(j, newRange.LastColumn().ColumnNumber() - 3).GetString()
                                //&& page.Cell(i, newRange.LastColumn().ColumnNumber() - 3).GetString() != "")
                                {
                                    //for (int k = 0; k < 3; k++)
                                    //{
                                    //    page.Cell(i, newRange.LastColumn().ColumnNumber() - k).SetValue<string>(oldwb.Worksheet(page.Name).Cell(j, newRange.LastColumn().ColumnNumber() - k).GetString());
                                    //    page.Cell(i, newRange.LastColumn().ColumnNumber() - k).Style.Font.FontColor = XLColor.Black;
                                    //}

                                    page.Cell(i, newRange.LastColumn().ColumnNumber() + 1).SetValue<string>(oldwb.Worksheet(page.Name).Cell(j, oldRange.LastColumn().ColumnNumber()).GetString() + " sd");

                                    break;

                                }

                            }
                        }
                    }
                    catch (System.ArgumentException)
                    {

                    }
                }

            }
            newwb.SaveAs(@"Result\Аналитика РНР Хауз 02.06.xlsx");
        }
        static void Cleaning()
        {
            var files = Directory.GetFiles("Clean");
            XLWorkbook wb = new XLWorkbook(files.First());
            bool Belfan = Regex.Match(files.First(), "Белфан", RegexOptions.IgnoreCase).Success;
            List <int> todel = new List<int>();
            IXLWorksheet Missed;
            
                Missed = wb.Worksheets.Add("Упущенные");
                Missed.Cell("A1").Value = "Клиент";
                Missed.Cell("B1").Value = "Ответственный"; 
                Missed.Cell("C1").Value = "Примечание";
                Missed.Cell("D1").Value = "Примечание по CRM";
                Missed.Cell("E1").Value = "Дата упущения";

            foreach (var page in wb.Worksheets)
            {
                if (page.Name != "Архив" && page.Name!= "Упущенные" && !Regex.Match(page.Name, "> 4 недель", RegexOptions.IgnoreCase).Success)
                {
                    var Rng = page.RangeUsed();
                    var lastcol = Rng.LastColumn().ColumnNumber() - 1;
                    DateTime dateNext;
                    for (int i = 2; i <= Rng.LastRow().RowNumber(); i++)
                    {
                        if (!DateTime.TryParse(page.Cell(i, lastcol).GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out dateNext))
                        {
                            DateTime.TryParse(page.Cell(i, lastcol).GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out dateNext);
                        }
                        if ((Regex.Match(page.Cell(i, lastcol - 2).GetString(), "успе", RegexOptions.IgnoreCase).Success
                            || Regex.Match(page.Cell(i, lastcol - 2).GetString(), "удал", RegexOptions.IgnoreCase).Success
                            || Regex.Match(page.Cell(i, lastcol - 2).GetString(), "не открывает", RegexOptions.IgnoreCase).Success
                            || Regex.Match(page.Cell(i, lastcol - 2).GetString(), "не целев", RegexOptions.IgnoreCase).Success
                            || Regex.Match(page.Cell(i, lastcol - 2).GetString(), "нецелев", RegexOptions.IgnoreCase).Success
                            || Belfan)
                            && Regex.Match(page.Cell(i, lastcol - 1).GetString(), "закрыт", RegexOptions.IgnoreCase).Success
                            || (Regex.Match(page.Cell(i, lastcol - 1).GetString(), "работ", RegexOptions.IgnoreCase).Success && dateNext > DateTime.Today.AddDays(-4)
                          ))
                        {
                            var arch = wb.Worksheet("Архив");
                            int lastrowarch = arch.RangeUsed().LastRowUsed().RowNumber() + 1;
                            int lastcolarch = arch.RangeUsed().LastColumnUsed().ColumnNumber();
                            arch.Cell(lastrowarch, lastcolarch).SetValue<string>(String.Format("{0:dd.MM.yyyy}", dateNext));
                            arch.Cell(lastrowarch, 1).Value = page.Cell(i, 1).Value;
                            arch.Cell(lastrowarch, 1).Hyperlink = page.Cell(i, 1).GetHyperlink();
                            arch.Cell(lastrowarch, lastcolarch - 1).SetValue<string>(page.Cell(i, lastcol - 1).GetString());
                            arch.Cell(lastrowarch, lastcolarch - 2).SetValue<string>(page.Cell(i, lastcol - 2).GetString());
                            arch.Cell(lastrowarch, lastcolarch - 3).SetValue<string>(page.Cell(i, lastcol - 3).GetString());
                            arch.Cell(lastrowarch, lastcolarch - 4).SetValue<string>(page.Cell(i, lastcol - 4).GetString());
                            page.Cell(i, 1).Value = "";
                            if (Belfan)
                            {
                                
                                
                                
                                var usedRng = Missed.RangeUsed();
                                int curRowNum = usedRng.LastRow().RowNumber() + 1;
                                var ccell = page.Cell(i, lastcol - 2).GetString();
                                if (!(Regex.Match(page.Cell(i, lastcol - 2).GetString(), "успе", RegexOptions.IgnoreCase).Success
                                        )
                                        && Regex.Match(page.Cell(i, lastcol - 1).GetString(), "закрыт", RegexOptions.IgnoreCase).Success)
                                {
                                    Missed.Cell("A" + curRowNum).Value = arch.Cell(lastrowarch, 1).Value;
                                    Missed.Cell("A" + curRowNum).Hyperlink = arch.Cell(lastrowarch, 1).GetHyperlink(); ;
                                    Missed.Cell("B" + curRowNum).SetValue<string>(page.Cell(i, lastcol - 4).GetString());
                                    Missed.Cell("C" + curRowNum).SetValue<string>(page.Cell(i, lastcol - 3).GetString());
                                    Missed.Cell("D" + curRowNum).SetValue<string>(page.Cell(i, lastcol - 2).GetString());
                                    Missed.Cell("E" + curRowNum).SetValue<string>(String.Format("{0:dd.MM.yyyy}", dateNext));
                                }


                            }
                        }
                    }
                    for (int i = Rng.LastRow().RowNumber(); i >=2 ; i--)
                    {
                        
                        if (page.Cell(i, 1).GetString() == "")
                        {
                            page.Row(i).Delete();
                            //i--;
                        }

                    }
                   

                }
            }

            var missedRng = Missed.RangeUsed();
            missedRng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            missedRng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; ;
            missedRng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            Missed.Columns("1:2,4:5").Width = 20;
            Missed.Column(3).Width = 60;
            Missed.Style.Alignment.WrapText = true;
            Missed.Range("A1:E1").Style.Font.Bold = true;
            if (!Belfan)
                wb.Worksheets.Delete("Упущенные");
            var client = wb.Worksheet(1).Cell(2, 1);
            wb.SaveAs(@"Result\Аналитика " + DateTime.Now.ToString("dd.MM") + ".xlsx");
        }


        static void CleaningWork()
        {
            var files = Directory.GetFiles("Clean");
            XLWorkbook wb = new XLWorkbook(files.First());
            List<int> todel = new List<int>();
            IXLWorksheet page;

            page = wb.Worksheet("Архив");

            
                    var Rng = page.RangeUsed();
                    var lastcol = Rng.LastColumn().ColumnNumber();
                    DateTime dateNext;
                    for (int i = 2; i <= Rng.LastRow().RowNumber(); i++)
                    {
                        if (!DateTime.TryParse(page.Cell(i, lastcol - 1).GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out dateNext))
                        {
                            DateTime.TryParse(page.Cell(i, lastcol - 1).GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out dateNext);
                        }
                        if (!Regex.Match(page.Cell(i, lastcol - 4).GetString(), "успе", RegexOptions.IgnoreCase).Success

                           || (Regex.Match(page.Cell(i, lastcol - 4).GetString(), "успе", RegexOptions.IgnoreCase).Success && dateNext > DateTime.Today.AddDays(-30)) )
                          
                        {
                            
                            page.Cell(i, 1).Value = "";
                        }
                    }
                    for (int i = Rng.LastRow().RowNumber(); i >= 2; i--)
                    {

                        if (page.Cell(i, 1).GetString() == "")
                        {
                            page.Row(i).Delete();
                            //i--;
                        }

                    }

                    
           
            wb.SaveAs(@"Result\Атекс успешные " + DateTime.Now.ToString("dd.MM") + ".xlsx");
        }

        static void ReWrite()
        {
            var files = Directory.GetFiles("Clean");
            XLWorkbook wb = new XLWorkbook(files.First());
            
            wb.SaveAs(@"Result\Аналитика " + DateTime.Now.ToString("dd.MM") + ".xlsx");
        }
    }
}
