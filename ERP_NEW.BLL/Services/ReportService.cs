using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

using ERP_NEW.BLL.Interfaces;
using ERP_NEW.DAL.Interfaces;
using ERP_NEW.BLL.DTO.ModelsDTO;
using ERP_NEW.BLL.DTO.ReportsDTO;
using ERP_NEW.BLL.Infrastructure;
using System.IO;
using SpreadsheetGear;
using System.Diagnostics;
using System.Windows.Forms;
using ERP_NEW.DAL.Entities.ReportModel;
using FirebirdSql.Data.FirebirdClient;

using ERP_NEW.BLL.DTO.SelectedDTO;
using Words = Microsoft.Office.Interop.Word;
using System.Globalization;
using ERP_NEW.BLL.NameCaseLib;
using System.Text.RegularExpressions;
using Nager.Date;
using AutoMapper;
using ERP_NEW.DAL.Entities.Models;
using ERP_NEW.DAL.Entities.QueryModels;
using ERP_NEW.BLL.DTO;



namespace ERP_NEW.BLL.Services
{
    public class ReportService : IReportService
    {
        private string GeneratedReportsDir = Utils.HomePath + @"\Temp\";

        private Words._Application word;
        private Words._Document document;

        private IUnitOfWork Database { get; set; }

        
        private IRepository<FixedAssetsOrderJournalPrint> fixedAssetsOrderJournalPrint;
      

        private IMapper mapper;

        public ReportService(IUnitOfWork uow)
        {
            Database = uow;

           var config = new MapperConfiguration(cfg =>
            {       
               
                cfg.CreateMap<FixedAssetsOrderJournalPrint, FixedAssetsOrderJournalPrintDTO>();
                cfg.CreateMap<FixedAssetsOrderJournalPrintDTO, FixedAssetsOrderJournalPrint>();
            });

            mapper = config.CreateMapper();
        }



        #region TimeSheet report's


        public void PrintTimeSheet(List<EmployeesInfoDTO> source, DateTime currentDate)
        {

            string templateName = " ";
            int days = DateTime.DaysInMonth(currentDate.Year, currentDate.Month);

            switch (days)
            {
                case 28:
                    templateName = @"\Templates\TimeSheet28daysTemplate.xls";
                    break;
                case 29:
                    templateName = @"\Templates\TimeSheet29daysTemplate.xls";
                    break;
                case 30:
                    templateName = @"\Templates\TimeSheet30daysTemplate.xls";
                    break;
                case 31:
                    templateName = @"\Templates\TimeSheet31daysTemplate.xls";
                    break;
            }

            try
            {
                Factory.GetWorkbook(GeneratedReportsDir + templateName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не знайдено шаблон документа!\n" + ex.Message, "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var Workbook = Factory.GetWorkbook(GeneratedReportsDir + templateName);
            var Worksheet = Workbook.Worksheets[0];
            var Сells = Worksheet.Cells;
            int startWith = 8;
            int weekendDays = 0;
            string nameDepart = "";

            IRange cells = Worksheet.Cells;
            int recCount = source.Count();
            Dictionary<string, byte> HeaderColumn = new Dictionary<string, byte>();
            int startPosition = 4, currentPosition = startPosition + 2;

        


            Сells["X2"].Value = currentDate.Year;

            cells["P2"].Value = currentDate.Month;

            int currentColumn = 2;

            for (int i = 1; i <= days; i++)
            {
                if (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, i), CountryCode.UA) || DateSystem.IsWeekend(new DateTime(currentDate.Year, currentDate.Month, i), CountryCode.UA))
                {
                    weekendDays++;
                    cells[vsS[currentColumn + i] + "5" + ":" + vsS[currentColumn + i] + "8"].Interior.Color = Color.DodgerBlue;
                }

                //           cells["" + (startWith + source.Count) + ":" + (startWith + source.Count)].Delete();
            }
            for (int i = 0; i < recCount; i++)
            {
                currentPosition++;
            }
            string FullNameWithoutProfession = "";
            int depId = 0;

            //SpreadsheetGear.IRange cellsa = Worksheet.Cells[1,200];
            //cellsa.Merge();

            for (int i = 0; i < source.Count; i++)
            {
                //pin chief in excel
                depId = source[i].DepartmentID.Value;

                switch (depId)
                {
                    case 3://бухгалтерия
                        Сells[source.Count + 15, 2].Value = "Сергієнко Л.В.";
                        break;
                    case 10://админист
                        Сells[source.Count + 15, 2].Value = "Пархоменко Н. М.";
                        break;
                    case 14://технический
                        Сells[source.Count + 15, 2].Value = "Дорошенко В. А. ";
                        break;
                    case 17://плановый
                        Сells[source.Count + 15, 2].Value = "Пінчук М. М. ";
                        break;
                    case 18://маркетинг
                        Сells[source.Count + 15, 2].Value = "Шалаєвський І. М.";
                        break;
                    case 20://материально технический
                        Сells[source.Count + 15, 2].Value = "Романенко І. В. ";
                        break;
                    case 21://конструкторский
                        Сells[source.Count + 15, 2].Value = "Малюсейко В. М. ";
                        break;
                    case 25://транспортный
                        Сells[source.Count + 15, 2].Value = "Дементєєв Р. А. ";
                        break;
                    case 26://технический контороль ОТК
                        Сells[source.Count + 15, 2].Value = "Лисенко О. В. ";
                        break;
                    case 28://энерго-механ
                        Сells[source.Count + 15, 2].Value = "Зайцев В. І.";
                        break;
                    case 29://сто
                        Сells[source.Count + 15, 2].Value = "Бондаренко В. В. ";
                        break;
                    case 30://охрана
                        Сells[source.Count + 15, 2].Value = "Вуцело С. Г. ";
                        break;
                    case 32://комплексное обслуживание и ремонт
                        Сells[source.Count + 15, 2].Value = "";
                        break;
                    case 39://столовая
                        Сells[source.Count + 15, 2].Value = "Вороновська Н. Ф. ";
                        break;
                    case 43://информационные технологии
                        Сells[source.Count + 15, 2].Value = "Шишкін О. М. ";
                        break;
                    case 46://технологическое бюро
                        Сells[source.Count + 15, 2].Value = "Бобрівник В. В. ";
                        break;
                    case 51://инструментальное господар.
                        Сells[source.Count + 15, 2].Value = "Ведмідь В. Ю. ";
                        break;
                    case 53://асуп
                        Сells[source.Count + 15, 2].Value = "Телятник  М. С. ";
                        break;
                    case 54://системы автоматич проектирования
                        Сells[source.Count + 15, 2].Value = "Горбенко С. Г. ";
                        break;
                    case 56://охрана труда
                        Сells[source.Count + 15, 2].Value = "Баранник С. А. ";
                        break;
                    case 58://юрист
                        Сells[source.Count + 15, 2].Value = "Яковенко Н. В. ";
                        break;
                    case 61://договорник
                        Сells[source.Count + 15, 2].Value = "Дузік О. В. ";
                        break;
                    case 63://готовая продукция
                        Сells[source.Count + 15, 2].Value = "Костиренко С. В.";
                        break;
                    case 65://лаборатория сварщиков
                        Сells[source.Count + 15, 2].Value = "Позябін В.І. ";
                        break;
                    case 66://научно техничес
                        Сells[source.Count + 15, 2].Value = " ";
                        break;
                    case 68://господарс.
                        Сells[source.Count + 15, 2].Value = "Вуцело С. Г.";
                        break;
                }


                cells["" + startWith + ":" + startWith].Insert();
                int indexOfChar = source[i].FullName.IndexOf('(');
                FullNameWithoutProfession = source[i].FullName.Substring(0, indexOfChar);
                int dopWeekend = 0;
                currentColumn = 2;

                for (int j = 1; j <= days; j++)
                {
                    int startCell = j + 4;

                    DateTime daysOfWeek = new DateTime(currentDate.Year, currentDate.Month, j);

                    if (daysOfWeek.DayOfWeek == DayOfWeek.Monday)
                    {
                        if (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA))
                        {
                            cells[vsS[currentColumn + j] + startWith].Value =  "ВС";
                            // cells[vsS[currentColumn + j-1] + startWith].Value = "вMon-1";
                        }
                      //  else { cells[vsS[currentColumn + j] + startWith].Value = 100; }
                    }
                   // if (daysOfWeek.DayOfWeek == DayOfWeek.Monday && !(DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA)) && (daysOfWeek.DayOfWeek == DayOfWeek.Sunday) && (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA)))
                      //  else   { cells[vsS[currentColumn + j] + startWith].Value = 100; }



                    if (daysOfWeek.DayOfWeek == DayOfWeek.Tuesday)
                    {
                        if (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA))
                        {
                            cells[vsS[currentColumn + j] + startWith].Value =  "ВС";
                            cells[vsS[currentColumn + j - 1] + startWith].Value = 7;
                        }
                        else { cells[vsS[currentColumn + j] + startWith].Value = 8; }
                    }


                    if (daysOfWeek.DayOfWeek == DayOfWeek.Wednesday)
                    {
                        if (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA))
                        {
                            cells[vsS[currentColumn + j] + startWith].Value =  "ВС";
                            cells[vsS[currentColumn + j - 1] + startWith].Value = 7;
                        }
                        else { cells[vsS[currentColumn + j] + startWith].Value = 8; }
                    }

                    if (daysOfWeek.DayOfWeek == DayOfWeek.Thursday)
                    {
                        if (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA))
                        {
                            cells[vsS[currentColumn + j] + startWith].Value =  "ВС";
                            cells[vsS[currentColumn + j - 1] + startWith].Value = 7;
                        }
                        else { cells[vsS[currentColumn + j] + startWith].Value = 8; }

                    }
                    if (daysOfWeek.DayOfWeek == DayOfWeek.Friday)
                    {
                        if (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA))
                        {
                            cells[vsS[currentColumn + j] + startWith].Value =  "ВС";
                            cells[vsS[currentColumn + j - 1] + startWith].Value = 7;
                        }
                        else { cells[vsS[currentColumn + j] + startWith].Value = 8; }

                    }

                        int curcolSut = 2;
                        if ((daysOfWeek.DayOfWeek == DayOfWeek.Saturday) && (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA)))
                        {
                            if (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA))
                            {
                                cells[vsS[currentColumn + j] + startWith].Value ="ВС";
                                cells[vsS[currentColumn + j - 1] + startWith].Value = 7;

                                cells[vsS[currentColumn + j + curcolSut] + startWith].Value =  "BC";
                                cells[vsS[currentColumn + j + curcolSut] + "5" + ":" + vsS[currentColumn + j + curcolSut] + "8"].Interior.Color = Color.DodgerBlue;
                            }
                            else { cells[vsS[currentColumn + j] + startWith].Value = 8; }
                        }
                        else
                        {
                            if (j < days - 1)
                            {
                                cells[vsS[currentColumn + j + curcolSut] + startWith].Value = 8;
                            }
                        }

                        int curcolSund = 1;


                        if (((daysOfWeek.DayOfWeek == DayOfWeek.Saturday)||(daysOfWeek.DayOfWeek == DayOfWeek.Sunday) )&& (DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA)))
                        {
                            dopWeekend += curcolSund;
                            
                            cells[vsS[currentColumn + j + curcolSund] + startWith].Value = "BC5";

                            cells[vsS[currentColumn + j + 1] + startWith].Value = "BC";
                            cells[vsS[currentColumn + j + curcolSund] + "5" + ":" + vsS[currentColumn + j + curcolSund] + "8"].Interior.Color = Color.DodgerBlue;

                        }
                        else
                        {
                            //if (j < days - 1)
                            //{
                            //  cells[vsS[currentColumn + j + curcolSund] + startWith].Value = 8;
                            //}
                        }
                        DateTime dt = new DateTime(currentDate.Year, currentDate.Month, j);
                        int dtw = dt.Date.Day;

                        DateTime startAtMonday = DateTime.Now.AddDays(DayOfWeek.Monday - DateTime.Now.DayOfWeek);
                        bool dayH = DateSystem.IsPublicHoliday(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA);
                        if ((dtw == 1) && (daysOfWeek.DayOfWeek == DayOfWeek.Monday) || (dayH = false))
                            cells[vsS[currentColumn + j] + startWith].Value = 8;

                        if ((dtw == 2) && (daysOfWeek.DayOfWeek == DayOfWeek.Monday) || (dayH = false))
                            cells[vsS[currentColumn + j] + startWith].Value = 8;

                        //2024 year won't work    
                        //For december 2019 if need special weeked
                        //if (currentDate.Month == 12)
                        //{
                        //    cells[vsS[currentColumn + 21] + startWith].Value = 8;
                        //    cells[vsS[currentColumn + 21] + "5" + ":" + vsS[currentColumn + 21] + "8"].Interior.Color = Color.White;
                        //    cells[vsS[currentColumn + 28] + startWith].Value = 8;
                        //    cells[vsS[currentColumn + 28] + "5" + ":" + vsS[currentColumn + 28] + "8"].Interior.Color = Color.White;

                        //    cells[vsS[currentColumn + 30] + startWith].Value = "ВС";
                        //    cells[vsS[currentColumn + 30] + "5" + ":" + vsS[currentColumn + 30] + "8"].Interior.Color = Color.DodgerBlue;

                        //    cells[vsS[currentColumn + 31] + startWith].Value = "ВС";
                        //    cells[vsS[currentColumn + 31] + "5" + ":" + vsS[currentColumn + 31] + "8"].Interior.Color = Color.DodgerBlue;
                        //}

                        if (DateSystem.IsWeekend(new DateTime(currentDate.Year, currentDate.Month, j), CountryCode.UA))
                            cells[vsS[currentColumn + j] + startWith].Value =  "ВС";

                    Сells["A" + startWith].EntireRow.AutoFit();
                    Сells["A:BB"].Style.Font.Size = 11;

                    Сells["A" + startWith].Value = FullNameWithoutProfession;

                    Сells["B" + 3].Value = source[i].DepartmentName;
                    Сells["B" + 3].Font.Bold = true;

                    Сells["B" + startWith].Orientation = 0;
                    Сells["B" + startWith].Value = source[i].AccountNumber;
                    Сells["B" + 8].HorizontalAlignment = HAlign.Center;

                    Сells["C" + startWith].WrapText = true;
                    Сells["C" + startWith].Value = source[i].ProfessionName;


                    // all days of working by employees
                    Сells[vsS[days + 4] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 4].Formula = "=countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"<9\"" + ")";
                  
                    //annual vacation by employees
                    Сells[vsS[days + 6] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 6].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"В\"" +
                      ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"В\"" +
                        ")," + "\"\"" + ")";
                    
                    //------------------
                    Сells[vsS[days + 7] + startWith].Orientation = 0;
                    //формулы для работы с макросами. Подсчет В,ВД,ТН
                    cells[startWith - 1, days + 7].Formula = "=IF(" + vsS[days + 6] + startWith +"<>"+ "\"\"" + " ," + vsS[days + 6] + startWith + "-" + vsS[days + 31] + startWith +","+"\"\""+ ")";
                    cells[startWith - 1, days + 31].Formula = "=SumV( D" + startWith + ":" + vsS[days + 2] + startWith +","+"BK7"+")";

                    cells[startWith - 1, days + 9].Formula = "=IF(" + vsS[days + 8] + startWith + "<>" + "\"\"" + " ," + vsS[days + 8] + startWith + "-" + vsS[days + 32] + startWith + "," + "\"\"" + ")";
                    cells[startWith - 1, days + 32].Formula = "=SumTN( D" + startWith + ":" + vsS[days + 2] + startWith + "," + "BM7" + ")";

                    cells[startWith - 1, days + 13].Formula = "=IF(" + vsS[days + 12] + startWith + "<>" + "\"\"" + " ," + vsS[days + 12] + startWith + "-" + vsS[days + 33] + startWith + "," + "\"\"" + ")";
                    cells[startWith - 1, days + 33].Formula = "=SumVD( D" + startWith + ":" + vsS[days + 2] + startWith + "," + "BL7" + ")";
                    //    "=IF((" + vsS[days + 31] + startWith + "- " +vsS[days + 32] + startWith +
                    //   ")" + ">0," + "(" + vsS[days + 31] + startWith + "- " + vsS[days + 32] + startWith + ")," + "\"\"" + ")";

               /*     cells[startWith - 1, days + 31].Formula = "=CountCellsByColor( D" + startWith + ":" + vsS[days + 2] + startWith + "," + "A1" + ")";
                    cells[startWith - 1, days + 32].Formula =
                   "=countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"<9\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ДД\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"СТ\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"НА\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ТН\"" + ")+"+ 
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ВП\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"УВ\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ВД\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"С\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"НД\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"Д\"" + ")+" +
                   "countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"П\"" + ")";
            */    //    ")";


                    //--------------------
                    
                    //temporary disability (sick)
                    Сells[vsS[days + 8] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 8].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ТН\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ТН\"" + ")," + 0 + ")";

                    //non-appearance with the permission of the administration
                    Сells[vsS[days + 10] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 10].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"НА\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"НА\"" + ")," + 0 + ")";

                    //non-attendance from the administration's initiative
                    Сells[vsS[days + 11] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 11].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"НД\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"НД\"" + ")," + 0 + ")";

                    //assignment   
                    Сells[vsS[days + 12] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 12].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ВД\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ВД\"" + ")," + 0 + ")";

                    //absenteeism
                    Сells[vsS[days + 14] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 14].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"П\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"П\"" + ")," + 0 + ")";

                    //short week
                    Сells[vsS[days + 15] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 15].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"СТ\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"СТ\"" + ")," + 0 + ")";

                    //weekend and holiday days
                    //Сells[vsS[days + 16] + startWith].Orientation = 0;
                    //cells[startWith - 1, days + 16].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ВС\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"ВС\"" + ")," + "\"\"" + ")";

                    // all days weekend of working by employees
                    Сells[vsS[days + 16] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 16].Value = weekendDays + dopWeekend;


                    // total calendar days
                    Сells[vsS[days + 17] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 17].Formula = "=counta( D" + startWith + ":" + vsS[days + 2] + startWith + ")";

                    //hours worked
                    Сells[vsS[days + 18] + startWith].Orientation = 0;
                    cells[startWith - 1, days + 18].Formula = "=SUM( D" + startWith + ":" + vsS[days + 2] + startWith + ")";

                    ////hours worked in holiday or weekend
                    //Сells[vsS[days + 19] + startWith].Orientation = 0;
                    //cells[startWith - 1, days + 19].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"РВ\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"РВ\"" + ")," + "\"\"" + ")";
                    ////hours worked overdue
                    //Сells[vsS[days + 20] + startWith].Orientation = 0;
                    //cells[startWith - 1, days + 20].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"НУ\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"НУ\"" + ")," + "\"\"" + ")";
                    ////hours worked in night time
                    //Сells[vsS[days + 21] + startWith].Orientation = 0;
                    //cells[startWith - 1, days + 21].Formula = "=IF(countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"РН\"" + ")>0,countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"РН\"" + ")," + "\"\"" + ")";

                    //part-time
                    cells[startWith - 1, days + 26].Formula = "=IF(countif(" + vsS[days + 22] + startWith + " ," + "\"С\"" + ")>0,SUM(" + vsS[days + 4] + startWith + ")," + "\"\"" + ")";

                    // worked part-time weekend
                    cells[startWith - 1, days + 27].Formula = "=IF(countif(" + vsS[days + 22] + startWith + " ," + "\"С\"" + ")>0,SUM(" + vsS[days + 5] + startWith + ")," + "\"\"" + ")";

                    // weekend part-time weekend and holiday
                    cells[startWith - 1, days + 28].Formula = "=IF(countif(" + vsS[days + 22] + startWith + " ," + "\"С\"" + ")>0,SUM(" + vsS[days + 16] + startWith + ")," + "\"\"" + ")";

                    // weekend part-time all day
                    cells[startWith - 1, days + 29].Formula = "=IF(countif(" + vsS[days + 22] + startWith + " ," + "\"С\"" + ")>0,SUM(" + vsS[days + 17] + startWith + ")," + "\"\"" + ")";

                    // weekend part-time all time
                    cells[startWith - 1, days + 30].Formula = "=IF(countif(" + vsS[days + 22] + startWith + " ," + "\"С\"" + ")>0,SUM(" + vsS[days + 18] + startWith + ")," + "\"\"" + ")";


                    //совмещение
                    //Сells[vsS[days + 22] + startWith].Orientation = 0;
                    //cells[startWith - 1, days + 22].Formula = "=countif( D" + startWith + ":" + vsS[days + 2] + startWith + " ," + "\"с\"" + ")";

                    Сells[vsS[days + 5] + startWith].Orientation = 0;
                    Сells[vsS[days + 7] + startWith].Orientation = 0;
                    Сells[vsS[days + 9] + startWith].Orientation = 0;
                    Сells[vsS[days + 11] + startWith].Orientation = 0;
                    Сells[vsS[days + 12] + startWith].Orientation = 0;
                    Сells[vsS[days + 13] + startWith].Orientation = 0;
                    Сells[vsS[days + 15] + startWith].Orientation = 0;
                    Сells[vsS[days + 16] + startWith].Orientation = 0;
                    Сells[vsS[days + 17] + startWith].Orientation = 0;
                    Сells[vsS[days + 18] + startWith].Orientation = 0;
                    Сells[vsS[days + 19] + startWith].Orientation = 0;
                    Сells[vsS[days + 20] + startWith].Orientation = 0;
                    Сells[vsS[days + 21] + startWith].Orientation = 0;
                    Сells[vsS[days + 22] + startWith].Orientation = 0;
                    Сells[vsS[days + 23] + startWith].Orientation = 0;
                    Сells[vsS[days + 24] + startWith].Orientation = 0;

                    //style line in table
                    cells[startWith - 1, j].Borders.LineStyle = LineStyle.None;
                    cells[startWith - 1, 0].Borders.LineStyle = LineStyle.None;
                    cells[startWith - 1, j + 22].Borders.LineStyle = LineStyle.None;
                    cells[startWith - 1, j].Borders.LineStyle = LineStyle.Continuous;
                    cells[startWith - 1, 0].Borders.LineStyle = LineStyle.Continuous;
                    cells[startWith - 1, j + 22].Borders.LineStyle = LineStyle.Continuous;

                    //bold text in number's month
                    cells[startWith - 1, j + 2].Font.Bold = true;
                    cells[startWith - 1, days - 2].Font.Bold = true;

                    //cells[days,j+22].Borders.Weight = BorderWeight.Medium; bold line
                    //cells[days,j+22].Borders[BordersIndex.EdgeBottom].LineStyle = LineStyle.Continous;
                }

                //cells[startWith, days].Borders[BordersIndex.EdgeBottom].LineStyle = LineStyle.Continous;
                //cells[startWith, days].Borders.LineStyle = LineStyle.Continous;

                //calculation of days: where 7 hours, day off or holiday 
                int previousDays = DateTime.DaysInMonth(currentDate.Year, currentDate.Month);

                var lastDay = new DateTime(currentDate.Year, currentDate.Month, 1).AddMonths(1);
                var lastDayqq = new DateTime(currentDate.Year, currentDate.Month, 1).AddMonths(1).AddDays(-1);
                for (int n = 1; n <= previousDays; ++n)
                {

                    if ((DateSystem.IsPublicHoliday(lastDay, CountryCode.UA)))
                    {
                        if (((DateSystem.IsWeekend(lastDay, CountryCode.UA) == false)))
                        {
                            cells[vsS[currentColumn + previousDays] + startWith].Value = 7;
                        }//????
                        else { cells[vsS[currentColumn + previousDays] + startWith].Value = 7;// "ВC"; 
                        }
                    }

                    else
                    {
                        if (((DateSystem.IsWeekend(lastDayqq, CountryCode.UA) == true)))
                        {
                            cells[vsS[currentColumn + previousDays] + startWith].Value = "ВC";
                        }
                        else { cells[vsS[currentColumn + previousDays] + startWith].Value = 8; }

                        n = previousDays;

                    }

                }

                Сells[source.Count + 8, days - 2].Value = "усього:";
                //sum days everybody of working of employees
                cells[source.Count + 8, days + 4].Formula = "=SUM(" + vsS[(days + 4)] + 8 + ":" + vsS[(days + 4)] + (source.Count + 8) + ")"; //строка,столбец

                //sum weekend days by employees
                cells[source.Count + 8, days + 5].Formula = "=IF(SUM(" + vsS[(days + 5)] + 8 + ":" + vsS[(days + 5)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 5)] + 8 + ":" + vsS[(days + 5)] + (source.Count + 8) + "))"; //строка,столбец

                //sum annual vacation by employees
                cells[source.Count + 8, days + 6].Formula = "=IF(SUM(" + vsS[(days + 6)] + 8 + ":" + vsS[(days + 6)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 6)] + 8 + ":" + vsS[(days + 6)] + (source.Count + 8) + "))"; //строка,столбец

                //------------------
                cells[source.Count + 8, days + 7].Formula = "=IF(SUM(" + vsS[(days + 7)] + 8 + ":" + vsS[(days + 7)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 7)] + 8 + ":" + vsS[(days + 7)] + (source.Count + 8) + "))"; //строка,столбец
                //------------------


                //sum temporary disability (sick)
                cells[source.Count + 8, days + 8].Formula = "=IF(SUM(" + vsS[(days + 8)] + 8 + ":" + vsS[(days + 8)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 8)] + 8 + ":" + vsS[(days + 8)] + (source.Count + 8) + "))"; //строка,столбец

                //sum days unclear without payment
                cells[source.Count + 8, days + 10].Formula = "=IF(SUM(" + vsS[(days + 10)] + 8 + ":" + vsS[(days + 10)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 10)] + 8 + ":" + vsS[(days + 10)] + (source.Count + 8) + "))"; //строка,столбец

                //sum days with initiative administration
                cells[source.Count + 8, days + 11].Formula = "=IF(SUM(" + vsS[(days + 11)] + 8 + ":" + vsS[(days + 11)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 11)] + 8 + ":" + vsS[(days + 11)] + (source.Count + 8) + "))"; //строка,столбец

                //sum assignment 
                cells[source.Count + 8, days + 12].Formula = "=IF(SUM(" + vsS[(days + 12)] + 8 + ":" + vsS[(days + 12)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 12)] + 8 + ":" + vsS[(days + 12)] + (source.Count + 8) + "))"; //строка,столбец

                //sum others reasons
                cells[source.Count + 8, days + 14].Formula = "=IF(SUM(" + vsS[(days + 14)] + 8 + ":" + vsS[(days + 14)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 14)] + 8 + ":" + vsS[(days + 14)] + (source.Count + 8) + "))"; //строка,столбец

                //sum part week
                cells[source.Count + 8, days + 15].Formula = "=IF(SUM(" + vsS[(days + 15)] + 8 + ":" + vsS[(days + 15)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 15)] + 8 + ":" + vsS[(days + 15)] + (source.Count + 8) + "))"; //строка,столбец


                //sum holiday and weekend everybody of working of employees
                cells[source.Count + 8, days + 16].Formula = "=IF(SUM(" + vsS[(days + 16)] + 8 + ":" + vsS[(days + 16)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 16)] + 8 + ":" + vsS[(days + 16)] + (source.Count + 8) + "))"; //строка,столбец


                //sum days in mounth everybody of working of employees
                cells[source.Count + 8, days + 17].Formula = "=SUM(" + vsS[(days + 17)] + 8 + ":" + vsS[(days + 17)] + (source.Count + 8) + ")"; //строка,столбец

                //sum ALL hours worked by employees
                cells[source.Count + 8, days + 18].Formula = "=SUM(" + vsS[(days + 18)] + 8 + ":" + vsS[(days + 18)] + (source.Count + 8) + ")"; //строка,столбец

                ////sum worked out time
                //cells[source.Count + 8, days + 19].Formula = "=IF(SUM(" + vsS[(days + 19)] + 8 + ":" + vsS[(days + 19)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 19)] + 8 + ":" + vsS[(days + 19)] + (source.Count + 8) + "))"; //строка,столбец

                ////sum overdue time
                //cells[source.Count + 8, days + 20].Formula = "=IF(SUM(" + vsS[(days + 20)] + 8 + ":" + vsS[(days + 20)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 20)] + 8 + ":" + vsS[(days + 20)] + (source.Count + 8) + "))"; //строка,столбец

                ////sum worked night time
                //cells[source.Count + 8, days + 21].Formula = "=IF(SUM(" + vsS[(days + 21)] + 8 + ":" + vsS[(days + 21)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 21)] + 8 + ":" + vsS[(days + 21)] + (source.Count + 8) + "))"; //строка,столбец

                //sum worked part  time work days   "=IF(SUM(" + vsS[(days + 16)] + 8 + ":" + vsS[(days + 16)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 16)] + 8 + ":" + vsS[(days + 16)] + (source.Count + 8) + "))";
                cells[source.Count + 9, days + 4].Formula = "=IF(SUM(" + vsS[(days + 26)] + 8 + ":" + vsS[(days + 26)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 26)] + 8 + ":" + vsS[(days + 26)] + (source.Count + 8) + "))"; //строка,столбец

                //sum worked part-time weekend
                cells[source.Count + 9, days + 5].Formula = "=IF(SUM(" + vsS[(days + 27)] + 8 + ":" + vsS[(days + 27)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 27)] + 8 + ":" + vsS[(days + 27)] + (source.Count + 8) +"))"; //строка,столбец

                //sum worked part-time weekend and holiday
                cells[source.Count + 9, days + 16].Formula ="=IF(SUM(" + vsS[(days + 28)] + 8 + ":" + vsS[(days + 28)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 28)] + 8 + ":" + vsS[(days + 28)] + (source.Count + 8) +"))"; //строка,столбец

                //sum sum worked part-time all days
                cells[source.Count + 9, days + 17].Formula ="=IF(SUM(" + vsS[(days + 29)] + 8 + ":" + vsS[(days + 29)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 29)] + 8 + ":" + vsS[(days + 29)] + (source.Count + 8) +"))"; //строка,столбец

                //sum sum worked part-time all time
                cells[source.Count + 9, days + 18].Formula ="=IF(SUM(" + vsS[(days + 30)] + 8 + ":" + vsS[(days + 30)] + (source.Count + 8) + ")=0," + "\"\"" + "," + "SUM(" + vsS[(days + 30)] + 8 + ":" + vsS[(days + 30)] + (source.Count + 8) +"))"; //строка,столбец

                Сells[source.Count + 9, days - 2].Value = "за сумісництвом:";
                //sum days everybody of working of employees
                //cells[source.Count + 9, days + 4].Formula = "=SUM(" + vsS[(days + 4)] + 8 + ":" + vsS[(days + 4)] + (source.Count + 8) + ")"; //строка,столбец

                Сells[source.Count + 10, 0].Value = "Вихідні і свята";
                Сells[source.Count + 11, 0].Value = "Щорічна відпустка";
                Сells[source.Count + 12, 0].Value = "Скорочений тиждень";
                Сells[source.Count + 13, 0].Value = "Відпустка без";
                Сells[source.Count + 14, 0].Value = "збереження з/п за";
                Сells[source.Count + 15, 0].Value = "згодою сторін";
                Сells[source.Count + 16, 0].Value = "Начальник Відділу";

                cells[source.Count + 16, 0].Font.Bold = true;
                Сells[source.Count + 10, 0].HorizontalAlignment = HAlign.Right;
                Сells[source.Count + 11, 0].HorizontalAlignment = HAlign.Right;
                Сells[source.Count + 12, 0].HorizontalAlignment = HAlign.Right;
                Сells[source.Count + 13, 0].HorizontalAlignment = HAlign.Right;
                Сells[source.Count + 14, 0].HorizontalAlignment = HAlign.Right;
                Сells[source.Count + 15, 0].HorizontalAlignment = HAlign.Right;
                Сells[source.Count + 16, 5].Borders[BordersIndex.EdgeBottom].LineStyle = LineStyle.Continous;
                Сells[source.Count + 16, 6].Borders[BordersIndex.EdgeBottom].LineStyle = LineStyle.Continous;
                Сells[source.Count + 16, 7].Borders[BordersIndex.EdgeBottom].LineStyle = LineStyle.Continous;
                Сells[source.Count + 16, 8].Borders[BordersIndex.EdgeBottom].LineStyle = LineStyle.Continous;

                Сells[source.Count + 10, 1].Value = "ВC";
                Сells[source.Count + 11, 1].Value = "В";
                Сells[source.Count + 12, 1].Value = "СТ";
                Сells[source.Count + 15, 1].Value = "НА";

                Сells[source.Count + 10, 3].Value = "Тимчасова непрацездатність";
                Сells[source.Count + 11, 3].Value = "Відпустка у зв'язку з пологами";
                Сells[source.Count + 12, 3].Value = "Відпустка для догляду за дитиною";

                Сells[source.Count + 10, 15].Value = "ТН";
                Сells[source.Count + 11, 15].Value = "ВП";
                Сells[source.Count + 12, 15].Value = "ДД";

                Сells[source.Count + 10, 20].Value = "Учбова відпустка";
                Сells[source.Count + 11, 20].Value = "Відрядження";
                Сells[source.Count + 12, 20].Value = "За сумісництвом";

                Сells[source.Count + 10, 27].Value = "УВ";
                Сells[source.Count + 11, 27].Value = "ВД";
                Сells[source.Count + 12, 27].Value = "С";

                Сells[source.Count + 10, 35].Value = "Неявки з ініциативи адміністрації";
                Сells[source.Count + 11, 35].Value = "Виконання держобов'язків";
                Сells[source.Count + 12, 35].Value = "Простій";

                Сells[source.Count + 10, 45].Value = "НД";
                Сells[source.Count + 11, 45].Value = "Д";
                Сells[source.Count + 12, 45].Value = "П";
                cells["" + (startWith + source.Count) + ":" + (startWith + source.Count)].Delete();

            }
       //      cells["" + (startWith + source.Count) + ":" + (startWith + source.Count)].Insert();

            try
            {
                for (int i = 0; i < source.Count; i++)
                {
                    nameDepart = source[i].DepartmentName;
                }
                // string homePath = Utils.HomePath + @"\\SERVER-TFS\Data\Табель обліку робочого часу\" + nameDepart + "\\";
               // var pathToSaveFile = new DirectoryInfo(@"\\SERVER-TFS\Data\Табель обліку робочого часу\" + nameDepart + "\\");



                //Workbook.SaveAs(homePath + "Табель обліку робочого часу  за " + ToMonthName(currentDate) + " " + nameDepart + ".xls", FileFormat.Excel8);

                //Process process = new Process();
                //process.StartInfo.Arguments = "\"" + homePath + "Табель обліку робочого часу  за " + ToMonthName(currentDate) + " " + nameDepart + ".xls" + "\"";

                //process.StartInfo.FileName = "Excel.exe";
                //process.Start();


                string pathhhh = Utils.HomePath + @"\D:\";

                string path = "";
                string subpath = "";
                path = Directory.Exists(@"D:\") ? @"D:\" : @"E:\";



                string allpath = path + subpath + "\"";//@"D:\TimeSheet_" + currentDate.Year
                Workbook.SaveAs(path + "Табель обліку робочого часу  за " + ToMonthName(currentDate) + " " + nameDepart + ".xls", FileFormat.Excel8);

                Process process = new Process();
                process.StartInfo.Arguments = "\"" + path + "Табель обліку робочого часу  за " + ToMonthName(currentDate) + " " + nameDepart + ".xls" + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();


            }
            catch (System.IO.IOException) { MessageBox.Show("Документ уже открыт!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не найден Microsoft Excel!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        public string ToMonthName( DateTime dateTime)
        {
            return CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(dateTime.Month);
        }


//------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        #endregion

        #region FixedAssetsOrder report's

        public void PrintFixedAssetsOder(FixedAssetsOrderJournalDTO model, List<FixedAssetsMaterialsDTO> materialsListSource, DateTime endDate, DateTime firstDay)
        {
            List<FixedAssetsMaterialsDTO> materialsList = new List<FixedAssetsMaterialsDTO>();
            string typeMaterial = "";
            string templateName = " ";
            templateName = @"\Templates\FixedAssetsPrintItem.xls";

            var Workbook = Factory.GetWorkbook(GeneratedReportsDir + templateName);
            var Worksheet = Workbook.Worksheets[0];
            var Сells = Worksheet.Cells;            
            IRange cells = Worksheet.Cells;
            int startRow = 3;
            int indexRow = startRow + 1;
            string indexRowStr;

            if (model.Id == null)
            {
                MessageBox.Show("За обраний період немає даних!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //Head document
            cells["B2"].Value = "Карточка основного засобу за період: з " + firstDay.ToShortDateString() + " по " + endDate.ToShortDateString();
            cells["B2"].Font.Size = 14;
            cells["B2"].Font.Bold = true;
            cells["B2"].HorizontalAlignment = SpreadsheetGear.HAlign.Center;
            cells["B2"].VerticalAlignment = SpreadsheetGear.VAlign.Center;
            cells["B2:" + "P2"].Merge();

            //Head table1          
            indexRowStr = indexRow.ToString();
            //ColumnWidth
            Worksheet.Cells["B:B"].ColumnWidth = 12.47;
            Worksheet.Cells["C:C"].ColumnWidth = 32.9;
            Worksheet.Cells["D:D"].ColumnWidth = 16.19;
            Worksheet.Cells["E:E"].ColumnWidth = 26.19;
            Worksheet.Cells["F:F"].ColumnWidth = 14.86;
            Worksheet.Cells["G:G"].ColumnWidth = 13.04;
            Worksheet.Cells["H:H"].ColumnWidth = 11.43;
            Worksheet.Cells["I:I"].ColumnWidth = 25.43;

            Worksheet.Cells["J:J"].ColumnWidth = 16.14;
            Worksheet.Cells["K:K"].ColumnWidth = 12.33;
            Worksheet.Cells["L:L"].ColumnWidth = 13.9;
            Worksheet.Cells["M:M"].ColumnWidth = 14.04;
            Worksheet.Cells["N:N"].ColumnWidth = 24.43;
            Worksheet.Cells["O:O"].ColumnWidth = 24.43;
            Worksheet.Cells["P:P"].ColumnWidth = 12.33;
   //         int rowcount = materialsListSource.Count;

            //TITLE
            cells["B" + indexRowStr].Value = "Інвентарний номер";
            cells["C" + indexRowStr].Value = "Найменування";
            cells["D" + indexRowStr].Value = "Бал./рах.";
            cells["E" + indexRowStr].Value = "Відповідальна особа";
            cells["F" + indexRowStr].Value = "Термін використання (міс.)";
            cells["G" + indexRowStr].Value = "Дата приняття до обліку";
            cells["H" + indexRowStr].Value = "Дата зняття з обліку";
            cells["I" + indexRowStr].Value = "Група";
            cells["J" + indexRowStr].Value = "Первинна вартість";
            cells["K" + indexRowStr].Value = "Збільшення вартості";
            cells["L" + indexRowStr].Value = "Поточна вартість";
            cells["M" + indexRowStr].Value = "Залишкова вартість";
            cells["N" + indexRowStr].Value = "Сума амортизації";
            cells["O" + indexRowStr].Value = "Амортизація за місяць";

            //body table
            indexRow++;
            indexRowStr = indexRow.ToString();
            cells["B" + indexRowStr].Value = model.InventoryNumber;
            cells["C" + indexRowStr].Value = model.InventoryName;
            cells["D" + indexRowStr].Value = model.BalanceAccountNum;
            cells["E" + indexRowStr].Value = model.SupplierName;
            cells["F" + indexRowStr].Value = model.UsefulMonth;
            cells["G" + indexRowStr].Value = model.BeginDate;
            cells["H" + indexRowStr].Value = model.EndRecordDate;
            cells["I" + indexRowStr].Value = model.GroupName;
            cells["J" + indexRowStr].Value = model.BeginPrice;
            cells["K" + indexRowStr].Value = model.IncreasePrice;
            cells["L" + indexRowStr].Value = model.TotalPrice;
            cells["M" + indexRowStr].Value = model.CurrentPrice;
            cells["N" + indexRowStr].Value = model.PeriodAmortization;
            cells["O" + indexRowStr].Value = model.CurrentAmortization;
            cells["I" + indexRowStr + ":" + "O" + indexRowStr].NumberFormat = "### ### ##0.00";
            cells["L" + indexRowStr + ":" + "M" + indexRowStr].NumberFormat = "### ### ##0.00";
            // first row headtable
            cells["B" + (startRow + 1) + ":" + "P" + (startRow + 1)].WrapText = true;
            cells["B" + (startRow + 1) + ":" + "P" + (startRow + 1)].HorizontalAlignment = SpreadsheetGear.HAlign.Center;
            cells["B" + (startRow + 1) + ":" + "P" + (startRow + 1)].VerticalAlignment = SpreadsheetGear.VAlign.Center;
            cells["B" + (startRow + 1) + ":" + "P" + indexRow].Borders.LineStyle = LineStyle.Continous;
            // first row headtable
            cells["B" + (startRow + 1) + ":" + "P" + (startRow + 1)].WrapText = true;
            cells["B" + (startRow + 1) + ":" + "P" + (startRow + 1)].HorizontalAlignment = SpreadsheetGear.HAlign.Center;
            cells["B" + (startRow + 1) + ":" + "P" + (startRow + 1)].VerticalAlignment = SpreadsheetGear.VAlign.Center;
            cells["B" + (startRow + 1) + ":" + "P" + indexRow].Borders.LineStyle = LineStyle.Continous;

            //table 2
            indexRow++;
            startRow = indexRow++;
                indexRowStr = indexRow.ToString();
                cells["B" + indexRowStr].Value = "Ном. номер";
                cells["C" + indexRowStr].Value = "Найменування";
                cells["D" + indexRowStr].Value = "Рах. нарахування амортизації";
                cells["E" + indexRowStr].Value = "Номер надходження";
                cells["F" + indexRowStr].Value = "Дата надходження";
                cells["G" + indexRowStr].Value = "Балансовий рахунок";
                cells["H" + indexRowStr].Value = "К-сть";
                cells["I" + indexRowStr].Value = "Ціна";
                cells["J" + indexRowStr].Value = "Сума";
                cells["K" + indexRowStr].Value = "Дата списання";
                cells["L" + indexRowStr].Value = "Сума списання";
                cells["M" + indexRowStr].Value = "Сумма до обліку";
                cells["N" + indexRowStr].Value = "Тип";
                
                //body table 2
                for (var i = 0; i < materialsListSource.Count; i++)
                {
                    indexRow++;
                    indexRowStr = indexRow.ToString();

                    cells["B" + indexRowStr].Value = materialsListSource[i].Nomenclature;//((FixedAssetsMaterialsDTO)fixedAssetsOrderBS[i]).Nomenclature;
                    cells["C" + indexRowStr].Value = materialsListSource[i].Name;
                    cells["D" + indexRowStr].Value = materialsListSource[i].FixedNum;
                    cells["E" + indexRowStr].Value = materialsListSource[i].ReceiptNum;
                    cells["F" + indexRowStr].Value = materialsListSource[i].OrderDate;
                    cells["G" + indexRowStr].Value = materialsListSource[i].OrderNum;
                    cells["H" + indexRowStr].Value = materialsListSource[i].Quantity;
                    cells["I" + indexRowStr].Value = materialsListSource[i].UnitPrice;
                    cells["J" + indexRowStr].Value = materialsListSource[i].TotalPrice;
                    cells["K" + indexRowStr].Value = materialsListSource[i].ExpDate;
                    cells["L" + indexRowStr].Value = materialsListSource[i].Price;
                    cells["M" + indexRowStr].Value = materialsListSource[i].FixedPrice;
                    switch(materialsListSource[i].Flag)
                    {
                        case 0:
                    typeMaterial = "Основний засіб";
                    break;
                        case 1:
                    typeMaterial = "Збільшення вартості";
                    break;
                        case 2:
                    typeMaterial = "Корегування";
                    break;
                        default: typeMaterial = "";
                    break;
                    }
                    cells["N" + indexRowStr].Value = typeMaterial;

                    //Interval I->J
                    cells["I" + indexRowStr + ":" + "J" + indexRowStr].NumberFormat = "### ### ##0.00";
                    cells["L" + indexRowStr + ":" + "M" + indexRowStr].NumberFormat = "### ### ##0.00";
                }

                cells["B" + (startRow + 1) + ":" + "N" + (startRow + 1)].WrapText = true;
                cells["B" + (startRow + 1) + ":" + "N" + (startRow + 1)].HorizontalAlignment = SpreadsheetGear.HAlign.Center;
                cells["B" + (startRow + 1) + ":" + "N" + (startRow + 1)].VerticalAlignment = SpreadsheetGear.VAlign.Center;
                cells["B" + (startRow + 1) + ":" + "N" + indexRow].Borders.LineStyle = LineStyle.Continous;
                indexRow = indexRow + 2;
                startRow = startRow + materialsListSource.Count + 3;
                indexRowStr = indexRow.ToString();
            try
            {
                Workbook.SaveAs(GeneratedReportsDir + "Карточка ОЗ " + model.InventoryNumber + ".xls", FileFormat.Excel8);                Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Карточка ОЗ " +model.InventoryNumber+ ".xls";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ вже відкритий!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        public void PrintInventoryCardForSoftware(FixedAssetsOrderJournalDTO model)
        {
            string templateName = " ";
            templateName = @"\Templates\InventoryCardPrintForSoftware.xls";
            var Workbook = Factory.GetWorkbook(GeneratedReportsDir + templateName);
            var Worksheet = Workbook.Worksheets[0];
            var Сells = Worksheet.Cells;

            IRange cells = Worksheet.Cells;
            decimal yearAmortization = Math.Round(((decimal)model.BeginPrice / (short)model.UsefulMonth) * 12, 2);               

            int rows = 40;
            int cols = 30;

            cells[0, 0, rows, cols].Replace("{currInventoryName}", model.InventoryName, LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currInventoryNumber}", model.InventoryNumber, LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currBeginDate}", Convert.ToDateTime(model.BeginDate).ToShortDateString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currSupplier_Name}", model.SupplierName, LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currUsefulMonth}", model.UsefulMonth.ToString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currBalance_Account_Num}", model.BalanceAccountNum, LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currBeginPrice}", model.BeginPrice.ToString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currAmortizationYearPrice}", yearAmortization.ToString(), LookAt.Part, SearchOrder.ByRows, false);

            try
            {
                Worksheet.SaveAs(GeneratedReportsDir + "Інвентарна картка обліку ПО №" + model.InventoryNumber.Replace("/", "_") + ".xls", FileFormat.Excel8);

                Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Інвентарна картка обліку ПО №" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls" + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ вже відкритий!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не знайдений Microsoft Excel!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }              
        }
        
        public void PrintFixedAssetsOrderAct(FixedAssetsOrderJournalDTO model, List<FixedAssetsMaterialsDTO> materialsListSource)
        {
            string templateName = " ";
            templateName = @"\Templates\FixedAssetsOrderAct.xls";
            var Workbook = Factory.GetWorkbook(GeneratedReportsDir + templateName);
            var Worksheet = Workbook.Worksheets[0];
            var Сells = Worksheet.Cells;
            IRange cells = Worksheet.Cells;

            List<FixedAssetsMaterialsDTO> newMaterialsList = new List<FixedAssetsMaterialsDTO>();
            decimal sumPrice=0;
           
            int rows = 120;
            int cols = 31;
            DateTime dt = DateTime.Now;

            cells[1, 1, rows, cols].Replace("{currYear}", dt.Year.ToString(), LookAt.Part, SearchOrder.ByRows, false);

            if (materialsListSource.Count > 0)
            {
                newMaterialsList = materialsListSource.Where(x => x.Flag < 2).ToList();
                sumPrice = Math.Round(newMaterialsList.Where(r => r.Flag < 2).Select(s => s.FixedPrice).Sum(), 2);
                float percentUsefullMonth = (100 / (Convert.ToInt16(model.UsefulMonth) / 12));

                cells["M32"].Value = model.InventoryName.ToString() + "   ";
                cells["O27"].Value = model.InventoryNumber;
                cells["D27"].Value = model.BalanceAccountNum;
                cells["R27"].Value = model.FixedAccountNum;
                cells["V27"].Value = percentUsefullMonth.ToString() + "%";

                cells["N27"].Value = sumPrice;
                cells["N27"].NumberFormat = "### ### ##0.00";
                cells["C27"].Value = model.RegionName;

                int materialRows = 44;
                int count = newMaterialsList.Count;
                cells[materialRows + ":" + (materialRows + count)].Insert();
                cells["A43:AC43"].Copy(cells["A" + materialRows + ":AC" + (materialRows + count - 1)], PasteType.Formats, PasteOperation.None, false, false);

                foreach (var item in newMaterialsList)
                {
                    cells["B" + materialRows].Value = item.Name;
                    cells["G" + materialRows].Value = item.Nomenclature;
                    cells["M" + materialRows].Value = "";
                    cells["Q" + materialRows].NumberFormat = "YYYY";
                    cells["Q" + materialRows].Value = item.ExpDate;
                    cells["T" + materialRows].NumberFormat = "MM.YYYY";
                    cells["T" + materialRows].Value = item.ExpDate;
                    cells["Y" + materialRows].Value = "";
                    cells["B" + materialRows].RowHeight = 30;
                    materialRows++;
                };
            }

            try
            {
                Worksheet.SaveAs(GeneratedReportsDir + "Акт прийому-передачі інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls", FileFormat.Excel8);

                Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Акт прийому-передачі інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls" + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ вже відкритий!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не знайдений Microsoft Excel!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            
        }

        public void PrintFixedAssetsOrderActForSoftware(FixedAssetsOrderJournalDTO model)
        {
            string templateName = " ";
            templateName = @"\Templates\FixedAssetsOrderActForSoftware.xls";
            var Workbook = Factory.GetWorkbook(GeneratedReportsDir + templateName);
            var Worksheet = Workbook.Worksheets[0];
            var Сells = Worksheet.Cells;
            IRange cells = Worksheet.Cells;

            decimal yearAmortization = Math.Round(((decimal)model.BeginPrice / (short)model.UsefulMonth) * 12, 2);

            int rows = 100;
            int cols = 30;
            DateTime dt = DateTime.Now;

            cells[0, 0, rows, cols].Replace("{currYear}", dt.Year.ToString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currInventoryName}", model.InventoryName.ToString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currInventoryNumber}", model.InventoryNumber.ToString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currBeginDate}", Convert.ToDateTime(model.BeginDate).ToShortDateString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currSupplier_Name}", model.SupplierName.ToString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currUsefulMonth}", model.UsefulMonth.ToString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currBalance_Account_Num}", model.BalanceAccountNum.ToString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currBeginPrice}", model.BeginPrice.ToString(), LookAt.Part, SearchOrder.ByRows, false);
            cells[0, 0, rows, cols].Replace("{currAmortizationYearPrice}", yearAmortization.ToString(), LookAt.Part, SearchOrder.ByRows, false);

            try
            {
                Worksheet.SaveAs(GeneratedReportsDir + "Акт введення в господарський оборот об'єкта № " + model.InventoryNumber.ToString().Replace("/", "_") + ".xls", FileFormat.Excel8);
                Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Акт введення в господарський оборот об'єкта № " + model.InventoryNumber.ToString().Replace("/", "_") + ".xls" + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ вже відкритий!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не знайдений Microsoft Excel!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        public void PrintFixedAssetsOrderActWriteOff(FixedAssetsOrderJournalDTO model, List<FixedAssetsMaterialsDTO> materialsListSource, int monthSource, int yearSource)
        {
            DateTime expDateMaterial = new DateTime();
            int yearExpDateMaterial, monthExpDateMaterial;
            string templateName = " ";
            List<FixedAssetsMaterialsDTO> newMaterialsList = new List<FixedAssetsMaterialsDTO>();

            templateName = @"\Templates\FixedAssetsAddedPriceTemplate.xls";
            var Workbook = Factory.GetWorkbook(GeneratedReportsDir + templateName);
            var Worksheet = Workbook.Worksheets[0];
            var Сells = Worksheet.Cells;
            IRange cells = Worksheet.Cells;

            if (materialsListSource.Count > 0)
            {
                for (int i = 0; i < materialsListSource.Count; i++)
                {
                    expDateMaterial = materialsListSource[i].ExpDate.Value;
                    yearExpDateMaterial = expDateMaterial.Year;
                    monthExpDateMaterial = expDateMaterial.Month;
                    //from grid Mateials find ExpDate=date edit'ov on main form 
                    if (materialsListSource[i].Flag == 1 && monthExpDateMaterial == monthSource && yearExpDateMaterial == yearSource)
                        newMaterialsList.Add(materialsListSource[i]);
                }
                int materialRows = 21;
                if (newMaterialsList.Count != 0)
                {
                    foreach (var item in newMaterialsList)
                    {
                        cells["" + materialRows + ":" + materialRows].Insert();
                        cells["A" + materialRows].HorizontalAlignment = HAlign.Left;
                        cells["A" + materialRows + ":" + "B" + materialRows].Merge();
                        cells["A" + materialRows].Value = model.RegionName;
                        cells["C" + materialRows + ":" + "D" + materialRows].Merge();
                        cells["C" + materialRows].Value = model.BalanceAccountNum;

                        switch (model.GroupId)
                        {
                            case 2:
                            case 10:
                                cells["G" + materialRows].Value = "154";
                                break;
                            default: cells["G" + materialRows].Value = "152";
                                break;
                        }
                        cells["K" + materialRows + ":" + "L" + materialRows].Merge();
                        cells["K" + materialRows].Value = item.TotalPrice;
                        cells["M" + materialRows].Value = model.InventoryNumber;
                        materialRows++;
                    };                    
                }
                else
                {
                    MessageBox.Show("Відсутні матеріали, які задовольняють задані умови", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return; 
                }
            }

            try
            {
                Worksheet.SaveAs(GeneratedReportsDir + "Акт приймання-здачі відремонтованих об'єктів інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls", FileFormat.Excel8);

                Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Акт приймання-здачі відремонтованих об'єктів інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls" + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ вже відкритий!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не знайдений Microsoft Excel!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        
        }

        public string GetFixedAssetsGroupByAccountNumber(string accountNumber)
        {
            switch (accountNumber)
            {
                case "103":
                    return "3";
                    break;
                case "104":
                case "104/1":
                    return "4";
                    break;
                case "105":
                    return "5";
                    break;
                case "106":
                case "127":
                    return "6";
                    break;
                case "109":
                    return "9";
                    break;
                default:
                    return "Невідома группа";
                    break;
            }
        }

        public void FixedAssetsDecreeInput(FixedAssetsOrderRegJournalDTO model, FixedAssetsMaterialsDTO modelMaterials)
        {
            //DataTable currTable = null;
            var pad = new Ua();
            string templateName = @"\Templates\FixedAssetsDecreeInputTemplate.xls";
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(GeneratedReportsDir +templateName);
            SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets[0];
            SpreadsheetGear.IRange cells = worksheet.Cells;

            int rows = 120;
            int cols = 31;
            DateTime dt = DateTime.Now;
            DateTime amortizationDate = new DateTime(((DateTime)model.DateOrder).Year, ((DateTime)model.DateOrder).Month, 1);
            amortizationDate = amortizationDate.AddMonths(1);

            // Add in cell date, month, year
            cells["B6"].HorizontalAlignment = HAlign.Center;
            cells["D6"].HorizontalAlignment = HAlign.Center;
            cells["F6"].HorizontalAlignment = HAlign.Center;
            string s = ((DateTime)model.DateOrder).Month.ToString();
            int month = Int32.Parse(s);

           string rez = RuDateAndMoneyConverter.MonthName(month, Utils.TextCase.Genitive);

            cells["B" + 6].Value = ((DateTime)model.DateOrder).Day.ToString();
            cells["D" + 6].Value = rez;
            cells["F" + 6].Value = ((DateTime)model.DateOrder).Year.ToString();
            cells["K" + 6].Value = model.NumberOrder;
            double years = Convert.ToDouble(model.UsefulMonth) / 12.0;

            cells["A" + 8].HorizontalAlignment = HAlign.Left;

            for (int i = 13; i < 26; i++)
            {
                cells["B" + i].HorizontalAlignment = HAlign.Left;
                cells["A" + i].HorizontalAlignment = HAlign.Center;
            }

            cells["B14"].Font.Bold = true;
            cells["B" + 13].HorizontalAlignment = HAlign.Left;
            cells["A" + 8].Value = "« Про введення в експлуатацію " + model.InventoryName + " Інв. " + model.InventoryNumber + " »";
            //SourseData["InventoryName"] + ;
            cells["A" + 13].Value = " 1" + " . ";
            cells["B" + 13].Value = " Ввести в експлуатацію з " + ((DateTime)model.DateOrder).ToShortDateString() + "р.";
            cells["B" + 14].Value = " " + model.InventoryName;
            cells["B" + 15].Value = " Вартістю " + modelMaterials.TotalPrice + " грн. в кількості 1 шт.";
            cells["B" + 16].Value = " та надати наступний інвентарний номер: №" + model.InventoryNumber + ".";
            cells["A" + 18].Value = " 2 . ";
            cells["B" + 18].Value = " Для цілей бухгалтерського обліку використати рахунок №" + model.BalanceAccountNum + ", ";
            cells["B" + 19].Value = " податкового обліку група №" + GetFixedAssetsGroupByAccountNumber(model.BalanceAccountNum.ToString()) + ",";
            cells["B" + 20].Value = " бухгалтерії вести облік вищевказаного основного засобу,";
            cells["B" + 21].Value = " амортизацію нарахувати з " + amortizationDate.ToShortDateString() + " р.";
            cells["A" + 23].Value = " 3 . ";
            cells["B" + 23].Value = " Термін корисного використання " + years + " " +NumberToYear(Convert.ToInt32(years)) + ".";
            cells["A" + 25].Value = " 4 . ";
            cells["B" + 25].Value = " Контроль за виконанням цього наказу покладаю  ";
            cells["B" + 26].Value = " на Першого заступника директора - Кондрашова В.В.. ";
            try
            {
                worksheet.SaveAs(GeneratedReportsDir + "Наказ № 00-00-00-інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls", FileFormat.Excel8);
                Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Наказ № 00-00-00-інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls" + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ уже открыт!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не найден Microsoft Excel!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
            }       
        }

        public static string NumberToYear(int number)
        {
            string year = "";
            switch (number)
            {
                case 1:
                    year = "рік";
                    break;
                case 2:
                case 3:
                case 4:
                    year = "роки";
                    break;
                case 5:
                case 6:
                case 7:
                case 8:
                case 9:
                case 10:
                case 11:
                case 12:
                case 13:
                case 14:
                case 15:
                case 16:
                case 17:
                case 18:
                case 19:
                case 20:
                    year = "років";
                    break;
                default:
                    break;

            }
            return year;
        }


        public void FixedAssetsDecreeAddedPrice(FixedAssetsOrderRegJournalDTO model, FixedAssetsMaterialsDTO modelMaterials)
        {
            var pad = new Ua();
            string templateName = @"\Templates\FixedAssetsDecreeAddedPriceTemplate.xls";
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(GeneratedReportsDir + templateName);
            SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets[0];
            SpreadsheetGear.IRange cells = worksheet.Cells;

            DateTime amortizationDate = new DateTime(Convert.ToDateTime(modelMaterials.OrderDate).AddMonths(1).Year, Convert.ToDateTime(modelMaterials.OrderDate).AddMonths(1).Month, 1);
            // Add in cell date, month, year
            cells["B7"].HorizontalAlignment = HAlign.Center;
            cells["D7"].HorizontalAlignment = HAlign.Center;
            cells["F7"].HorizontalAlignment = HAlign.Center;
            string s = ((DateTime)modelMaterials.OrderDate).Month.ToString();
            int month = Int32.Parse(s);
            string rez = RuDateAndMoneyConverter.MonthName(month, Utils.TextCase.Genitive);

            cells["B" + 7].Value = ((DateTime)modelMaterials.ExpDate).Day.ToString();
            cells["D" + 7].Value = rez;
            cells["F" + 7].Value = ((DateTime)modelMaterials.ExpDate).Year.ToString();
            cells["K" + 7].Value = model.NumberOrder;

            var fixedAssetsOrderName = pad.Q(model.InventoryName.ToString());

            double years = Convert.ToDouble(model.UsefulMonth) / 12.0;

            for (int i = 9; i < 27; i++)
            {
                cells["B" + i].HorizontalAlignment = HAlign.Left;
                cells["A" + i].HorizontalAlignment = HAlign.Center;
            }

            cells["A" + 9].HorizontalAlignment = HAlign.Left;
            cells["A" + 9].Value = "« Про збільшення вартості  " + fixedAssetsOrderName[3] + "-" + " інв. №" + model.InventoryNumber + " »";
            cells["A" + 14].Value = " 1 . ";
            cells["B" + 14].Value = "З " + ((DateTime)modelMaterials.OrderDate).ToShortDateString() + "р. збільшити вартість " + fixedAssetsOrderName[3] + "- інв. №" + model.InventoryNumber;
            cells["B" + 16].Value = "На суму ";
            //cells["E" + 16].Value = SourseDataMaterials["UNIT_PRICE"];
            cells["E" + 16].Value = modelMaterials.TotalPrice;

            cells["I" + 16].Value = " грн. ";
            cells["A" + 18].Value = " 2 . ";
            cells["B" + 18].Value = "Для цілей бухгатерського обліку використовувати рахунок № " + model.BalanceAccountNum + ", ";
            cells["B" + 19].Value = "податкового обліку група №" + GetFixedAssetsGroupByAccountNumber(model.BalanceAccountNum.ToString()) + ", ";
            cells["B" + 20].Value = "бугалтерії вести облік вищевказаного засобу, ";
            cells["B" + 21].Value = "амортизацію нарахувати з " + amortizationDate.ToShortDateString() + " р.";
            cells["A" + 23].Value = " 3 . ";
            cells["B" + 23].Value = "Термін корисного використання " + years + " " + NumberToYear(Convert.ToInt32(years)) + ".";//RuDateAndMoneyConverter.NumberToYear(Convert.ToInt32(years)) + " . ";
            cells["A" + 25].Value = " 4 . ";
            cells["B" + 25].Value = "Контроль за виконанням цього наказу покладаю на ";
            cells["B" + 26].Value = "Першого заступника директора Кондрашова В.В.. ";

            try
            {
                worksheet.SaveAs(GeneratedReportsDir + "Наказ на збільшення вартості № 00-00-00-інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls", FileFormat.Excel8);

                Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Наказ на збільшення вартості № 00-00-00-інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls" + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ уже открыт!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не найден Microsoft Excel!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); }


        }

        public void PrintFixedAssetsOrderActForSaleSoftWare(FixedAssetsOrderRegJournalDTO model)
        {
            string templateName = @"\Templates\FixedAssetsOrderActForSaleSoftware.xls";
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(GeneratedReportsDir + templateName);
            SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets[0];
            SpreadsheetGear.IRange cells = worksheet.Cells;
            DateTime dt = DateTime.Now;
            cells["C" + 23].Value = "1";
            cells["D" + 23].Value = model.InventoryName;
            cells["E" + 23].Value = model.InventoryNumber;
            cells["D" + 23].Value = model.InventoryName;
            cells["H" + 23].Value = model.BeginDate;
            cells["M" + 23].Value = model.SupplierName;
            cells["S" + 23].Value = model.UsefulMonth;
            cells["W" + 23].Value = model.BalanceAccountNum;
            cells["AL" + 23].Value = model.DateOrder;
            cells["AN" + 23].Value = model.SoldPrice;
            cells["AQ" + 23].Value = model.TransferPrice;
            cells["AI" + 29].Value = model.SoldPrice;

            try
            {
                worksheet.SaveAs(GeneratedReportsDir + 
                    "Акт вибуття (ліквідації) об'єкта права інтелектуальної власності у складі нематеріальних актів № " +
                    model.InventoryNumber.ToString().Replace("/", "_") + ".xls", FileFormat.Excel8);

                Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + 
                    "Акт вибуття (ліквідації) об'єкта права інтелектуальної власності у складі нематеріальних актів № " +
                    model.InventoryNumber.ToString().Replace("/", "_") + ".xls" + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ уже открыт!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не найден Microsoft Excel!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        public void PrintFixedAssetsDecreeSold(FixedAssetsOrderRegJournalDTO model)
        {
            string templateName = @"\Templates\FixedAssetsDecreeSoldTemplate.xls";
            SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(GeneratedReportsDir + templateName);
            var pad = new Ua();

            SpreadsheetGear.IWorksheet worksheet = workbook.Worksheets[0];
            SpreadsheetGear.IRange cells = worksheet.Cells;

            // Add in cell date, month, year
            cells["B7"].HorizontalAlignment = HAlign.Center;
            cells["D7"].HorizontalAlignment = HAlign.Center;
            cells["F7"].HorizontalAlignment = HAlign.Center;

            string s = ((DateTime)model.DateOrder).Month.ToString();
            int month = Int32.Parse(s);
            string rez = RuDateAndMoneyConverter.MonthName(month, Utils.TextCase.Genitive);
            cells["B" + 7].Value = ((DateTime)model.DateOrder).Day.ToString();
            cells["D" + 7].Value = rez;
            cells["F" + 7].Value = ((DateTime)model.DateOrder).Year.ToString();
            cells["J" + 7].Value = model.NumberOrder;

            for (int i = 13; i < 27; i++)
            {
                cells["B" + i].HorizontalAlignment = HAlign.Left;
                cells["A" + i].HorizontalAlignment = HAlign.Center;
            }
            cells["B" + 15].HorizontalAlignment = HAlign.Center;
            cells["B" + 23].HorizontalAlignment = HAlign.Center;

            cells["A" + 9].HorizontalAlignment = HAlign.Left;
            cells["A" + 9].Value = "« Про продаж замортизованих основних засобів " + model.InventoryName +
                " інвентарний № " + model.InventoryNumber + " »";
            cells["A" + 14].Value = " 1 . ";
            cells["B" + 14].Value = "Продати замортизований " + model.InventoryName + ":";
            cells["B" + 15].Value = "інвентарний № " + model.InventoryNumber + " . ";
            cells["A" + 17].Value = " 2 . ";
            cells["B" + 17].Value = "Відповідальним за продаж призначаю: ";
            cells["B" + 18].Value = "Першого заступника директора - Кондрашова В.В.. ";

            cells["A" + 20].Value = " 3 . ";
            cells["B" + 20].Value = "Головному бухгалтеру Сергієнко Л.В. виконати необхідні ";
            cells["B" + 21].Value = "бухгалтерські операції при продажу та при знятті ";
            cells["B" + 22].Value = "Основних засобів з бухгалтерського обліку:";
            cells["B" + 23].Value = "інвентарний № " + model.InventoryNumber + " . ";
            cells["A" + 25].Value = " 4 . ";
            cells["B" + 25].Value = "Контроль за виконанням цього наказу покладаю на ";
            cells["B" + 26].Value = "Першого заступника директора - Кондрашова В.В.. ";

            try
            {
                worksheet.SaveAs(GeneratedReportsDir + "Наказ на продаж № 00-00-00-інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls", FileFormat.Excel8);

                Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Наказ на продаж № 00-00-00-інв.№" + model.InventoryNumber.ToString().Replace("/", "_") + ".xls" + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ уже открыт!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не найден Microsoft Excel!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning); }

        }


        public void PrintAllJournalFixedAssetsOder(List<FixedAssetsOrderRegJournalDTO> modelList, DateTime beginDate, DateTime endDate)
        {
            string templateName = " ";
            templateName = @"\Templates\TemplateWithStamp.xls";

            var Workbook = Factory.GetWorkbook(GeneratedReportsDir + templateName);
            var Worksheet = Workbook.Worksheets[0];
            var Сells = Worksheet.Cells;
            IRange cells = Worksheet.Cells;
            int startRow = 3;
            int indexRow = startRow + 1;

            cells["D6"].Font.Bold = true;
            cells["A8:G8"].Font.Bold = true;
            cells["A8:G8"].HorizontalAlignment = HAlign.Center;

            cells["A8:G8"].Borders.LineStyle = LineStyle.Continuous;

            cells["A8:f8"].ColumnWidth = 20;
            cells["G8"].ColumnWidth = 28;
            string periodArchiveStr = beginDate.ToShortDateString() + " - " + endDate.ToShortDateString();
            cells["D" + 6].Value = "Журнал регістрації наказів по Основним засобам за " + periodArchiveStr;
            cells["A" + 8].Value = "№ п./п.";
            cells["b" + 8].Value ="Інвентарний номер";
            cells["c" + 8].Value ="Дата";
            cells["d" + 8].Value ="Бал.рах.";
            cells["e" + 8].Value ="Номер наказу	";
            cells["f" + 8].Value ="Зміст";	
            cells["g" + 8].Value ="Тип наказу";
            int col = 9;
            foreach (var item in modelList)
            {
                cells["a" + col].Value = item.Pos;
                cells["b" + col].Value = item.InventoryNumber;
                cells["c" + col].Value = item.DateOrder;
                cells["d" + col].Value = item.BalanceAccountNum;
                cells["e" + col].Value = item.NumberOrder;
                cells["f" + col].Value = item.InventoryName;
                cells["g" + col].Value = item.TypeOrder;

                cells["a"+col].Borders.LineStyle = LineStyle.Continuous;
                cells["b" + col].Borders.LineStyle = LineStyle.Continuous;
                cells["c" + col].Borders.LineStyle = LineStyle.Continuous;
                cells["d" + col].Borders.LineStyle = LineStyle.Continuous;
                cells["e" + col].Borders.LineStyle = LineStyle.Continuous;
                cells["f" + col].Borders.LineStyle = LineStyle.Continuous;
                cells["g" + col].Borders.LineStyle = LineStyle.Continuous;

                col++;
            }

            try
            {
                Workbook.SaveAs(GeneratedReportsDir + "Журнал регістрації наказів по Основним засобам за " + periodArchiveStr + ".xls", FileFormat.Excel8); Process process = new Process();
                process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Журнал регістрації наказів по Основним засобам за " + periodArchiveStr + ".xls";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ вже відкритий!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        #region Report
    
        public void FixedAssetsReportGroupShort(List<FixedAssetsOrderByGroupShortReportDTO> model, DateTime startDate, DateTime endDate)
        {
            if (model.Count == 0)
            {
                MessageBox.Show("За вибраний період немає даних!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                string templateName = " ";
                templateName = @"\Templates\FixedAssetsByGroupShort.xls";
                var Workbook = Factory.GetWorkbook(GeneratedReportsDir + templateName);
                var Worksheet = Workbook.Worksheets[0];
                var cells = Worksheet.Cells;
                String activBalance_Account_Num = String.Empty;
                String activGroupName = String.Empty;

                Dictionary<String, String> mumIndex = new Dictionary<String, String>();
                mumIndex.Add(vsS[1], "18");//23
                mumIndex.Add(vsS[2], "19");//91
                mumIndex.Add(vsS[3], "21");//92
                mumIndex.Add(vsS[4], "23");//93

                StringBuilder sumB = new StringBuilder("=SUM(");
                StringBuilder sumC = new StringBuilder("=SUM(");
                StringBuilder sumD = new StringBuilder("=SUM(");
                StringBuilder sumE = new StringBuilder("=SUM(");
                StringBuilder sumF = new StringBuilder("=SUM(");

                int captionPosition = 6;
                int startRow = captionPosition + 3;
                int activRow = startRow;
                String sumColIndex = "F";

                Action<int> WriteSum = (sendActivRow) =>
                {
                    cells[String.Format("{0}{1}", sumColIndex, sendActivRow)].NumberFormat = "### ### ##0.00";
                    cells[String.Format("{0}{1}", sumColIndex, sendActivRow)].Value = "=SUM(B" + sendActivRow + ":E" + sendActivRow + ")";
                };

                Action<int, int, int, string, Boolean, Boolean> WriteColumSum = (sendActivRow, sendStartActivSumRowIndex, sendEndActivSumRowIndex, numName, isLastSum, isNotGlobalSum) =>
                {
                    string activColumName = "B";
                    string activAdressName = string.Empty;
                    //A
                    activColumName = "A";
                    cells[String.Format("{0}{1}", activColumName, sendActivRow)].Value = isNotGlobalSum ? String.Format("{0} по рахунку {1}", "Всього", numName) : "Всього";
                    cells[String.Format("{1}{0}:{2}{0}", sendActivRow, activColumName, "F")].Interior.Color = isNotGlobalSum ? Color.LightGreen : Color.Silver;
                    //B
                    activColumName = "B";
                    activAdressName = String.Format("{0}{1}", activColumName, sendActivRow);
                    cells[activAdressName].Value = isNotGlobalSum ? String.Format("=SUM({0}{1}:{0}{2})", activColumName, sendStartActivSumRowIndex, sendEndActivSumRowIndex) : sumB.ToString();
                    cells[activAdressName].NumberFormat = "### ### ##0.00";
                    sumB.Append(String.Format("{0}{1}", activAdressName, isLastSum ? ")" : "+"));
                    //C
                    activColumName = "C";
                    activAdressName = String.Format("{0}{1}", activColumName, sendActivRow);
                    cells[activAdressName].Value = isNotGlobalSum ? String.Format("=SUM({0}{1}:{0}{2})", activColumName, sendStartActivSumRowIndex, sendEndActivSumRowIndex) : sumC.ToString();
                    cells[activAdressName].NumberFormat = "### ### ##0.00";
                    sumC.Append(String.Format("{0}{1}", activAdressName, isLastSum ? ")" : "+"));
                    //D
                    activColumName = "D";
                    activAdressName = String.Format("{0}{1}", activColumName, sendActivRow);
                    cells[activAdressName].Value = isNotGlobalSum ? String.Format("=SUM({0}{1}:{0}{2})", activColumName, sendStartActivSumRowIndex, sendEndActivSumRowIndex) : sumD.ToString();
                    cells[activAdressName].NumberFormat = "### ### ##0.00";
                    sumD.Append(String.Format("{0}{1}", activAdressName, isLastSum ? ")" : "+"));
                    //E
                    activColumName = "E";
                    activAdressName = String.Format("{0}{1}", activColumName, sendActivRow);
                    cells[activAdressName].Value = isNotGlobalSum ? String.Format("=SUM({0}{1}:{0}{2})", activColumName, sendStartActivSumRowIndex, sendEndActivSumRowIndex) : sumE.ToString();
                    cells[activAdressName].NumberFormat = "### ### ##0.00";
                    sumE.Append(String.Format("{0}{1}", activAdressName, isLastSum ? ")" : "+"));
                    //F
                    activColumName = "F";
                    activAdressName = String.Format("{0}{1}", activColumName, sendActivRow);
                    cells[activAdressName].Value = isNotGlobalSum ? String.Format("=SUM({0}{1}:{0}{2})", activColumName, sendStartActivSumRowIndex, sendEndActivSumRowIndex) : sumF.ToString();
                    cells[activAdressName].NumberFormat = "### ### ##0.00";
                    sumF.Append(String.Format("{0}{1}", activAdressName, isLastSum ? ")" : "+"));
                };

                var activAdressRange = String.Format("A" + captionPosition + ":{0}" + captionPosition, sumColIndex);
                cells[activAdressRange].Merge();
                cells["A" + captionPosition].Value = "Відомість основних засобів по групам (скорочено)";
                cells[activAdressRange].Font.Bold = true;
                cells[activAdressRange].HorizontalAlignment = SpreadsheetGear.HAlign.Center;
                cells[activAdressRange].VerticalAlignment = SpreadsheetGear.VAlign.Center;

                activAdressRange = String.Format("A" + (captionPosition + 1) + ":{0}" + (captionPosition + 1), sumColIndex);
                cells[activAdressRange].Merge();
                cells["A" + (captionPosition + 1)].Value = "за період з " + startDate.ToShortDateString() + " по " + endDate.ToShortDateString();
                cells[activAdressRange].Font.Bold = true;
                cells[activAdressRange].HorizontalAlignment = SpreadsheetGear.HAlign.Center;
                cells[activAdressRange].VerticalAlignment = SpreadsheetGear.VAlign.Center;

                int startActivSumRowIndex = -1;
                int endActivSumRowIndex = -1;

                activRow--;
                foreach (var item in model)
                {
                    if (String.Compare(activGroupName, item.Name.ToString(), true) != 0)
                    {
                        if (String.Compare(activGroupName, String.Empty, true) != 0)
                        {
                            WriteSum(activRow);
                        }

                        if (String.Compare(activBalance_Account_Num, item.Num.ToString(), true) != 0)
                        {
                            if (String.Compare(activBalance_Account_Num, String.Empty, true) != 0)
                            {
                                activRow++;
                                endActivSumRowIndex = activRow - 1;
                                WriteColumSum(activRow, startActivSumRowIndex, endActivSumRowIndex, activBalance_Account_Num, false, true);
                            }
                            startActivSumRowIndex = activRow + 1;
                            activGroupName = String.Empty;
                            activBalance_Account_Num = item.Num.ToString();

                        }
                        activRow++;
                        activGroupName = item.Name.ToString();
                    }


                    cells[String.Format("{0}{1}", "A", activRow)].Value = item.Name;

                    var columName = mumIndex.FirstOrDefault(X => String.Compare(X.Value, item.Fixed_Account_Id.ToString(), true) == 0).Key;
                    if (columName != null)
                    {
                        cells[String.Format("{0}{1}", columName, activRow)].Value = item.PeriodAmortization;
                        cells[String.Format("{0}{1}", columName, activRow)].NumberFormat = "### ### ##0.00";
                    }
                }
                WriteSum(activRow);
                endActivSumRowIndex = activRow;
                activRow++;
                WriteColumSum(activRow, startActivSumRowIndex, endActivSumRowIndex, activBalance_Account_Num, true, true);
                //Global ROW sum
                activRow++;
                WriteColumSum(activRow, startActivSumRowIndex, endActivSumRowIndex, activBalance_Account_Num, true, false);

                PrintSignatures(cells, activRow + 3);

                cells["A" + startRow + ":" + sumColIndex + activRow].Borders.LineStyle = LineStyle.Continous;
                cells["A" + startRow + ":" + sumColIndex + activRow].Font.Size = 12;

                try
                {
                    Workbook.SaveAs(GeneratedReportsDir + "Відомість ОЗ по групам (скорочено)" + "за період з " + startDate.ToShortDateString() + " по " + endDate.ToShortDateString() + ".xls", FileFormat.Excel8);

                    Process process = new Process();
                    process.StartInfo.Arguments = "\"" + GeneratedReportsDir + "Відомість ОЗ по групам (скорочено)" + "за період з " + startDate.ToShortDateString() + " по " + endDate.ToShortDateString() + ".xls";// +"\"";
                    process.StartInfo.FileName = "Excel.exe";
                    process.Start();
                }
                catch (System.IO.IOException) { MessageBox.Show("Документ вже відкритий!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не знайдена програма Microsoft Excel!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }

            }
        }
    
        public void InputFixedAssetsForGroup(List<InputFixedAssetsForGroupDTO> model, DateTime startDate, DateTime endDate)
        {
            if (model.Count == 0)
            {
                MessageBox.Show("За вибраний період немає даних!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string templateName = " ";
            templateName = @"\Templates\InputFixedAssetsForGroup.xls";
            var Workbook = Factory.GetWorkbook(GeneratedReportsDir + templateName);
           
            var Worksheet = Workbook.Worksheets[0];
            var cells = Worksheet.Cells;

            int captionPosition = 7;
            int startGroupPosition = -100;
            int startPosition = captionPosition + 2;
            int currentPosition = startPosition;
            int n = 1;
            string activGroup = "";
            DateTime beginHeaderDate = Convert.ToDateTime(startDate).AddDays(-1);
            DateTime endHeaderDate = Convert.ToDateTime(endDate);

            StringBuilder SumG = new StringBuilder("=");
            StringBuilder SumH = new StringBuilder("=");
            StringBuilder SumI = new StringBuilder("=");


            cells["D" + captionPosition + ":" + "H" + captionPosition].Merge();
            cells["D" + captionPosition].HorizontalAlignment = HAlign.Center;
            cells["D" + captionPosition].VerticalAlignment = VAlign.Center;
            cells["D" + captionPosition].Font.Bold = true;
            cells["D" + captionPosition].Value = "Введення основних засобів з " + beginHeaderDate.ToShortDateString() + " по " + endHeaderDate.ToShortDateString() + "(по групам)";
            cells["G" + (captionPosition + 1)].Value = cells["G" + (captionPosition + 1)].Value + beginHeaderDate.ToShortDateString();
            cells["H" + (captionPosition + 1)].Value = cells["H" + (captionPosition + 1)].Value + endHeaderDate.ToShortDateString();

            Action<int, int> EndSumWrite = (currentPos, startGroupPos) =>
            {
                cells["A" + currentPos + ":I" + currentPos].Borders.LineStyle = LineStyle.Continous;
                cells["A" + currentPos + ":I" + currentPos].Interior.Color = Color.LightGreen;
                cells["A" + currentPos + ":I" + currentPos].Font.Bold = true;

                cells["A" + currentPos].Value = "Всього:";
                cells["G" + currentPos].Value = "=SUM(G" + startGroupPos.ToString() + ":G" + (currentPos - 1).ToString() + ")";
                cells["H" + currentPos].Value = "=SUM(H" + startGroupPos.ToString() + ":H" + (currentPos - 1).ToString() + ")";
                cells["I" + currentPos].Value = "=SUM(I" + startGroupPos.ToString() + ":I" + (currentPos - 1).ToString() + ")";


                cells["G" + currentPosition + ":" + "I" + currentPosition].NumberFormat = "### ### ##0.00";

                //global sum add
                SumG.AppendFormat("+G{0}", currentPosition.ToString());
                SumH.AppendFormat("+H{0}", currentPosition.ToString());
                SumI.AppendFormat("+I{0}", currentPosition.ToString());
            };

            for (int i = 0; i < model.Count; i++)
            {
                if (activGroup != model[i].Group_Id.ToString())
                {
                    if (startGroupPosition != -100)
                    {
                        EndSumWrite(currentPosition, startGroupPosition);
                        currentPosition++;
                    }
                    activGroup = model[i].Group_Id.ToString();
                    //Group
                    cells["A" + currentPosition].Value = "Група: " + model[i].GroupName;
                    cells["A" + currentPosition + ":" + "I" + currentPosition].Merge();
                    cells["A" + currentPosition.ToString() + ":" + "I" + currentPosition.ToString()].Font.Bold = true;
                    currentPosition++;
                    startGroupPosition = currentPosition;
                }

                cells["A" + currentPosition].Value = n;
                cells["B" + currentPosition].Value = model[i].Num;
                cells["C" + currentPosition].Value = model[i].YearSet;
                cells["D" + currentPosition].Value = model[i].MonthSet;
                cells["E" + currentPosition].Value = model[i].InventoryNumber;
                cells["E" + currentPosition].HorizontalAlignment = HAlign.Right;
                cells["F" + currentPosition].Value = model[i].InventoryName;
                cells["G" + currentPosition].Value = model[i].FYearPrice;
                cells["H" + currentPosition].Value = model[i].LYearPrice;
                cells["I" + currentPosition].Value = model[i].Difference;
                cells["G" + currentPosition + ":" + "I" + currentPosition].NumberFormat = "### ### ##0.00";
                currentPosition++;
                n++;
            }
            EndSumWrite(currentPosition, startGroupPosition);
            currentPosition++;

            cells["A" + currentPosition].Value = "Сума:";
            cells["A" + currentPosition + ":" + "I" + currentPosition].Font.Bold = true;
            cells["G" + currentPosition].Value = SumG.ToString();
            cells["H" + currentPosition].Value = SumH.ToString();
            cells["I" + currentPosition].Value = SumI.ToString();
            cells["A" + currentPosition + ":" + "I" + currentPosition].NumberFormat = "### ### ##0.00";

            cells["A" + currentPosition + ":I" + currentPosition].Borders.LineStyle = LineStyle.Continous;
            cells["A" + currentPosition + ":I" + currentPosition].Interior.Color = Color.Silver;
            cells["A" + currentPosition + ":I" + currentPosition].Font.Bold = true;

            PrintSignatures(cells, currentPosition + 3);

            cells["A" + startPosition + ":I" + currentPosition].Borders.LineStyle = LineStyle.Continous;

            try
            {
                string documentAddresName = GeneratedReportsDir + "Введення основних засобів з " + beginHeaderDate.ToShortDateString() + " по " + endHeaderDate.ToShortDateString() + "(по групам).xls";
                Workbook.SaveAs(documentAddresName, FileFormat.Excel8);

                Process process = new Process();
                process.StartInfo.Arguments = "\"" + documentAddresName + "\"";
                process.StartInfo.FileName = "Excel.exe";
                process.Start();
            }
            catch (System.IO.IOException) { MessageBox.Show("Документ вже відкритий!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (System.ComponentModel.Win32Exception) { MessageBox.Show("Не знайдена програма Microsoft Excel!", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }


        #endregion

        #endregion

  
        public void Dispose()
        {
            Database.Dispose();
        }
    }
}

