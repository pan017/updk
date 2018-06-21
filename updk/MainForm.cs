using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace updk
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        string connectionString;
        DateTime beginDate;
        DateTime endDate;
        List<TestResultValues> testResultValues;
        public MainForm(string connectionString)
        {
            InitializeComponent();
            this.connectionString = connectionString;
        }

        public MainForm(string connectionString, DateTime beginDate, DateTime endDate)
        {
            InitializeComponent();
            this.connectionString = connectionString;
            this.beginDate = beginDate;
            this.endDate = endDate;
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            mainDatagridView.AutoGenerateColumns = false;
            testResultValues = new List<TestResultValues>(); 
            string sqlExpression = @"SELECT  Patients.LastName + ' ' + Patients.FirstName + ' ' +
                Patients.Patronymic  AS FIO, Gender.Name AS Gender, Groups.Name AS [Group], TestResults.Name, 
                CASE  WHEN  
                    CONVERT(VARCHAR(MAX),[TestResultValues].Value) IS NULL THEN ' '
                    ELSE CONVERT(VARCHAR(MAX), [TestResultValues].Value) 
               END AS Result, Tests.Name, Patients.StageDate
               FROM [UPDK5].[dbo].[TestResultValues]  
                    INNER JOIN dbo.TestResults ON dbo.TestResults.ID = dbo.[TestResultValues].TestResultID 
                    INNER JOIN FinishedTests ON TestResultValues.FinishedTestID = FinishedTests.ID  
                    INNER JOIN TestPacks ON TestPacks.ID = FinishedTests.TestPackID  
                    INNER JOIN Patients ON PatientID = Patients.ID 
                    INNER JOIN Gender ON Gender.ID = Patients.GenderID  
                    INNER JOIN Groups ON GroupID = Groups.ID 
                    INNER JOIN TestPackTypes ON TestPacks.TestPackTypeID = TestPackTypes.ID  
                    INNER JOIN Tests ON Tests.ID = TestResults.TestID" +
                String.Format(" WHERE  TestPacks.Date > '{0}' ",  beginDate.ToShortDateString() == endDate.ToShortDateString() ? new DateTime(1900, 1, 1).ToShortDateString() : beginDate.ToShortDateString())
            +
                $"AND  TestPacks.Date < '{endDate.ToShortDateString()}'";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(sqlExpression, connection);
                SqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows) // если есть данные
                {

                    while (reader.Read()) // построчно считываем данные
                    {
                        testResultValues.Add(new TestResultValues
                        { 
                            FIO = reader.GetString(0),
                            Gender = reader.GetString(1),
                            Group = reader.GetString(2),
                            TestResultName = reader.GetString(3),
                            TestResultValue = reader.GetString(4),
                            TestName = reader.GetString(5),
                            StageDate = reader.GetDateTime(6)
                        });
                    }
                }

                reader.Close();
            }
            MessageBox.Show("Подключение установленно!");
            foreach (var item in testResultValues)
            {
                mainDatagridView.Rows.Add(item.FIO, item.Gender, item.Group, item.TestName, item.TestResultName, item.TestResultValue);
            }
            exportExcel();
        }
        bool IsTheSameCellValue(int column, int row)
        {
            DataGridViewCell cell1 = mainDatagridView[column, row];
            DataGridViewCell cell2 = mainDatagridView[column, row - 1];
            if (cell1.Value == null || cell2.Value == null)
            {
                return false;
            }
            return cell1.Value.ToString() == cell2.Value.ToString();
        }

        private void mainDatagridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex < 1 || e.ColumnIndex < 0)
                return;
            if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
            {
                e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
            }
            else
            {
                e.AdvancedBorderStyle.Top = mainDatagridView.AdvancedCellBorderStyle.Top;
            }  
        }

        private void mainDatagridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex == 0)
                return;
            if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
            {
                e.Value = "";
                e.FormattingApplied = true;
            }
        }

        void exportExcel()
        {
            List<string> tempNames = new List<string>();
            tempNames = testResultValues.Select(x => x.FIO).ToList();
            List<string> names = tempNames.Distinct().ToList();
            Excel.Application excelApp = new Excel.Application();
            // Создаём экземпляр рабочий книги Excel
            Excel.Workbook workBook;
            // Создаём экземпляр листа Excel
            Excel.Worksheet workSheet;
            int rowNumber = 4;
            workBook = excelApp.Workbooks.Add(@"D:\УПДК\1.xlsx");
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
            foreach (var item in names)
            {

                workSheet.Cells[rowNumber, 1] = item;
                DateTime date1 = testResultValues.FirstOrDefault(x => x.FIO == item).StageDate;
                System.TimeSpan diff1 = DateTime.Now - date1;
                
                workSheet.Cells[rowNumber, 3] = diff1.Days/365;
                rowNumber++;

            }
            rowNumber = 4;
            for (int i = 0; i < names.Count; i++)
            {
                for (int j = 0; j < testResultValues.Count; j++)
                {
                    if (testResultValues[j].TestName == "Уровень восприятия скорости и расстояния" && testResultValues[j].FIO == names[i])
                    {
                        workSheet.Cells[rowNumber, 9] = testResultValues[j].TestResultValue;
                        int result;
                        if (int.TryParse(testResultValues[j].TestResultValue, out result))
                        {
                            if (result >= 10)
                            {
                                workSheet.Cells[rowNumber, 4] = "норма";
                            }
                            if (result >= 7 && result < 10)
                            {
                                workSheet.Cells[rowNumber, 4] = "Удовл";
                            }
                            if (result < 7)
                            {
                                workSheet.Cells[rowNumber, 4] = "Неуд";
                              
                            }
                        }
                    }
                    if (testResultValues[j].TestName == "Оценка склонности к риску" && 
                        testResultValues[j].FIO == names[i] &&  testResultValues[j].TestResultName == "Количество баллов")
                    {
                        workSheet.Cells[rowNumber, 10] = testResultValues[j].TestResultValue;
                        int result;
                        if (int.TryParse(testResultValues[j].TestResultValue, out result))
                        {
                            if (result >= 0 && result <=3)
                            {
                                workSheet.Cells[rowNumber, 5] = "Хорошо";
                            }
                            if (result >=4 && result <= 6)
                            {
                                workSheet.Cells[rowNumber, 5] = "Удовл";
                            }
                            if (result >= 7 && result < 10)
                            {
                                workSheet.Cells[rowNumber, 5] = "Неуд";

                            }
                        }
                    }
                        
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Среднее время реагирования в задании №1")
                        workSheet.Cells[rowNumber, 11] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество правильных ответов на зрительные стимулы в задании №1")
                        workSheet.Cells[rowNumber, 12] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Разница средних времен реагирования между заданием №2 и заданием №1")
                        workSheet.Cells[rowNumber, 13] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество правильных ответов на слуховые стимулы в задании №2")
                        workSheet.Cells[rowNumber, 14] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество ошибочных ответов на зрительные стимулы")
                        workSheet.Cells[rowNumber, 15] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Среднее время реагирования в задании №2 на зрительные стимулы")
                        workSheet.Cells[rowNumber, 16] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество правильных реагирований на зрительные стимулы в задании № 2")
                        workSheet.Cells[rowNumber, 17] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество реагирований при отсутствии сигнала в задании № 1 на зрительные стимулы")
                        workSheet.Cells[rowNumber, 18] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество реагирований при отсутствии сигнала в задании № 2 на зрительные стимулы")
                        workSheet.Cells[rowNumber, 19] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Распределение внимания" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Разница количества правильных ответов на зрительные стимулы (№1 - №2)")
                        workSheet.Cells[rowNumber, 20] = testResultValues[j].TestResultValue;

                    if (testResultValues[j].TestName == "Эмоциональная устойчивость" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество ошибок без помехи (N1)")
                        workSheet.Cells[rowNumber, 21] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Эмоциональная устойчивость" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество пропусков без помехи")
                        workSheet.Cells[rowNumber, 22] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Эмоциональная устойчивость" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Среднеарифметическое время реагирования без помехи (ВР1)")
                        workSheet.Cells[rowNumber, 23] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Эмоциональная устойчивость" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество ошибок с помехой (N2)")
                        workSheet.Cells[rowNumber, 24] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Эмоциональная устойчивость" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество пропусков с помехой")
                        workSheet.Cells[rowNumber, 25] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Эмоциональная устойчивость" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Среднеарифметическое время реагирования с помехой (ВР2)")
                        workSheet.Cells[rowNumber, 26] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Эмоциональная устойчивость" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Разница среднеарифметических времен реагирования (ВР2 - ВР1)")
                        workSheet.Cells[rowNumber, 27] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Эмоциональная устойчивость" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Разница количества ошибок с помехой и без помехи (N2 - N1)")
                        workSheet.Cells[rowNumber, 28] = testResultValues[j].TestResultValue;

                    if (testResultValues[j].TestName == "Сложная двигательная реакция - М" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Среднее время реагирования в задании №1")
                        workSheet.Cells[rowNumber, 29] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Сложная двигательная реакция - М" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество ошибок в задании №2")
                        workSheet.Cells[rowNumber, 30] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Сложная двигательная реакция - М" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Время выбора")
                        workSheet.Cells[rowNumber, 31] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Сложная двигательная реакция - М" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Количество нажатий на кнопку при отсутствии сигнала")
                        workSheet.Cells[rowNumber, 32] = testResultValues[j].TestResultValue;
                    if (testResultValues[j].TestName == "Сложная двигательная реакция - М" && testResultValues[j].FIO == names[i] && testResultValues[j].TestResultName == "Среднее время реагирования в задании №2")
                        workSheet.Cells[rowNumber, 33] = testResultValues[j].TestResultValue;

                }
                {
                    var res1 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Эмоциональная устойчивость" && x.TestResultName == "Количество ошибок без помехи (N1)");
                    var res2 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Эмоциональная устойчивость" && x.TestResultName == "Количество пропусков без помехи");
                    var res3 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Эмоциональная устойчивость" && x.TestResultName == "Среднеарифметическое время реагирования без помехи (ВР1)");
                    var res4 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Эмоциональная устойчивость" && x.TestResultName == "Количество ошибок с помехой (N2)");
                    var res5 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Эмоциональная устойчивость" && x.TestResultName == "Количество пропусков с помехой");
                    var res6 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Эмоциональная устойчивость" && x.TestResultName == "Среднеарифметическое время реагирования с помехой (ВР2)");
                    var res7 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Эмоциональная устойчивость" && x.TestResultName == "Разница среднеарифметических времен реагирования (ВР2 - ВР1)");
                    var res8 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Эмоциональная устойчивость" && x.TestResultName == "Разница количества ошибок с помехой и без помехи (N2 - N1)");
                    if (res1 != null && res2 != null && res3 != null && res4 != null && res5 != null && res6 != null && res7 != null && res8 != null)
                    {
                        if (double.Parse(res1.TestResultValue) <= 2
                            && int.Parse(res2.TestResultValue) <= 1
                            && double.Parse(res3.TestResultValue) * 1000 <= 900 
                            && int.Parse(res4.TestResultValue) <= 4
                            && int.Parse(res5.TestResultValue) <= 2
                            && double.Parse(res6.TestResultValue) * 1000 <= 1250
                            && double.Parse(res7.TestResultValue) * 1000 <= 350
                            && int.Parse(res8.TestResultValue) < 2)
                        {
                            workSheet.Cells[rowNumber, 7] = "Хорошо";
                        }
                        else if (double.Parse(res1.TestResultValue) <= 2
                            && int.Parse(res2.TestResultValue) <= 1
                            && double.Parse(res3.TestResultValue) * 1000 <= 900
                            && int.Parse(res4.TestResultValue) <= 5
                            && int.Parse(res5.TestResultValue) <= 3
                            && double.Parse(res6.TestResultValue) * 1000 <= 1350
                            && double.Parse(res7.TestResultValue) * 1000 <= 350
                            && int.Parse(res8.TestResultValue) < 3)
                        {
                            workSheet.Cells[rowNumber, 7] = "Удовл";
                        }
                        else
                        {
                            workSheet.Cells[rowNumber, 7] = "Неуд";
                        }
                    }
                }
                {
                    var res1 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Распределение внимания" && x.TestResultName == "Среднее время реагирования в задании №1");
                    var res2 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Распределение внимания" && x.TestResultName == "Количество правильных ответов на зрительные стимулы в задании №1");
                    var res3 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Распределение внимания" && x.TestResultName == "Разница средних времен реагирования между заданием №2 и заданием №1");
                    var res4 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Распределение внимания" && x.TestResultName == "Количество правильных ответов на слуховые стимулы в задании №2");
                    var res5 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Распределение внимания" && x.TestResultName == "Разница количества правильных ответов на зрительные стимулы (№1 - №2)");

                    if (res1 != null && res2 != null && res3 != null && res4 != null && res5 != null)
                    {
                        if (double.Parse(res1.TestResultValue) * 1000 <= 600 && int.Parse(res2.TestResultValue) >= 17 && double.Parse(res3.TestResultValue) * 1000 <= 300 && int.Parse(res4.TestResultValue) >= 9 && int.Parse(res5.TestResultValue) <= 3)
                        {
                            workSheet.Cells[rowNumber, 6] = "Хорошо";
                        }
                        else if (double.Parse(res1.TestResultValue) * 1000 <= 600 && int.Parse(res2.TestResultValue) >= 17 && double.Parse(res3.TestResultValue) * 1000 <= 350 && int.Parse(res4.TestResultValue) >= 8 && int.Parse(res5.TestResultValue) <= 4)
                        {
                            workSheet.Cells[rowNumber, 6] = "Удовл";
                        }
                        else
                        {
                            workSheet.Cells[rowNumber, 6] = "Неуд";
                        }
                    }

                }
                {
                    var res1 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Сложная двигательная реакция - М" && x.TestResultName == "Среднее время реагирования в задании №1");
                    var res2 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Сложная двигательная реакция - М" && x.TestResultName == "Количество ошибок в задании №2");
                    var res3 = testResultValues.FirstOrDefault(x => x.FIO == names[i] && x.TestName == "Сложная двигательная реакция - М" && x.TestResultName == "Время выбора");

                    if (res1 != null && res2 != null && res3 != null)
                    {
                        if (double.Parse(res1.TestResultValue) * 1000 <= 360
                            && int.Parse(res2.TestResultValue) <= 4
                            && double.Parse(res3.TestResultValue) * 1000 <= 300)
                        {
                            workSheet.Cells[rowNumber, 8] = "Хорошо";
                        }
                        else if (double.Parse(res1.TestResultValue) * 1000 <= 360
                                && int.Parse(res2.TestResultValue) <= 5
                                && double.Parse(res3.TestResultValue) * 1000 <= 350)
                        {
                            workSheet.Cells[rowNumber, 8] = "Удовл";
                        }
                        else
                        {
                            workSheet.Cells[rowNumber, 8] = "Неуд";
                        }
                    }
                }
                    rowNumber++;
            }
            // Открываем созданный excel-файл
            excelApp.Visible = true;
            excelApp.UserControl = true;

            //String targetfile = @"D:\УПДК\rep.xlsx";
            //if (File.Exists(targetfile))
            //    File.Delete(targetfile);
            //File.Copy(@"D:\УПДК\1.xlsx", targetfile);

            //var excelapp = new Excel.Application();
            //excelapp.Visible = false;
            ////Получаем набор ссылок на объекты Workbook
            ////var excelappworkbooks = excelapp.Workbooks;
            ////Открываем книгу и получаем на нее ссылку
            //var excelappworkbook = excelapp.Workbooks.Add(@"D:\УПДК\1.xlsx");
            ////Получаем массив ссылок на листы выбранной книги
            //var excelsheets = excelappworkbook.Worksheets;
            ////Получаем ссылку на лист 1
            //var excelworksheet = (Excel.Worksheet)excelsheets.Item[1];

            //Excel.Range excelcells;
            //int rowNumber = 2;

            //object _missingObj = System.Reflection.Missing.Value;
            //List<string> tempNames = new List<string>();
            //tempNames = testResultValues.Select(x => x.FIO).ToList();
            //List<string> names = tempNames.Distinct().ToList();
            //try
            //{
            //    foreach (var item in names)
            //    {

            //        excelcells = (Excel.Range)excelworksheet.Cells[rowNumber, 2];
            //        excelcells.Value2 = item;
               
            //        rowNumber++;
                   
            //    }

            //    excelappworkbook.SaveAs(targetfile, Excel.XlFileFormat.xlWorkbookNormal, _missingObj, _missingObj, _missingObj, _missingObj,
            //            Excel.XlSaveAsAccessMode.xlExclusive, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj);



            //}
            //catch (Exception e)
            //{

            //}
            //finally
            //{

            //    excelworksheet = null;
            //    excelsheets = null;
            //    try
            //    {// Тут уж простите за заплатку, выпадало исключение, если файл с таким именем существует и пользователь откажется перезаписывать его
            //     // Не знаю почему, времени нет разбираться.
            //     //excelappworkbook.Save();
            //    }
            //    catch { }

            //    excelappworkbook.Close(false, _missingObj, _missingObj);// Закрываем книгу
            //                                                            //excelAppWorkbooks = excelApp.Workbooks;     // Далее проверяем есть ли ещё другие открытые книги, ведь во время работы нашей программы пользователь мог открыть другую книгу
            //                                                            //if (excelAppWorkbooks.Count == 0)
            //    excelapp.Quit();            // Если нет то закрываем приложение
            //    excelappworkbook = null;        // Продолжаем обнулять ссылки
            //                                    //excelAppWorkbooks = null;
            //    excelapp = null;
            //    GC.Collect();

            //}


        }
    }
}
