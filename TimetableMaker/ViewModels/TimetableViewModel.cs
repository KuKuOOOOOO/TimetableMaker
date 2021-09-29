using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TimetableMaker.Models;
using System.Windows.Input;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.IO;

namespace TimetableMaker.ViewModels
{
    class TimetableViewModel : ViewModelBase
    {
        private TimetableModel Timetable;
        private List<string> ClassList = new List<string>();
        private List<DateTime> StartTimeList = new List<DateTime>();
        private List<DateTime> EndTimeList = new List<DateTime>();
        public TimetableViewModel() // Construct
        {
            Timetable = new TimetableModel();
        }

        public string ClassName    //Binding class name
        {
            get { return Timetable._ClassName; }
            set
            {
                Timetable._ClassName = value;
                OnPropertyChanged();
            }
        }
        public string TeacherName  //Binding teacher name
        {
            get { return Timetable._TeacherName; }
            set
            {
                Timetable._TeacherName = value;
                OnPropertyChanged();
            }
        }
        public DateTime StartTime    //Binding start time
        {
            get
            {
                if (Timetable._StartTime == DateTime.MinValue)
                    return DateTime.Now;
                return Timetable._StartTime;
            }
            set
            {
                if (Timetable._StartTime != value)
                {
                    Timetable._StartTime = value;
                    OnPropertyChanged();
                }
            }
        }
        public DateTime EndTime
        {
            get
            {
                DateTime today = DateTime.Now;
                DateTime tomorrow = today.AddDays(1);
                if (Timetable._EndTime == DateTime.MinValue)
                    return tomorrow;
                return Timetable._EndTime;
            }
            set
            {
                if (Timetable._EndTime != value)
                {
                    Timetable._EndTime = value;
                    OnPropertyChanged();
                }
            }
        }
        // Add class in the table
        public ICommand AddClassCommand
        {
            get { return new RelayCommand(AddClass, CanExecute); }
        }
        public void AddClass()
        {
            TimeSpan SubDay = EndTime - StartTime;
            int SubHour = EndTime.Hour - StartTime.Hour;
            List<int> HourScope = new List<int>();
            for (int i = StartTime.Hour; i < EndTime.Hour; i++)
                HourScope.Add(i);
            if (TeacherName == null)
            {
                System.Windows.MessageBox.Show("教師名稱無法為空白", "Alert", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            else if (SubDay.Days <= 0 || SubHour <= 0)
            {
                System.Windows.MessageBox.Show("開始時間必須小於結束時間\n(課程結束日期>課程開始日期)&(課程結束時間>課程開始時間)", "Alert", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            else if (ClassName == null)
            {
                System.Windows.MessageBox.Show("課程名稱不能為空白", "Alert", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            else if (HourScope.Contains(12)) 
            {
                System.Windows.MessageBox.Show("課程時間無法跨越中午午餐時間(12:00)", "Alert", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            else
            {
                ClassList.Add(ClassName);
                StartTimeList.Add(StartTime);
                EndTimeList.Add(EndTime);
                System.Windows.MessageBox.Show(ClassName + "\n" + StartTime.ToString() + " ~ " + EndTime.ToString() + "\n新增成功", "TimetableMaker", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Asterisk);
                ClassName = "";
                StartTime = DateTime.MinValue;
                EndTime = DateTime.MinValue;
            }
        }
        // Loading classtable in Excel
        public ICommand LoadingCommand
        {
            get { return new RelayCommand(LoadingExcel, CanExecute); }
        }
        public void LoadingExcel()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "選擇要開啟的Excel";
            openFileDialog.Filter = "Excel活頁簿 (.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 1;
            openFileDialog.DefaultExt = ".xlsx";
            openFileDialog.Multiselect = true;
            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true) 
            {
                string path = openFileDialog.FileName;
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                app.DisplayAlerts = false;
                Workbook wb = app.Workbooks.Open(path, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet ws = wb.Worksheets[1];
                var cellValue = (string)(ws.Range["G5"]).Value;

                wb.Close();
                app.Quit();
            }

        }
        // Preview the class
        public ICommand PreviewCommand
        {
            get { return new RelayCommand(Preview, CanExecute); }
        }
        public void Preview()
        {
            if (ClassList.Count == 0)
            {
                System.Windows.MessageBox.Show("尚未輸入課表", "Information", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return;
            }
            else
            {
                string Message = TeacherName + " 老師的課程:\n";
                for (int i = 0; i < ClassList.Count; i++)
                    Message += ClassList[i] + " : " + StartTimeList[i].ToString("yyyy/MM/dd-HH:mm") + " ~ " + EndTimeList[i].ToString("yyyy/MM/dd-HH:mm") + "\n";
                System.Windows.MessageBox.Show(Message, "Information", System.Windows.MessageBoxButton.OK);
            }
        }
        // Export to the Excel
        public ICommand ExportExcelCommand
        {
            get { return new RelayCommand(ExportExcel, CanExecute); }
        }
        public void ExportExcel()
        {
            string XlsxPath = ExcelCalssDemo();
            if (XlsxPath == "Error")
            {
                System.Windows.MessageBox.Show("請將考核表放置與程式同個路徑並將檔名更改為\"Assessment.xlsx\"", "Alert", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.DisplayAlerts = false;
            Workbook wb = app.Workbooks.Open(@XlsxPath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet ws = wb.Worksheets[1];
            #region Print Class
            // Print Teacher name
            if (TeacherName != null)
                ws.Range["A3"].Value = TeacherName + " 老師";
            else
            {
                System.Windows.MessageBox.Show("教師名稱無法為空白", "Alert", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
            // Print Class name
            int InititalCell = 8;   // Time initial's Cell is A8
            for (int i = 0; i < 12; i++)
            {
                var StringCellTime = (string)(ws.Range["A" + (InititalCell + (i * 3)).ToString()] as Microsoft.Office.Interop.Excel.Range).Value;
                string[] CellTimeArray = StringCellTime.Split(new char[2] { ' ', '~' }, StringSplitOptions.RemoveEmptyEntries);
                DateTime CellStartTime = DateTime.ParseExact(CellTimeArray[0], "HH:mm", null);
                DateTime CellEndTime = DateTime.ParseExact(CellTimeArray[1], "HH:mm", null);
                //StringCellTime = StringCellTime.Split('~');
                DateTime CellTime = DateTime.Now; //DateTime.FromOADate(DoubleCellTime);
                for (int j = 0; j < EndTimeList.Count; j++)
                {
                    if (CellEndTime.Hour == EndTimeList[j].Hour)
                    {
                        int Span = EndTimeList[j].Hour - StartTimeList[j].Hour;
                        int DecreaseTime = i;
                        for (int k = Span; k > 0; k--)
                        {
                            string StartTimeWeek = System.Globalization.DateTimeFormatInfo.GetInstance(new System.Globalization.CultureInfo("zh-TW")).DayNames[(byte)StartTimeList[j].DayOfWeek];
                            switch (StartTimeWeek)
                            {
                                case "星期一":
                                    ws.Range["C" + (InititalCell + (DecreaseTime * 3)).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                    break;
                                case "星期二":
                                    ws.Range["E" + (InititalCell + (DecreaseTime * 3)).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                    break;
                                case "星期三":
                                    ws.Range["G" + (InititalCell + (DecreaseTime * 3)).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                    break;
                                case "星期四":
                                    ws.Range["I" + (InititalCell + (DecreaseTime * 3)).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                    break;
                                case "星期五":
                                    ws.Range["K" + (InititalCell + (DecreaseTime * 3)).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                    break;
                                case "星期六":
                                    ws.Range["M" + (InititalCell + (DecreaseTime * 3)).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                    break;
                                case "星期日":
                                    ws.Range["O" + (InititalCell + (DecreaseTime * 3)).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                    break;
                            }
                            if (DecreaseTime > 0)
                                DecreaseTime--;
                            else
                                DecreaseTime = 0;
                        }
                    }
                }
            }
            ws.Name = "教師課表";
            #endregion
            // Save file dialog
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.InitialDirectory = @"C:\";
            saveDlg.FileName = TeacherName + "課表";
            saveDlg.DefaultExt = ".xlsx";
            saveDlg.Filter = "Excel活頁簿 (.xlsx)|*.xlsx";
            Nullable<bool> result = saveDlg.ShowDialog();
            if (result == true)
            {
                try
                {
                    string path = saveDlg.FileName;
                    wb.SaveAs(path);
                    wb.Close();
                    app.Quit();
                    //System.IO.FileInfo fi = new System.IO.FileInfo(@"TimeTable.xlsx");
                    //fi.Delete();
                    System.Windows.MessageBox.Show("課表已輸出至" + path, "TimetableMaker", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Asterisk);
                    TeacherName = "";
                    ClassList.Clear();
                    ClassName = "";
                    StartTimeList.Clear();
                    EndTimeList.Clear();
                    StartTime = DateTime.MinValue;
                    EndTime = DateTime.MinValue;
                    if (File.Exists(XlsxPath))
                        File.Delete(XlsxPath);
                }
                catch (IOException IOex)
                {
                    wb.Close();
                    app.Quit();
                    if (File.Exists(XlsxPath))
                        File.Delete(XlsxPath);
                    System.Windows.MessageBox.Show(IOex.ToString(), "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    wb.Close();
                    app.Quit();
                    if (File.Exists(XlsxPath))
                        File.Delete(XlsxPath);
                    System.Windows.MessageBox.Show(ex.ToString(), "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
            }
        }
        public string ExcelCalssDemo()
        {
            string EXEpath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string WorkDir = Path.GetDirectoryName(EXEpath);
            string XlsxPath = WorkDir + @"\TimeTable.xlsx";
            string AssessmentPath = WorkDir + @"\Assessment.xlsx";
            if (File.Exists(XlsxPath))
                File.Delete(XlsxPath);
            // Check Assessment.xlsx is exists
            if (!File.Exists(AssessmentPath))
                return "Error";
            string currentSheet = "儲備講師";
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.DisplayAlerts = false;
            Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = wb.Worksheets[1];
            #region Include Assessment.xlsx
            // Open Assessment.xlsx
            Workbook Assessment_wb = app.Workbooks.Open(AssessmentPath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet Assessment_ws = Assessment_wb.Worksheets[currentSheet];
            // Check Assessment.xlsx's sheet have currentSheet
            foreach (Worksheet sheet in Assessment_wb.Worksheets)
            {
                if (sheet.Name == currentSheet)
                    Assessment_ws.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
            }
            Assessment_wb.Close();  // Assessment.xlsx close
            #endregion
            try
            {
                //Excel merge cell
                ws.Range["A1:P2"].Merge();
                ws.Range["A3:P4"].Merge();
                for (int i = 1; i <= 39; i += 3)
                {
                    ws.Range["A" + (i + 4).ToString() + ":B" + (i + 6).ToString()].Merge();
                    ws.Range["C" + (i + 4).ToString() + ":D" + (i + 6).ToString()].Merge();
                    ws.Range["E" + (i + 4).ToString() + ":F" + (i + 6).ToString()].Merge();
                    ws.Range["G" + (i + 4).ToString() + ":H" + (i + 6).ToString()].Merge();
                    ws.Range["I" + (i + 4).ToString() + ":J" + (i + 6).ToString()].Merge();
                    ws.Range["K" + (i + 4).ToString() + ":L" + (i + 6).ToString()].Merge();
                    ws.Range["M" + (i + 4).ToString() + ":N" + (i + 6).ToString()].Merge();
                    ws.Range["O" + (i + 4).ToString() + ":P" + (i + 6).ToString()].Merge();
                }
                //Excel border's line style
                ws.Range["A1:P43"].Borders.LineStyle = XlLineStyle.xlContinuous;

                ws.Range["A1:P2"].Cells.Interior.Color = XlRgbColor.rgbMediumAquamarine;
                ws.Range["A3:P4"].Cells.Interior.Color = XlRgbColor.rgbMistyRose;
                ws.Range["C5:P7"].Cells.Interior.Color = XlRgbColor.rgbLightGoldenrodYellow;
                ws.Range["A8:B43"].Cells.Interior.Color = XlRgbColor.rgbHoneydew;
                #region Excel cell text
                ws.Range["A1"].Value = "課表產生器";
                //Week
                ws.Range["C5"].Value = "星期一";
                ws.Range["E5"].Value = "星期二";
                ws.Range["G5"].Value = "星期三";
                ws.Range["I5"].Value = "星期四";
                ws.Range["K5"].Value = "星期五";
                ws.Range["M5"].Value = "星期六";
                ws.Range["O5"].Value = "星期日";
                //Section
                ws.Range["A8"].Value = "08:00 ~ 09:00";
                ws.Range["A11"].Value = "09:00 ~ 10:00";
                ws.Range["A14"].Value = "10:00 ~ 11:00";
                ws.Range["A17"].Value = "11:00 ~ 12:00";
                ws.Range["A20"].Value = "13:00 ~ 14:00";
                ws.Range["A23"].Value = "14:00 ~ 15:00";
                ws.Range["A26"].Value = "15:00 ~ 16:00";
                ws.Range["A29"].Value = "16:00 ~ 17:00";
                ws.Range["A32"].Value = "17:00 ~ 18:00";
                ws.Range["A35"].Value = "18:00 ~ 19:00";
                ws.Range["A38"].Value = "19:00 ~ 20:00";
                ws.Range["A41"].Value = "20:00 ~ 21:00";
                #endregion
                ws.Range["A1:P43"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                wb.SaveAs(XlsxPath);
                wb.Close();
                app.Quit();
                File.SetAttributes(XlsxPath, File.GetAttributes(XlsxPath) | FileAttributes.Hidden);
            }
            catch (IOException IOex)
            {
                wb.Close();
                app.Quit();
                if (File.Exists(XlsxPath))
                    File.Delete(XlsxPath);
                System.Windows.MessageBox.Show(IOex.ToString(), "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                wb.Close();
                app.Quit();
                if (File.Exists(XlsxPath))
                    File.Delete(XlsxPath);
                System.Windows.MessageBox.Show(ex.ToString(), "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            return XlsxPath;
        }
        public bool CanExecute()
        {
            return true;
        }
    }
}
