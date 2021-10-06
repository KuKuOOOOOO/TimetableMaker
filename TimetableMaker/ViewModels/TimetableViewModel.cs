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
            if (ClassList.Count != 0)
            {
                System.Windows.MessageBox.Show("已有相關課程紀錄至程式內\n請先輸出成Excel課表後再試一次", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }
            else
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
                    try
                    {
                        int InitialCell = 8, LastCell = 79;
                        var TeacherValue = (string)(ws.Range["A3"]).Value;
                        if (TeacherValue == null || TeacherValue == "")
                        {
                            System.Windows.MessageBox.Show("無法找到教師名稱\n請確認課表是否正確", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                            return;
                        }
                        else
                        {
                            TeacherValue = TeacherValue.Substring(0, TeacherValue.IndexOf(' '));
                            TeacherName = TeacherValue;
                        }
                        #region Loading all class function
                        // Monday
                        for (int i = InitialCell; i < LastCell; i++)
                        {
                            var MondayValue = (string)(ws.Range["C" + i.ToString()]).Value;
                            if (MondayValue == null || MondayValue == "")
                                continue;
                            else
                            {
                                if (ClassList.Contains(MondayValue.Substring(0, MondayValue.IndexOf('\n'))))
                                    continue;
                                else
                                {
                                    ClassList.Add(MondayValue.Substring(0, MondayValue.IndexOf('\n')));
                                    string StringDate = MondayValue.Substring(MondayValue.IndexOf('\n') + 1);
                                    string StringStartDate = StringDate.Substring(0, StringDate.IndexOf('~'));
                                    string StringEndDate = StringDate.Substring(StringDate.IndexOf('~') + 1);
                                    string StringComment = (string)ws.Range["C" + i.ToString()].Comment.Text();
                                    string StringStartTime = StringComment.Substring(0, StringComment.IndexOf("~"));
                                    string StringEndTime = StringComment.Substring(StringComment.IndexOf('~') + 1);
                                    DateTime StartDateTime = DateTime.Parse(StringStartDate).Add(TimeSpan.Parse(StringStartTime));
                                    DateTime EndDateTime = DateTime.Parse(StringEndDate).Add(TimeSpan.Parse(StringEndTime));
                                    StartTimeList.Add(StartDateTime);
                                    EndTimeList.Add(EndDateTime);
                                }
                            }
                        }
                        // Tuesday
                        for (int i = InitialCell; i < LastCell; i++)
                        {
                            var TuesdayValue = (string)(ws.Range["E" + i.ToString()]).Value;
                            if (TuesdayValue == null || TuesdayValue == "")
                                continue;
                            else
                            {
                                if (ClassList.Contains(TuesdayValue.Substring(0, TuesdayValue.IndexOf('\n'))))
                                    continue;
                                else
                                {
                                    ClassList.Add(TuesdayValue.Substring(0, TuesdayValue.IndexOf('\n')));
                                    string StringDate = TuesdayValue.Substring(TuesdayValue.IndexOf('\n') + 1);
                                    string StringStartDate = StringDate.Substring(0, StringDate.IndexOf('~'));
                                    string StringEndDate = StringDate.Substring(StringDate.IndexOf('~') + 1);
                                    string StringComment = (string)ws.Range["E" + i.ToString()].Comment.Text();
                                    string StringStartTime = StringComment.Substring(0, StringComment.IndexOf("~"));
                                    string StringEndTime = StringComment.Substring(StringComment.IndexOf('~') + 1);
                                    DateTime StartDateTime = DateTime.Parse(StringStartDate).Add(TimeSpan.Parse(StringStartTime));
                                    DateTime EndDateTime = DateTime.Parse(StringEndDate).Add(TimeSpan.Parse(StringEndTime));
                                    StartTimeList.Add(StartDateTime);
                                    EndTimeList.Add(EndDateTime);
                                }
                            }
                        }
                        // Wednesday
                        for (int i = InitialCell; i < LastCell; i++)
                        {
                            var WednesdayValue = (string)(ws.Range["G" + i.ToString()]).Value;
                            if (WednesdayValue == null || WednesdayValue == "")
                                continue;
                            else
                            {
                                if (ClassList.Contains(WednesdayValue.Substring(0, WednesdayValue.IndexOf('\n'))))
                                    continue;
                                else
                                {
                                    ClassList.Add(WednesdayValue.Substring(0, WednesdayValue.IndexOf('\n')));
                                    string StringDate = WednesdayValue.Substring(WednesdayValue.IndexOf('\n') + 1);
                                    string StringStartDate = StringDate.Substring(0, StringDate.IndexOf('~'));
                                    string StringEndDate = StringDate.Substring(StringDate.IndexOf('~') + 1);
                                    string StringComment = (string)ws.Range["G" + i.ToString()].Comment.Text();
                                    string StringStartTime = StringComment.Substring(0, StringComment.IndexOf("~"));
                                    string StringEndTime = StringComment.Substring(StringComment.IndexOf('~') + 1);
                                    DateTime StartDateTime = DateTime.Parse(StringStartDate).Add(TimeSpan.Parse(StringStartTime));
                                    DateTime EndDateTime = DateTime.Parse(StringEndDate).Add(TimeSpan.Parse(StringEndTime));
                                    StartTimeList.Add(StartDateTime);
                                    EndTimeList.Add(EndDateTime);
                                }
                            }
                        }
                        // Thursday
                        for (int i = InitialCell; i < LastCell; i++)
                        {
                            var ThursdayValue = (string)(ws.Range["I" + i.ToString()]).Value;
                            if (ThursdayValue == null || ThursdayValue == "")
                                continue;
                            else
                            {
                                if (ClassList.Contains(ThursdayValue.Substring(0, ThursdayValue.IndexOf('\n'))))
                                    continue;
                                else
                                {
                                    ClassList.Add(ThursdayValue.Substring(0, ThursdayValue.IndexOf('\n')));
                                    string StringDate = ThursdayValue.Substring(ThursdayValue.IndexOf('\n') + 1);
                                    string StringStartDate = StringDate.Substring(0, StringDate.IndexOf('~'));
                                    string StringEndDate = StringDate.Substring(StringDate.IndexOf('~') + 1);
                                    string StringComment = (string)ws.Range["I" + i.ToString()].Comment.Text();
                                    string StringStartTime = StringComment.Substring(0, StringComment.IndexOf("~"));
                                    string StringEndTime = StringComment.Substring(StringComment.IndexOf('~') + 1);
                                    DateTime StartDateTime = DateTime.Parse(StringStartDate).Add(TimeSpan.Parse(StringStartTime));
                                    DateTime EndDateTime = DateTime.Parse(StringEndDate).Add(TimeSpan.Parse(StringEndTime));
                                    StartTimeList.Add(StartDateTime);
                                    EndTimeList.Add(EndDateTime);
                                }
                            }
                        }
                        // Friday
                        for (int i = InitialCell; i < LastCell; i++)
                        {
                            var FridayValue = (string)(ws.Range["K" + i.ToString()]).Value;
                            if (FridayValue == null || FridayValue == "")
                                continue;
                            else
                            {
                                if (ClassList.Contains(FridayValue.Substring(0, FridayValue.IndexOf('\n'))))
                                    continue;
                                else
                                {
                                    ClassList.Add(FridayValue.Substring(0, FridayValue.IndexOf('\n')));
                                    string StringDate = FridayValue.Substring(FridayValue.IndexOf('\n') + 1);
                                    string StringStartDate = StringDate.Substring(0, StringDate.IndexOf('~'));
                                    string StringEndDate = StringDate.Substring(StringDate.IndexOf('~') + 1);
                                    string StringComment = (string)ws.Range["K" + i.ToString()].Comment.Text();
                                    string StringStartTime = StringComment.Substring(0, StringComment.IndexOf("~"));
                                    string StringEndTime = StringComment.Substring(StringComment.IndexOf('~') + 1);
                                    DateTime StartDateTime = DateTime.Parse(StringStartDate).Add(TimeSpan.Parse(StringStartTime));
                                    DateTime EndDateTime = DateTime.Parse(StringEndDate).Add(TimeSpan.Parse(StringEndTime));
                                    StartTimeList.Add(StartDateTime);
                                    EndTimeList.Add(EndDateTime);
                                }
                            }
                        }
                        // Saturday
                        for (int i = InitialCell; i < LastCell; i++)
                        {
                            var SaturdayValue = (string)(ws.Range["M" + i.ToString()]).Value;
                            if (SaturdayValue == null || SaturdayValue == "")
                                continue;
                            else
                            {
                                if (ClassList.Contains(SaturdayValue.Substring(0, SaturdayValue.IndexOf('\n'))))
                                    continue;
                                else
                                {
                                    ClassList.Add(SaturdayValue.Substring(0, SaturdayValue.IndexOf('\n')));
                                    string StringDate = SaturdayValue.Substring(SaturdayValue.IndexOf('\n') + 1);
                                    string StringStartDate = StringDate.Substring(0, StringDate.IndexOf('~'));
                                    string StringEndDate = StringDate.Substring(StringDate.IndexOf('~') + 1);
                                    string StringComment = (string)ws.Range["M" + i.ToString()].Comment.Text();
                                    string StringStartTime = StringComment.Substring(0, StringComment.IndexOf("~"));
                                    string StringEndTime = StringComment.Substring(StringComment.IndexOf('~') + 1);
                                    DateTime StartDateTime = DateTime.Parse(StringStartDate).Add(TimeSpan.Parse(StringStartTime));
                                    DateTime EndDateTime = DateTime.Parse(StringEndDate).Add(TimeSpan.Parse(StringEndTime));
                                    StartTimeList.Add(StartDateTime);
                                    EndTimeList.Add(EndDateTime);
                                }
                            }
                        }
                        // Sunday
                        for (int i = InitialCell; i < LastCell; i++)
                        {
                            var SundayValue = (string)(ws.Range["O" + i.ToString()]).Value;
                            if (SundayValue == null || SundayValue == "")
                                continue;
                            else
                            {
                                if (ClassList.Contains(SundayValue.Substring(0, SundayValue.IndexOf('\n'))))
                                    continue;
                                else
                                {
                                    ClassList.Add(SundayValue.Substring(0, SundayValue.IndexOf('\n')));
                                    string StringDate = SundayValue.Substring(SundayValue.IndexOf('\n') + 1);
                                    string StringStartDate = StringDate.Substring(0, StringDate.IndexOf('~'));
                                    string StringEndDate = StringDate.Substring(StringDate.IndexOf('~') + 1);
                                    string StringComment = (string)ws.Range["O" + i.ToString()].Comment.Text();
                                    string StringStartTime = StringComment.Substring(0, StringComment.IndexOf("~"));
                                    string StringEndTime = StringComment.Substring(StringComment.IndexOf('~') + 1);
                                    DateTime StartDateTime = DateTime.Parse(StringStartDate).Add(TimeSpan.Parse(StringStartTime));
                                    DateTime EndDateTime = DateTime.Parse(StringEndDate).Add(TimeSpan.Parse(StringEndTime));
                                    StartTimeList.Add(StartDateTime);
                                    EndTimeList.Add(EndDateTime);
                                }
                            }
                        }
                        #endregion
                        System.Windows.MessageBox.Show("讀取完成\n點擊預覽課表可確認相關課程", "Information", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                        wb.Close();
                        app.Quit();
                    }
                    catch(Exception ex)
                    {
                        System.Windows.MessageBox.Show("請確認讀取的檔案是否正確無誤(***課表.xlsx)\n" + ex.ToString(), "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                        wb.Close();
                        app.Quit();
                        return;
                    }

                }
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
            for (int i = 0; i < 72; i++)
            {
                var StringCellTime = (string)(ws.Range["A" + (InititalCell + i).ToString()] as Microsoft.Office.Interop.Excel.Range).Value;
                string[] CellTimeArray = StringCellTime.Split(new char[2] { ' ', '~' }, StringSplitOptions.RemoveEmptyEntries);
                DateTime CellStartTime = DateTime.ParseExact(CellTimeArray[0], "HH:mm", null);
                DateTime CellEndTime = DateTime.ParseExact(CellTimeArray[1], "HH:mm", null);
                //StringCellTime = StringCellTime.Split('~');
                DateTime CellTime = DateTime.Now; //DateTime.FromOADate(DoubleCellTime);
                for (int j = 0; j < EndTimeList.Count; j++)
                {
                    if (CellEndTime.ToString("HH:mm") == EndTimeList[j].ToString("HH:mm"))
                    {
                        Console.WriteLine("YES");
                        double DoubleSpan = (EndTimeList[j].TimeOfDay - StartTimeList[j].TimeOfDay).TotalMinutes;
                        int Span = Convert.ToInt32(DoubleSpan / 10);
                        int DecreaseTime = i;
                        for (int k = Span; k > 0; k--)
                        {
                            string StartTimeWeek = System.Globalization.DateTimeFormatInfo.GetInstance(new System.Globalization.CultureInfo("zh-TW")).DayNames[(byte)StartTimeList[j].DayOfWeek];
                            switch(StartTimeWeek)
                            {
                                case "星期一":
                                    {
                                        ws.Range["C" + (InititalCell + DecreaseTime).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                        ws.Range["C" + (InititalCell + DecreaseTime).ToString()].AddComment(StartTimeList[j].ToString("HH:mm") + " ~ " + EndTimeList[j].ToString("HH:mm"));
                                        if (k == 1)
                                        {
                                            ws.Range["C" + (InititalCell + DecreaseTime).ToString() + ":D" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = false;
                                            ws.Range["C" + (InititalCell + DecreaseTime).ToString() + ":D" + (InititalCell + DecreaseTime + Span - 1).ToString()].Merge();
                                            ws.Range["C" + (InititalCell + DecreaseTime).ToString() + ":D" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = true;
                                        }
                                    }
                                    break;
                                case "星期二":
                                    {
                                        ws.Range["E" + (InititalCell + DecreaseTime).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                        ws.Range["E" + (InititalCell + DecreaseTime).ToString()].AddComment(StartTimeList[j].ToString("HH:mm") + " ~ " + EndTimeList[j].ToString("HH:mm"));
                                        if (k == 1)
                                        {
                                            ws.Range["E" + (InititalCell + DecreaseTime).ToString() + ":F" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = false;
                                            ws.Range["E" + (InititalCell + DecreaseTime).ToString() + ":F" + (InititalCell + DecreaseTime + Span - 1).ToString()].Merge();
                                            ws.Range["E" + (InititalCell + DecreaseTime).ToString() + ":F" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = true;
                                        }
                                    }
                                    break;
                                case "星期三":
                                    {
                                        ws.Range["G" + (InititalCell + DecreaseTime).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                        ws.Range["G" + (InititalCell + DecreaseTime).ToString()].AddComment(StartTimeList[j].ToString("HH:mm") + " ~ " + EndTimeList[j].ToString("HH:mm"));
                                        if (k == 1)
                                        {
                                            ws.Range["G" + (InititalCell + DecreaseTime).ToString() + ":H" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = false;
                                            ws.Range["G" + (InititalCell + DecreaseTime).ToString() + ":H" + (InititalCell + DecreaseTime + Span - 1).ToString()].Merge();
                                            ws.Range["G" + (InititalCell + DecreaseTime).ToString() + ":H" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = true;
                                        }
                                    }
                                    break;
                                case "星期四":
                                    {
                                        ws.Range["I" + (InititalCell + DecreaseTime).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                        ws.Range["I" + (InititalCell + DecreaseTime).ToString()].AddComment(StartTimeList[j].ToString("HH:mm") + " ~ " + EndTimeList[j].ToString("HH:mm"));
                                        if (k == 1)
                                        {
                                            ws.Range["I" + (InititalCell + DecreaseTime).ToString() + ":J" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = false;
                                            ws.Range["I" + (InititalCell + DecreaseTime).ToString() + ":J" + (InititalCell + DecreaseTime + Span - 1).ToString()].Merge();
                                            ws.Range["I" + (InititalCell + DecreaseTime).ToString() + ":J" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = true;
                                        }
                                    }
                                    break;
                                case "星期五":
                                    {
                                        ws.Range["K" + (InititalCell + DecreaseTime).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                        ws.Range["K" + (InititalCell + DecreaseTime).ToString()].AddComment(StartTimeList[j].ToString("HH:mm") + " ~ " + EndTimeList[j].ToString("HH:mm"));
                                        if (k == 1)
                                        {
                                            ws.Range["K" + (InititalCell + DecreaseTime).ToString() + ":L" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = false;
                                            ws.Range["K" + (InititalCell + DecreaseTime).ToString() + ":L" + (InititalCell + DecreaseTime + Span - 1).ToString()].Merge();
                                            ws.Range["K" + (InititalCell + DecreaseTime).ToString() + ":L" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = true;
                                        }
                                    }
                                    break;
                                case "星期六":
                                    {
                                        ws.Range["M" + (InititalCell + DecreaseTime).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                        ws.Range["M" + (InititalCell + DecreaseTime).ToString()].AddComment(StartTimeList[j].ToString("HH:mm") + " ~ " + EndTimeList[j].ToString("HH:mm"));
                                        if (k == 1)
                                        {
                                            ws.Range["M" + (InititalCell + DecreaseTime).ToString() + ":N" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = false;
                                            ws.Range["M" + (InititalCell + DecreaseTime).ToString() + ":N" + (InititalCell + DecreaseTime + Span - 1).ToString()].Merge();
                                            ws.Range["M" + (InititalCell + DecreaseTime).ToString() + ":N" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = true;
                                        }
                                    }
                                    break;
                                case "星期日":
                                    {
                                        ws.Range["O" + (InititalCell + DecreaseTime).ToString()].Value = ClassList[j].ToString() + "\n" + StartTimeList[j].ToString("MM/dd") + " ~ " + EndTimeList[j].ToString("MM/dd");
                                        ws.Range["O" + (InititalCell + DecreaseTime).ToString()].AddComment(StartTimeList[j].ToString("HH:mm") + " ~ " + EndTimeList[j].ToString("HH:mm"));
                                        if (k == 1)
                                        {
                                            ws.Range["O" + (InititalCell + DecreaseTime).ToString() + ":P" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = false;
                                            ws.Range["O" + (InititalCell + DecreaseTime).ToString() + ":P" + (InititalCell + DecreaseTime + Span - 1).ToString()].Merge();
                                            ws.Range["O" + (InititalCell + DecreaseTime).ToString() + ":P" + (InititalCell + DecreaseTime + Span - 1).ToString()].WrapText = true;
                                        }
                                    }
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
                // Excel merge cell
                ws.Range["A1:P2"].Merge();
                ws.Range["A3:P4"].Merge();
                // Week Merge
                ws.Range["A5:B7"].Merge();
                ws.Range["C5:D7"].Merge();
                ws.Range["E5:F7"].Merge();
                ws.Range["G5:H7"].Merge();
                ws.Range["I5:J7"].Merge();
                ws.Range["K5:L7"].Merge();
                ws.Range["M5:N7"].Merge();
                ws.Range["O5:P7"].Merge();
                for (int i = 1; i <= 72; i++)
                {
                    ws.Range["A" + (i + 7).ToString() + ":B" + (i + 7).ToString()].Merge();
                    ws.Range["C" + (i + 7).ToString() + ":D" + (i + 7).ToString()].Merge();
                    ws.Range["E" + (i + 7).ToString() + ":F" + (i + 7).ToString()].Merge();
                    ws.Range["G" + (i + 7).ToString() + ":H" + (i + 7).ToString()].Merge();
                    ws.Range["I" + (i + 7).ToString() + ":J" + (i + 7).ToString()].Merge();
                    ws.Range["K" + (i + 7).ToString() + ":L" + (i + 7).ToString()].Merge();
                    ws.Range["M" + (i + 7).ToString() + ":N" + (i + 7).ToString()].Merge();
                    ws.Range["O" + (i + 7).ToString() + ":P" + (i + 7).ToString()].Merge();
                }
                //Excel border's line style
                ws.Range["A1:P79"].Borders.LineStyle = XlLineStyle.xlContinuous;

                ws.Range["A1:P2"].Cells.Interior.Color = XlRgbColor.rgbMediumAquamarine;
                ws.Range["A3:P4"].Cells.Interior.Color = XlRgbColor.rgbMistyRose;
                ws.Range["C5:P7"].Cells.Interior.Color = XlRgbColor.rgbLightGoldenrodYellow;
                ws.Range["A8:B79"].Cells.Interior.Color = XlRgbColor.rgbHoneydew;
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
                ws.Range["A8"].Value = "08:00 ~ 08:10";
                ws.Range["A9"].Value = "08:10 ~ 08:20";
                ws.Range["A10"].Value = "08:20 ~ 08:30";
                ws.Range["A11"].Value = "08:30 ~ 08:40";
                ws.Range["A12"].Value = "08:40 ~ 08:50";
                ws.Range["A13"].Value = "08:50 ~ 09:00";
                ws.Range["A14"].Value = "09:00 ~ 09:10";
                ws.Range["A15"].Value = "09:10 ~ 09:20";
                ws.Range["A16"].Value = "09:20 ~ 09:30";
                ws.Range["A17"].Value = "09:30 ~ 09:40";
                ws.Range["A18"].Value = "09:40 ~ 09:50";
                ws.Range["A19"].Value = "09:50 ~ 10:00";
                ws.Range["A20"].Value = "10:00 ~ 10:10";
                ws.Range["A21"].Value = "10:10 ~ 10:20";
                ws.Range["A22"].Value = "10:20 ~ 10:30";
                ws.Range["A23"].Value = "10:30 ~ 10:40";
                ws.Range["A24"].Value = "10:40 ~ 10:50";
                ws.Range["A25"].Value = "10:50 ~ 11:00";
                ws.Range["A26"].Value = "11:00 ~ 11:10";
                ws.Range["A27"].Value = "11:10 ~ 11:20";
                ws.Range["A28"].Value = "11:20 ~ 11:30";
                ws.Range["A29"].Value = "11:30 ~ 11:40";
                ws.Range["A30"].Value = "11:40 ~ 11:50";
                ws.Range["A31"].Value = "11:50 ~ 12:00";
                ws.Range["A32"].Value = "13:00 ~ 13:10";
                ws.Range["A33"].Value = "13:10 ~ 13:20";
                ws.Range["A34"].Value = "13:20 ~ 13:30";
                ws.Range["A35"].Value = "13:30 ~ 13:40";
                ws.Range["A36"].Value = "13:40 ~ 13:50";
                ws.Range["A37"].Value = "13:50 ~ 14:00";
                ws.Range["A38"].Value = "14:00 ~ 14:10";
                ws.Range["A39"].Value = "14:10 ~ 14:20";
                ws.Range["A40"].Value = "14:20 ~ 14:30";
                ws.Range["A41"].Value = "14:30 ~ 14:40";
                ws.Range["A42"].Value = "14:40 ~ 14:50";
                ws.Range["A43"].Value = "14:50 ~ 15:00";
                ws.Range["A44"].Value = "15:00 ~ 15:10";
                ws.Range["A45"].Value = "15:10 ~ 15:20";
                ws.Range["A46"].Value = "15:20 ~ 15:30";
                ws.Range["A47"].Value = "15:30 ~ 15:40";
                ws.Range["A48"].Value = "15:40 ~ 15:50";
                ws.Range["A49"].Value = "15:50 ~ 16:00";
                ws.Range["A50"].Value = "16:00 ~ 16:10";
                ws.Range["A51"].Value = "16:10 ~ 16:20";
                ws.Range["A52"].Value = "16:20 ~ 16:30";
                ws.Range["A53"].Value = "16:30 ~ 16:40";
                ws.Range["A54"].Value = "16:40 ~ 16:50";
                ws.Range["A55"].Value = "16:50 ~ 17:00";
                ws.Range["A56"].Value = "17:00 ~ 17:10";
                ws.Range["A57"].Value = "17:10 ~ 17:20";
                ws.Range["A58"].Value = "17:20 ~ 17:30";
                ws.Range["A59"].Value = "17:30 ~ 17:40";
                ws.Range["A60"].Value = "17:40 ~ 17:50";
                ws.Range["A61"].Value = "17:50 ~ 18:00";
                ws.Range["A62"].Value = "18:00 ~ 18:10";
                ws.Range["A63"].Value = "18:10 ~ 18:20";
                ws.Range["A64"].Value = "18:20 ~ 18:30";
                ws.Range["A65"].Value = "18:30 ~ 18:40";
                ws.Range["A66"].Value = "18:40 ~ 18:50";
                ws.Range["A67"].Value = "18:50 ~ 19:00";
                ws.Range["A68"].Value = "19:00 ~ 19:10";
                ws.Range["A69"].Value = "19:10 ~ 19:20";
                ws.Range["A70"].Value = "19:20 ~ 19:30";
                ws.Range["A71"].Value = "19:30 ~ 19:40";
                ws.Range["A72"].Value = "19:40 ~ 19:50";
                ws.Range["A73"].Value = "19:50 ~ 20:00";
                ws.Range["A74"].Value = "20:00 ~ 20:10";
                ws.Range["A75"].Value = "20:10 ~ 20:20";
                ws.Range["A76"].Value = "20:20 ~ 20:30";
                ws.Range["A77"].Value = "20:30 ~ 20:40";
                ws.Range["A78"].Value = "20:40 ~ 20:50";
                ws.Range["A79"].Value = "20:50 ~ 21:00";
                #endregion
                ws.Range["A1:P79"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

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
