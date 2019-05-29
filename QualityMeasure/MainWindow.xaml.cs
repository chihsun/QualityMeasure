using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ClosedXML.Excel;

namespace QualityMeasure
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        #region Definition
        public List<DiliveryData> gdata = new List<DiliveryData>();
        public List<Measurement> gmeasure = new List<Measurement>();
        public List<MElements> gcollect = new List<MElements>();
        public Dictionary<string, List<MElements>> gbackups = new Dictionary<string, List<MElements>>();
        public Dictionary<string, List<string>> gduplicate = new Dictionary<string, List<string>>();
        public List<string> SameEle = new List<string>();
        #endregion
        #region Class
        public class DiliveryData
        {
            public string Group { get; set; }
            public string Depart { get; set; }
            public string MeasureID { get; set; }
            public string MeasureName { get; set; }
            public string Frequency { get; set; }
            public string User { get; set; }

            public List<string> SameID = new List<string>();
        }

        public class Measurement
        {
            public string Group { get; set; }
            public string MeasureID { get; set; }
            public string MeasureName { get; set; }
            public string Numerator { get; set; }
            public string Denominator { get; set; }
            public string Threshold { get; set; }
            public string Frequency { get; set; }
            public string User { get; set; }
        }

        public class MElements
        {
            public string Element { get; set; }
            public DateTime Eledate { get; set; }
            public string ElementData { get; set; }
            public Dictionary<string, string> PreDdatas { get; set; }
            public MElements()
            {
                PreDdatas = new Dictionary<string, string>();
                Eledate = new DateTime();
            }
        }

        /*public class FileStatusHelper
        {
            [DllImport("kernel32.dll")]
            public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

            [DllImport("kernel32.dll")]
            public static extern bool CloseHandle(IntPtr hObject);

            public const int OF_READWRITE = 2;
            public const int OF_SHARE_DENY_NONE = 0x40;
            public static readonly IntPtr HFILE_ERROR = new IntPtr(-1);

            /// <summary>
            /// 查看檔案是否被佔用
            /// </summary>
            /// <param name="filePath"></param>
            /// <returns></returns>
            public static bool IsFileOccupied(string filePath)
            {
                IntPtr vHandle = _lopen(filePath, OF_READWRITE | OF_SHARE_DENY_NONE);
                CloseHandle(vHandle);
                return vHandle == HFILE_ERROR ? true : false;
            }
        }
        */
        #endregion
        private void BT_IMPORT_SOURCE(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.InitialDirectory = Environment.CurrentDirectory;
            dlg.Title = "選取資料檔";
            dlg.Filter = "xlsx files (*.*)|*.xlsx";
            if (dlg.ShowDialog() == true)
            {
                LoadFile(dlg.FileName);
            }
        }
        public void LoadFile(string fname)
        {
            if (!System.IO.File.Exists(fname))
                return;
            gdata.Clear();
            gmeasure.Clear();
            gduplicate.Clear();
            SameEle.Clear();
            gbackups.Clear();
            try
            {
                using (var wb = new XLWorkbook(fname))
                {
                    var ws = wb.Worksheet("工作表1");
                    
                    for (int i = 0; i < 500; i++)
                    {
                        if (string.IsNullOrEmpty(ws.Cell(i + 2, 1).GetString()))
                            break;
                        if (string.IsNullOrEmpty(ws.Cell(i + 2, 2).GetString()))
                            continue;
                        DiliveryData data = new DiliveryData
                        {
                            Group = ws.Cell(i + 2, 1).GetString().Trim(),
                            Depart = ws.Cell(i + 2, 2).GetString().Trim(),
                            MeasureID = ws.Cell(i + 2, 3).GetString().Trim(),
                            MeasureName = ws.Cell(i + 2, 5).GetString().Trim(),
                            User = ws.Cell(i + 2, 7).GetString().Trim()
                        };
                        for (int j = 12; j < 15; j++)
                        {
                            string content = ws.Cell(i + 2, j).GetString().Trim();
                            if (string.IsNullOrEmpty(content))
                                break;

                            if (SameEle.Contains(content) || (gduplicate.ContainsKey(content) && gduplicate[content].Contains(data.MeasureID)))
                                continue;

                            SameEle.Add(content);

                            if (gduplicate.ContainsKey(data.MeasureID))
                            {
                                if (!gduplicate[data.MeasureID].Contains(content))
                                {
                                    gduplicate[data.MeasureID].Add(content);
                                }
                            }
                            else
                            {
                                gduplicate.Add(data.MeasureID, new List<string>() { content });
                            }
                        }
                        gdata.Add(data);
                    }

                    if (gdata.Count > 0)
                    {
                        MessageBox.Show("匯入成功 : " + gdata.Count.ToString());
                        TxtBox1.Text += Environment.NewLine + "指標匯入數量 : " + gdata.Count + Environment.NewLine;
                        if (gduplicate.Count > 0)
                        {
                            TxtBox1.Text += Environment.NewLine + "相同意義要素組數量 : " + gduplicate.Count + Environment.NewLine;
                        }
                    }
                    else
                    {
                        MessageBox.Show("匯入失敗");
                    }

                    var ws2 = wb.Worksheet("工作表2");

                    for (int i = 0; i < 500; i++)
                    {
                        if (string.IsNullOrEmpty(ws2.Cell(i + 2, 1).GetString()))
                            break;
                        if (string.IsNullOrEmpty(ws2.Cell(i + 2, 2).GetString()))
                            continue;
                        Measurement data = new Measurement
                        {
                            Group = ws2.Cell(i + 2, 1).GetString().Trim(),
                            MeasureID = ws2.Cell(i + 2, 2).GetString().Trim(),
                            MeasureName = ws2.Cell(i + 2, 3).GetString().Trim(),
                            Numerator = ws2.Cell(i + 2, 4).GetString().Trim(),
                            Denominator = ws2.Cell(i + 2, 6).GetString().Trim()
                        };

                        gmeasure.Add(data);
                    }

                    MessageBox.Show("匯入成功 : " + gmeasure.Count.ToString());

                    wb.Dispose();

                    LoadDataBASE();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadDataBASE()
        {
            string fpath = Environment.CurrentDirectory + @"\要素備份";
            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            string fname = fpath + @"\指標收集存檔總檔.xlsx";
            if (!System.IO.File.Exists(fname))
                return;
            using (var wb = new XLWorkbook(fname))
            {
                var ws = wb.Worksheet(1);
                for (int i = 0; i < 500; i++)
                {
                    if (string.IsNullOrEmpty(ws.Cell(i + 2, 1).GetString()))
                        break;
                    //if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 2)))
                    //    continue;

                    if (gbackups.Count > 0 && gbackups.ContainsKey(ws.Cell(i + 2, 1).GetString()))
                    {
                        continue;
                    }
                    List<MElements> lme = new List<MElements>();
                    List<string> duplicate = new List<string>();
                    for (int j = 0; j < 12; j++)
                    {
                        if (string.IsNullOrEmpty(ws.Cell(1, j + 2).GetString()))
                            break;
                        if (string.IsNullOrEmpty(ws.Cell(2, j + 2).GetString()))
                            continue;
                        if (!DateTime.TryParse(ws.Cell(1, j + 2).GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dts))
                            continue;
                        if (dts > DateTime.Now.AddMonths(-13) && dts < DateTime.Now)
                        {
                            if (duplicate.Contains(dts.ToString("yyyy/MM")))
                                continue;
                            else
                                duplicate.Add(dts.ToString("yyyy/MM"));

                            MElements data = new MElements
                            {
                                Element = ws.Cell(i + 2, 1).GetString().Trim(),
                                ElementData = ws.Cell(i + 2, j + 2).GetString().Trim(),
                                Eledate = dts
                            };
                            lme.Add(data);
                        }
                    }
                    gbackups[ws.Cell(i + 2, 1).GetString().Trim()] = lme;

                    try
                    {
                        if (gduplicate.ContainsKey(ws.Cell(i + 2, 1).GetString().Trim()))
                        {
                            var glists = gduplicate.Where(o => o.Key == ws.Cell(i + 2, 1).GetString().Trim()).FirstOrDefault().Value;
                            foreach (var x in glists)
                            {
                                if (!gbackups.ContainsKey(x) &&
                                    gdata.Find(o => o.MeasureID == x) != null)
                                {
                                    gbackups[x] = lme;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message);
                    }
                }

                wb.Dispose();
            }
        }

            private void BT_TO_EXPORT_CLINIC(object sender, RoutedEventArgs e)
        {

        }

        private void BT_IMPORT_RESULT(object sender, RoutedEventArgs e)
        {

        }

        private void BT_IMPORT_OLDDATA(object sender, RoutedEventArgs e)
        {

        }

        private void BT_IMPORT_MEASUREDATA(object sender, RoutedEventArgs e)
        {

        }

        private void BT_EXPORT_ELEMENT(object sender, RoutedEventArgs e)
        {
            if (gbackups.Count <= 0)
                return;

            if (gcollect.Count > 0)
            {
                foreach (var x in gcollect)
                {
                    if (gbackups.ContainsKey(x.Element))
                    {
                        var dataex = gbackups[x.Element].FirstOrDefault(o => o.Element == x.Element && o.Eledate == x.Eledate);
                        if (dataex != null)
                            gbackups[x.Element].Remove(dataex);
                        gbackups[x.Element].Add(x);
                    }
                }
            }

            var sortbacks = gbackups.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value);

            string fpath = Environment.CurrentDirectory + @"\要素備份";

            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            string fname = @"\指標收集總存檔" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM", CultureInfo.InvariantCulture) + ".xlsx";
            //string fname2 = @"\指標收集存檔總檔.xlsx";
            try
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("工作表1");
                    ws.Style.Font.FontSize = 12;
                    ws.Style.Font.FontName = "微軟正黑體";

                    var wscol = ws.Columns("A");
                    wscol.Width = 15;
                    wscol.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                    ws.Cell(1, 1).Value = "指標要素";

                    for (int i = 0; i < 6; i++)
                    {
                        ws.Cell(1, i + 2).Value = DateTime.Now.AddMonths(-1 - i).ToString("yyyy/MM");
                        ws.Cell(1, i + 2).Style.DateFormat.Format = "yyyy/MM";
                        ws.Cell(1, i + 2).Style.Fill.BackgroundColor = XLColor.LightBlue;
                        ws.Cell(1, i + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Column(i + 2).AdjustToContents();
                    }
                    int index = 0;
                    foreach (var x in sortbacks)
                    {
                        ws.Cell(index + 2, 1).Value = x.Key;
                        ws.Cell(index + 2, 1).Style.Fill.BackgroundColor = XLColor.LightCyan;
                        foreach (var y in x.Value)
                        {
                            for (int i = 0; i < 6; i++)
                            {
                                int num;
                                if (y.Eledate.Year == DateTime.Now.AddMonths(-1 - i).Year
                                    && y.Eledate.Month == DateTime.Now.AddMonths(-1 - i).Month && int.TryParse(y.ElementData, out num))
                                {
                                    ws.Cell(index + 2, i + 2).Value = num;
                                }
                            }
                        }

                        index++;
                    }

                    wb.SaveAs(fpath + fname);

                    MessageBox.Show("匯出要素備份成功");

                    wb.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void BT_EXPORT_MEASURE(object sender, RoutedEventArgs e)
        {
            if (gmeasure.Count <= 0 || gbackups.Count <= 0)
                return;
            string fpath = Environment.CurrentDirectory + @"\要素備份";

            if (!Directory.Exists(fpath))
            {
                Directory.CreateDirectory(fpath);
            }
            string fname = @"\指標數據總資料" + DateTime.Now.AddMonths(-1).ToString("yyyy-M", CultureInfo.InvariantCulture) + ".xlsx";

            var sortbacks = gbackups.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value.OrderByDescending(o => o.ElementData).ToList());
            try
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("工作表1");
                    ws.Style.Font.FontSize = 12;
                    ws.Style.Font.FontName = "微軟正黑體";

                    var wscol = ws.Columns("A:C");
                    wscol.Width = 15;
                    wscol.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wscol.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    wscol.Style.Alignment.WrapText = true;

                    ws.Cells("A1:C1").Style.Fill.BackgroundColor = XLColor.LightBlue;
                    ws.Cell(1, 1).Value = "指標群組";
                    ws.Cell(1, 2).Value = "指標代號";
                    ws.Cell(1, 3).Value = "指標名稱";

                    for (int i = 0; i < 12; i++)
                    {
                        ws.Cell(1, i + 4).Value = DateTime.Now.AddMonths(-i - 1).ToString("yyyy/MM");
                        ws.Cell(1, i + 4).Style.DateFormat.Format = "yyyy/MM";
                        ws.Cell(1, i + 4).Style.Fill.BackgroundColor = XLColor.LightBlue;
                        ws.Cell(1, i + 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Column(i + 4).AdjustToContents();
                    }
                    int index = 2;
                    foreach (var x in gmeasure)
                    {
                        ws.Cell(index, 1).Value = x.Group;
                        ws.Cell(index, 2).Value = x.MeasureID;
                        ws.Cell(index, 3).Value = x.MeasureName;
                        ws.Range(index, 1, index + 2, 1).Merge();
                        ws.Range(index, 2, index + 2, 2).Merge();
                        ws.Range(index, 3, index + 2, 3).Merge();
                        int status = 0;
                        var Numes = sortbacks.FirstOrDefault(o => o.Key == x.Numerator).Value;

                        List<List<MElements>> DenosPlus = new List<List<MElements>>();

                        var Denos = sortbacks.FirstOrDefault(o => o.Key == x.Denominator).Value;

                        if (x.Denominator.Contains("+"))
                        {
                            status = 1;
                            var elements = x.Denominator.Split('+').ToList();
                            if (elements.Count > 0)
                            {
                                foreach (var ele in elements)
                                {
                                    var em = sortbacks.FirstOrDefault(o => o.Key == ele).Value;
                                    if (em != null)
                                        DenosPlus.Add(em);
                                }
                            }
                        }
                        else if (x.Denominator.Contains(".") && x.Denominator.Contains("-"))
                        {
                            status = 2;
                            var elements = x.Denominator.Split('-').ToList();
                            if (elements.Count > 0)
                            {
                                foreach (var ele in elements)
                                {
                                    var em = sortbacks.FirstOrDefault(o => o.Key == ele).Value;
                                    if (em != null)
                                        DenosPlus.Add(em);
                                }
                            }
                        }
                        for (int i = 0; i < 12; i++)
                        {
                            if (x.Numerator == "1")
                                ws.Cell(index + 1, i + 4).Value = 1;
                            else if (Numes != null)
                            {
                                var nume = Numes.FirstOrDefault(o => o.Eledate.Year == DateTime.Now.AddMonths(-i - 1).Year
                                && o.Eledate.Month == DateTime.Now.AddMonths(-i - 1).Month);
                                if (nume != null)
                                {
                                    int numok;
                                    if (int.TryParse(nume.ElementData, out numok))
                                        ws.Cell(index + 1, i + 4).Value = numok;
                                }
                            }

                            if (x.Denominator == "1")
                                ws.Cell(index + 2, i + 4).Value = "NA";
                            else if (Denos != null)
                            {
                                var deno = Denos.FirstOrDefault(o => o.Eledate.Year == DateTime.Now.AddMonths(-i - 1).Year
                                && o.Eledate.Month == DateTime.Now.AddMonths(-i - 1).Month);
                                if (deno != null)
                                {
                                    int numok;
                                    if (int.TryParse(deno.ElementData, out numok))
                                        ws.Cell(index + 2, i + 4).Value = numok;
                                }
                            }
                            else if (DenosPlus.Count > 0)
                            {
                                try
                                {
                                    int deno = 0;
                                    foreach (var ele in DenosPlus)
                                    {
                                        var de = ele.FirstOrDefault(o => o.Eledate.Year == DateTime.Now.AddMonths(-i - 1).Year
                                    && o.Eledate.Month == DateTime.Now.AddMonths(-i - 1).Month);
                                        if (de == null)
                                        {
                                            break;
                                        }
                                        int num;
                                        if (!Int32.TryParse(de.ElementData, out num))
                                            break;
                                        if (status == 1)
                                        {
                                            deno += num;
                                        }
                                        else if (status == 2)
                                        {
                                            if (deno == 0)
                                                deno = num;
                                            else
                                                deno -= num;
                                        }
                                    }
                                    if (deno > 0)
                                        ws.Cell(index + 2, i + 4).Value = deno;
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }
                            }

                            if (!string.IsNullOrEmpty(ws.Cell(index + 2, i + 4).GetString()))
                            {
                                double nu, de;
                                if (double.TryParse(ws.Cell(index + 1, i + 4).GetString(), out nu)
                                    && double.TryParse(ws.Cell(index + 2, i + 4).GetString() == "NA" ? "1" : ws.Cell(index + 2, i + 4).GetString(), out de))
                                {
                                    if (de != 0)
                                        ws.Cell(index, i + 4).Value = nu / de;
                                }
                            }
                        }
                        index += 3;
                    }

                    wb.SaveAs(fpath + fname);

                    MessageBox.Show("指標匯出結束");
                    wb.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void BT_TO_TCPI(object sender, RoutedEventArgs e)
        {

        }

        private void BT_TO_HACMI(object sender, RoutedEventArgs e)
        {

        }

        private void BT_TO_THIS(object sender, RoutedEventArgs e)
        {

        }
    }
}
