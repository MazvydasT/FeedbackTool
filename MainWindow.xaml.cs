using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using CefSharp;
using CefSharp.WinForms;

using FontAwesome.WPF;

using Newtonsoft.Json;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FeedbackTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ObservableCollection<UserData> dataGridSource = new ObservableCollection<UserData>();
        ChromiumWebBrowser browser;
        static string emailSelectionType { get; set; }

        public class JSObject
        {
            public System.Windows.Controls.DataGrid dataGridToQuery { get; set; }
            public System.Windows.Controls.Grid browserHostGrid { get; set; }

            public void cancel()
            {
                browserHostGrid.Dispatcher.Invoke(new Action(() =>
                {
                    browserHostGrid.Visibility = System.Windows.Visibility.Hidden;
                }));
            }

            public string getData()
            {
                return (string)dataGridToQuery.Dispatcher.Invoke(new Func<string>(() =>
                 {
                     if (emailSelectionType == "all")
                     {
                         return JsonConvert.SerializeObject(dataGridToQuery.Items);
                     }

                     else if (emailSelectionType == "selected")
                     {
                         return JsonConvert.SerializeObject(dataGridToQuery.SelectedItems);
                     }

                     var unselectedItems = new List<object>();

                     foreach (var item in dataGridToQuery.Items)
                     {
                         if (!dataGridToQuery.SelectedItems.Contains(item))
                         {
                             unselectedItems.Add(item);
                         }
                     }

                     return JsonConvert.SerializeObject(unselectedItems);
                 }));
            }
        }
        
        public MainWindow()
        {
            InitializeComponent();

            emailSelectionType = "all";

            updateButtons();

            dataGridSource.CollectionChanged += dataGridSource_CollectionChanged;

            dataGrid.ItemsSource = dataGridSource;

            datePickerTo.SelectedDate = DateTime.Now;

            double daysBetween = Properties.Settings.Default.daysBetween;

            if (daysBetween == 0)
            {
                datePickerFrom.SelectedDate = ((DateTime)datePickerTo.SelectedDate).AddYears(-1);
            }

            else
            {
                datePickerFrom.SelectedDate = ((DateTime)datePickerTo.SelectedDate).AddDays(-1 * daysBetween);
            }

            datePickerTo.SelectedDateChanged += datePickerTo_SelectedDateChanged;
            datePickerFrom.SelectedDateChanged += datePickerFrom_SelectedDateChanged;

            updateDataGrid();
            
            
            WindowsFormsHost host = new WindowsFormsHost();
            Cef.Initialize(new CefSettings(), true, true);
            browser = new ChromiumWebBrowser("");

            browser.RegisterJsObject("JSObject", new JSObject() { dataGridToQuery = dataGrid, browserHostGrid = browserHost }, false);

            browser.LoadingStateChanged += browser_LoadingStateChanged;

            host.Child = browser;
            browserHost.Children.Add(host);
        }

        void dataGridSource_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            if (dataGridSource.Count == 0)
            {
                goButton.IsEnabled = false;
            }

            else
            {
                goButton.IsEnabled = true;
            }
        }

        void browser_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
        {
            if (e.IsLoading)
            {
                 this.Dispatcher.Invoke(new Action(() =>
                 {
                     loadingCover.Visibility = System.Windows.Visibility.Visible;
                     browserHost.Visibility = System.Windows.Visibility.Hidden;
                 }));
            }

            else
            {
                this.Dispatcher.Invoke(new Action(() =>
                {
                    browserHost.Visibility = System.Windows.Visibility.Visible;
                    loadingCover.Visibility = System.Windows.Visibility.Hidden;
                }));
            }
        }

        void updateButtons()
        {
            string mainGoAction = Properties.Settings.Default.mainGoAction;

            if (mainGoAction == "email")
            {
                goActionMenuItem.Header = "Excel";
                goActionMenuItemIcon.Icon = FontAwesomeIcon.FileExcelOutline;

                goActionIcon.Icon = FontAwesomeIcon.Envelope;
            }

            else if (mainGoAction == "excel")
            {
                goActionMenuItem.Header = "E-mail";
                goActionMenuItemIcon.Icon = FontAwesomeIcon.Envelope;

                goActionIcon.Icon = FontAwesomeIcon.FileExcelOutline;
            }
        }

        

        void datePickerFrom_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            datePickerFrom.SelectedDateChanged -= datePickerFrom_SelectedDateChanged;

            updateDataGrid();

            datePickerFrom.SelectedDateChanged += datePickerFrom_SelectedDateChanged;
        }

        void datePickerTo_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            datePickerTo.SelectedDateChanged -= datePickerTo_SelectedDateChanged;

            updateDataGrid();

            datePickerTo.SelectedDateChanged += datePickerTo_SelectedDateChanged;
        }

        private void DG_Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Documents.Hyperlink link = (System.Windows.Documents.Hyperlink)e.OriginalSource;
            Process.Start(link.NavigateUri.AbsoluteUri);
        }

        private void updateDataGrid()
        {
            loadingCover.Visibility = System.Windows.Visibility.Visible;

            Task updateTask = new Task(() =>
            {
                DateTime from = (DateTime)this.Dispatcher.Invoke(new Func<DateTime>(() =>
                {
                    return (DateTime)datePickerFrom.SelectedDate;

                }));

                DateTime to = (DateTime)this.Dispatcher.Invoke(new Func<DateTime>(() =>
                {
                    return (DateTime)datePickerTo.SelectedDate;

                }));

                double dayDifference = (to - from).TotalDays;

                this.Dispatcher.Invoke(new Action(() =>
                {
                    labelTimeGap.Content = dayDifference + " days";
                }));

                Properties.Settings.Default.daysBetween = dayDifference;
                Properties.Settings.Default.Save();

                DataTable mainDataTable = new DataTable();
                DataTable noFeedbackSimsDataTable = new DataTable();

                using (SqlConnection connection = new SqlConnection("Data Source=GAY02016;Initial Catalog=VirtualManEng;Integrated Security=True"))
                {

                    SqlCommand mainCommand = new SqlCommand("SELECT DISTINCT vbesimulations_new.oemail, (CAST(COUNT(vbecssreport.refno) AS float)/CAST(COUNT(vbesimulations_new.oemail) AS float))*100.0 AS feedbackRate, COUNT(vbecssreport.refno) AS numberOfFeedbacks, COUNT(vbesimulations_new.oemail) - COUNT(vbecssreport.refno) AS numberOfNoFeedbacks FROM vbesimulations_new LEFT JOIN vbecssreport ON vbesimulations_new.refno = vbecssreport.refno WHERE (vbesimulations_new.issuedate BETWEEN @fromDate AND @toDate) AND (vbesimulations_new.project NOT LIKE '%test%') AND (vbesimulations_new.speciality = 'TF') AND (LTRIM(RTRIM(vbesimulations_new.refno)) <> '') GROUP BY vbesimulations_new.oemail HAVING COUNT(vbesimulations_new.oemail) - COUNT(vbecssreport.refno) > 0 ORDER BY feedbackRate,  numberOfNoFeedbacks DESC");
                    SqlCommand noFeedbackSimsCommand = new SqlCommand("SELECT vbesimulations_new.oemail, vbesimulations_new.refno, vbesimulations_new.issuedate, vbesimulations_new.simname, vbesimulations_new.project, vbesimulations_new.speciality FROM vbesimulations_new LEFT JOIN vbecssreport ON vbesimulations_new.refno = vbecssreport.refno WHERE (vbesimulations_new.issuedate BETWEEN @fromDate AND @toDate) AND (vbesimulations_new.project NOT LIKE '%test%') AND (vbecssreport.refno IS NULL) AND (vbesimulations_new.speciality = 'TF') AND (LTRIM(RTRIM(vbesimulations_new.refno)) <> '')");

                    mainCommand.Connection = connection;
                    mainCommand.CommandType = System.Data.CommandType.Text;

                    noFeedbackSimsCommand.Connection = connection;
                    noFeedbackSimsCommand.CommandType = System.Data.CommandType.Text;

                    mainCommand.Parameters.AddWithValue("@fromDate", from);
                    mainCommand.Parameters.AddWithValue("@toDate", to);

                    noFeedbackSimsCommand.Parameters.AddWithValue("@fromDate", from);
                    noFeedbackSimsCommand.Parameters.AddWithValue("@toDate", to);

                    try
                    {
                        connection.Open();

                        using (SqlDataReader dataReaderMainCommand = mainCommand.ExecuteReader())
                        {
                            mainDataTable.Load(dataReaderMainCommand);
                        }


                        using (SqlDataReader dataReaderNoFeedbackSimsCommand = noFeedbackSimsCommand.ExecuteReader())
                        {
                            noFeedbackSimsDataTable.Load(dataReaderNoFeedbackSimsCommand);
                        }
                    }

                    catch (Exception exception)
                    {
                        System.Windows.MessageBox.Show(exception.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                    }
                }

                this.Dispatcher.Invoke(new Action(() =>
                {
                    dataGridSource.Clear();
                }));

                foreach (DataRow row in mainDataTable.Rows)
                {
                    DataRow[] simRequests = noFeedbackSimsDataTable.Select("oemail = '" + row["oemail"] + "'", "issuedate DESC");

                    List<SimRequestData> simRequestsList = new List<SimRequestData>();

                    foreach (DataRow simRequest in simRequests)
                    {
                        simRequestsList.Add(new SimRequestData() { reference = simRequest["refno"], publishDate = simRequest["issuedate"], title = simRequest["simname"], project = simRequest["project"], speciality = simRequest["speciality"] });
                    }

                    this.Dispatcher.Invoke(new Action(() =>
                    {
                        dataGridSource.Add(new UserData() { email = row["oemail"], rate = row["feedbackRate"], feedbacks = row["numberOfFeedbacks"], noFeedbacks = row["numberOfNoFeedbacks"], simRequests = simRequestsList });
                    }));
                }
            });
            
            updateTask.ContinueWith(task => { loadingCover.Visibility = System.Windows.Visibility.Hidden; }, TaskScheduler.FromCurrentSynchronizationContext());

            updateTask.Start();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string mainGoAction = Properties.Settings.Default.mainGoAction;

            if (mainGoAction == "email")
            {
                selectedOption.IsEnabled = true;
                notSelectedOption.IsEnabled = true;

                if (dataGrid.SelectedItems.Count == 0)
                {
                    selectedOption.IsEnabled = false;
                }

                if (dataGrid.SelectedItems.Count == dataGrid.Items.Count)
                {
                    notSelectedOption.IsEnabled = false;
                }

                System.Windows.Controls.Button senderButton = sender as System.Windows.Controls.Button;
                System.Windows.Controls.ContextMenu contextMenu = senderButton.ContextMenu;

                contextMenu.IsEnabled = true;
                contextMenu.PlacementTarget = senderButton;
                contextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Left;
                contextMenu.IsOpen = true;
            }

            else if (mainGoAction == "excel")
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";

                if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(saveFileDialog.FileName, SpreadsheetDocumentType.Workbook))
                        {
                            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                            workbookPart.Workbook = new Workbook();

                            WorkbookStylesPart workbookStylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                            workbookStylePart.Stylesheet = new Stylesheet(
                                new DocumentFormat.OpenXml.Spreadsheet.Fonts(
                                    new DocumentFormat.OpenXml.Spreadsheet.Font(),
                                    new DocumentFormat.OpenXml.Spreadsheet.Font() { Color = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = 10 }, Underline = new DocumentFormat.OpenXml.Spreadsheet.Underline() },
                                    new DocumentFormat.OpenXml.Spreadsheet.Font() { Bold = new DocumentFormat.OpenXml.Spreadsheet.Bold() }),
                                new Fills(
                                    new Fill()),
                                new Borders(
                                    new DocumentFormat.OpenXml.Spreadsheet.Border()),
                                new CellFormats(
                                    new CellFormat() { NumberFormatId = 0, FormatId = 0, FontId = 0, BorderId = 0, FillId = 0 },
                                    new CellFormat() { NumberFormatId = 14, FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyNumberFormat = BooleanValue.FromBoolean(true) },
                                    new CellFormat() { NumberFormatId = 10, FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyNumberFormat = BooleanValue.FromBoolean(true) },
                                    new CellFormat() { NumberFormatId = 0, FormatId = 0, FontId = 1, BorderId = 0, FillId = 0 },
                                    new CellFormat() { NumberFormatId = 0, FormatId = 0, FontId = 2, BorderId = 0, FillId = 0 }));
                            workbookStylePart.Stylesheet.Save();

                            Hyperlinks hyperlinks = new Hyperlinks();

                            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                            worksheetPart.Worksheet = new Worksheet(new SheetData());

                            worksheetPart.Worksheet.Append(hyperlinks);

                            worksheetPart.Worksheet.Save();

                            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Table" };
                            sheets.Append(sheet);

                            SharedStringTablePart sharedStringTablePart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
                            sharedStringTablePart.SharedStringTable = new SharedStringTable();

                            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                            sheetData.Append(new Row(
                                new Cell()
                                {
                                    CellValue = new CellValue(insertSharedString(sharedStringTablePart, "Requestor E-mail")),
                                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                                    StyleIndex = 4
                                },

                                new Cell()
                                {
                                    CellValue = new CellValue(insertSharedString(sharedStringTablePart, "Feedback Rate")),
                                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                                    StyleIndex = 4
                                },

                                new Cell()
                                {
                                    CellValue = new CellValue(insertSharedString(sharedStringTablePart, "Number Of Feedbacks")),
                                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                                    StyleIndex = 4
                                },

                                new Cell()
                                {
                                    CellValue = new CellValue(insertSharedString(sharedStringTablePart, "Number Of No Feedbacks")),
                                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                                    StyleIndex = 4
                                },

                                new Cell()
                                {
                                    CellValue = new CellValue(insertSharedString(sharedStringTablePart, "Reference")),
                                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                                    StyleIndex = 4
                                },

                                new Cell()
                                {
                                    CellValue = new CellValue(insertSharedString(sharedStringTablePart, "Publish Date")),
                                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                                    StyleIndex = 4
                                },

                                new Cell()
                                {
                                    CellValue = new CellValue(insertSharedString(sharedStringTablePart, "Title")),
                                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                                    StyleIndex = 4
                                }));

                            int rowCounter = 1;

                            foreach (UserData userData in dataGridSource)
                            {
                                foreach (SimRequestData simRequestData in userData.simRequests)
                                {
                                    rowCounter++;

                                    sheetData.Append(new Row(
                                        new Cell()
                                        {
                                            CellValue = new CellValue(insertSharedString(sharedStringTablePart, userData.email.ToString())),
                                            DataType = new EnumValue<CellValues>(CellValues.SharedString)
                                        },

                                        new Cell()
                                        {
                                            CellValue = new CellValue((Convert.ToDouble(userData.rate) / 100d).ToString()),
                                            StyleIndex = 2
                                        },

                                        new Cell()
                                        {
                                            CellValue = new CellValue(userData.feedbacks.ToString())
                                        },

                                        new Cell()
                                        {
                                            CellValue = new CellValue(userData.noFeedbacks.ToString())
                                        },

                                        new Cell()
                                        {
                                            CellValue = new CellValue(insertSharedString(sharedStringTablePart, simRequestData.reference.ToString())),
                                            DataType = new EnumValue<CellValues>(CellValues.SharedString),
                                            StyleIndex = 3
                                        },

                                        new Cell()
                                        {
                                            CellValue = new CellValue(Convert.ToDateTime(simRequestData.publishDate).ToOADate().ToString()),
                                            StyleIndex = 1
                                        },

                                        new Cell()
                                        {
                                            CellValue = new CellValue(insertSharedString(sharedStringTablePart, simRequestData.title.ToString())),
                                            DataType = new EnumValue<CellValues>(CellValues.SharedString)
                                        }
                                    ));

                                    string hyperlinkID = "id" + rowCounter.ToString();

                                    hyperlinks.Append(new DocumentFormat.OpenXml.Spreadsheet.Hyperlink() { Reference = "E" + rowCounter.ToString(), Id = hyperlinkID });
                                    worksheetPart.AddHyperlinkRelationship(simRequestData.requestLink, true, hyperlinkID);
                                }
                            }

                            workbookPart.Workbook.Save();
                        }
                    }

                    catch (Exception exception)
                    {
                        System.Windows.MessageBox.Show(exception.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                    }
                }
            }
        }

        private string insertSharedString(SharedStringTablePart sharedStringTablePart, string stringToInsert)
        {
            int index = 0;

            var sharedStringTable = sharedStringTablePart.SharedStringTable;

            foreach (var item in sharedStringTable.Elements<SharedStringItem>())
            {
                if(item.InnerText == stringToInsert) {
                    return index.ToString();
                }

                index++;
            }

            sharedStringTable.AppendChild(new SharedStringItem(new Text(stringToInsert)));
            sharedStringTable.Save();

            return index.ToString();
        }

        private void actionSendEmail()
        {
            browser.GetMainFrame().LoadUrl("https://script.google.com/a/macros/jaguarlandrover.com/s/AKfycbyRFZLV_0KiUKoDZUI-gt9e7G6MWyoQS8wBFrFlodPREk5BFqw/exec");

            browserHost.Visibility = System.Windows.Visibility.Visible;
        }

        private void Window_Closing_1(object sender, CancelEventArgs e)
        {
            this.Hide();
            
            Cef.Shutdown();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button senderButton = sender as System.Windows.Controls.Button;
            System.Windows.Controls.ContextMenu contextMenu = senderButton.ContextMenu;

            contextMenu.IsEnabled = true;
            contextMenu.PlacementTarget = senderButton;
            contextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Left;
            contextMenu.HorizontalOffset = senderButton.ActualWidth;
            contextMenu.VerticalOffset = senderButton.ActualHeight;
            contextMenu.IsOpen = true;
        }

        private void goActionMenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            string mainGoAction = Properties.Settings.Default.mainGoAction;

            if (mainGoAction == "email")
            {
                Properties.Settings.Default.mainGoAction = "excel";
            }

            else if (mainGoAction == "excel")
            {
                Properties.Settings.Default.mainGoAction = "email";
            }

            Properties.Settings.Default.Save();
            updateButtons();
        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            emailSelectionType = "selected";

            actionSendEmail();
        }

        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            emailSelectionType = "all";

            actionSendEmail();
        }

        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            emailSelectionType = "notselected";

            actionSendEmail();
        }

    }

    public class UserData
    {
        public object email { get; set; }
        public object rate { get; set; }
        public object feedbacks { get; set; }
        public object noFeedbacks { get; set; }

        public List<SimRequestData> simRequests { get; set; }
    }

    public class SimRequestData
    {
        public object reference { get; set; }
        public object publishDate { get; set; }
        public string publishDateString
        {
            get
            {
                return ((DateTime)publishDate).ToString("dd/MM/yyyy");
            }
        }
        public object title { get; set; }
        public object project { get; set; }
        public object speciality { get; set; }

        private string refNoSlashes
        {
            get
            {
                return reference.ToString().Replace("/","");
            }
        }

        public Uri requestLink
        {
            get
            {
                return new Uri(String.Format(@"http://apps.pag.ford.com/virtualmaneng/vbereports/project_filter_2_2_conn_new.asp?reference={0}&sproject={1}&projectFolder=\\gay02016\prod$\AppsFarm\VME\ebc_admin\reports\{1}\{2}\{3}\", reference, project, speciality, refNoSlashes));
            }
        }
    }

    /*public class CustomMenuHandler : IMenuHandler
    {
        public bool OnBeforeContextMenu(IWebBrowser browser)
        {
            return false;
        }
    }*/
}
