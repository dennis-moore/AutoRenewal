using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using AutoRenewal.Models;
using System.Text.Json;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Diagnostics;
using AutoRenewal.Util;
using System.Text.RegularExpressions;
using System;
using Serilog;
using System.Text;

namespace AutoRenewal
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        public ObservableCollection<Organization> OrganizationList { get; private set; } = new ObservableCollection<Organization>();

        private Organization selectedOrganization;
        public Organization SelectedOrganization
        {
            get => selectedOrganization;
            set
            {
                if (value != selectedOrganization)
                {
                    selectedOrganization = value;
                    OnPropertyChanged();
                }
            }
        }

        private string inputPath = string.Empty;
        public string InputPath
        {
            get => inputPath; 
            set
            {
                if(inputPath != value)
                {
                    inputPath = value;
                    OnPropertyChanged();
                }
            }
        }

        public string progressText = string.Empty;
        public string ProgressText
        {
            get => progressText;
            set
            {
                if(progressText != value)
                {
                    progressText = value;
                    OnPropertyChanged();
                }
            }
        }

        private string OutputDirectory { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            Directory.CreateDirectory(Path.Combine(AppContext.BaseDirectory, "logs"));
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.Console()
                .WriteTo.File(Path.Combine("logs", "log.log"), rollOnFileSizeLimit: true)
                .CreateLogger();

            Log.Information("App Startup");

            OutputDirectory = Path.Combine(AppContext.BaseDirectory, "ED Analysis");
            Directory.CreateDirectory(OutputDirectory);
            Directory.CreateDirectory(Path.Combine(AppContext.BaseDirectory, "Configs"));

            // start a task to load in the runtime information
            Task.Run(async () =>
            {
                DirectoryInfo d = new DirectoryInfo(Path.Combine(AppContext.BaseDirectory, "Configs"));
                FileInfo[] Files = d.GetFiles("*.json");
                foreach (FileInfo file in Files)
                {
                    using (FileStream openStream = File.OpenRead(file.FullName))
                    {
                        var org = await JsonSerializer.DeserializeAsync<Organization>(openStream);
                        Dispatcher.Invoke(() => OrganizationList.Add(org));
                    }
                }
            });
        }

        private void InputBrowseBtnClick(object sender, RoutedEventArgs e)
        {
            var filename = OpenFileBrowser(".xlsx", "EXCEL Files (*.xlsx)|*.xlsx");
            InputPathTextBox.Text = filename;
        }

        private void StartBtnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                IWorkbook ExcelDoc = null;
                XWPFDocument WordDoc = null;
                var orgHelper = new OrgHelper(SelectedOrganization);
                var outputFileName = Path.GetFileNameWithoutExtension(InputPath);

                try
                {
                    using (FileStream file = new FileStream(InputPath, FileMode.Open, FileAccess.Read))
                    {
                        ExcelDoc = WorkbookFactory.Create(file);
                    }
                    string templatePath = Path.Combine(AppContext.BaseDirectory, "Templates", SelectedOrganization.TemplateFileName);
                    using (FileStream file = new FileStream(templatePath, FileMode.Open, FileAccess.ReadWrite))
                    {
                        WordDoc = new XWPFDocument(file);
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex.ToString());
                    ProgressText = "Unable to open input or output file";
                    MessageBox.Show(ex.Message);
                    return;
                }

                ProgressText = "In Progress";
                StringBuilder statusString = new StringBuilder();

                for (int i = 0; i < WordDoc.BodyElements.Count; i++)
                {
                    var element = WordDoc.BodyElements[i];
                    if (element.ElementType == BodyElementType.PARAGRAPH)
                    {
                        var paragraph = (XWPFParagraph)element;
                        if (paragraph.Text.Length > 0 && paragraph.Text[0] == '[') // only proceed if paragraph starts with '[', i.e. [Part 1]
                        {
                            // for each paragraph element, check if it has a mapping for the selected organization
                            var split = Regex.Split(paragraph.Text, @"(?<=[]])")[0];  //@"(?<=[.,;])"
                            var mapList = orgHelper.HasMapping(split);
                            if (mapList.Count > 0)
                            {
                                // fow now, i'm only considering the first map (fine for PS1 orgs)
                                var map = mapList[0];

                                var para = (XWPFParagraph)WordDoc.BodyElements[i + 1];
                                if(para.Text != string.Empty)
                                {
                                    statusString.AppendLine($"Unable to paste {map.ExcelCell} into {map.WordDesignator}. No empty line to paste into.");
                                    continue;
                                }

                                try
                                {
                                    ISheet sheet = ExcelDoc.GetSheet(map.SheetName);
                                    var rowVal = OrgHelper.GetRowValue(map.ExcelCell);
                                    var colVal = OrgHelper.GetColumnValue(map.ExcelCell);
                                    string contents = string.Empty;
                                    switch(sheet.GetRow(rowVal).GetCell(colVal).CellType)
                                    {
                                        case CellType.Blank:
                                            Log.Information($"blank cell found {map.ExcelCell}");
                                            continue;
                                        case CellType.Numeric:
                                            contents = sheet.GetRow(rowVal).GetCell(colVal).NumericCellValue.ToString();
                                            break;
                                        case CellType.String:
                                            contents = sheet.GetRow(rowVal).GetCell(colVal).StringCellValue;
                                            if (contents.Length < 7 && contents.Contains("N/A")) continue;
                                            break;
                                        case CellType.Unknown:
                                            statusString.AppendLine($"Badly formatted cell {map.SheetName} : {map.ExcelCell}");
                                            continue;
                                    }

                                    var run = para.CreateRun();
                                    run.SetText(contents);
                                }
                                catch(Exception ex)
                                {
                                    Log.Error(ex.ToString());
                                    statusString.AppendLine($"Unable to find {map.SheetName} : {map.ExcelCell}");
                                    continue;
                                }
                            }
                        }
                    }
                }

                FileStream sw = new FileStream(Path.Combine(OutputDirectory, $"{outputFileName}.docx"), FileMode.Create);
                WordDoc.Write(sw);
                sw.Close();

                ExcelDoc.Close();
                WordDoc.Close();

                if (statusString.ToString() == string.Empty)
                    ProgressText = "Completed Successfully";
                else
                {
                    statusString.AppendLine("Completed Successfully");
                    ProgressText = statusString.ToString();
                }
            }
            catch(Exception ex)
            {
                Log.Error(ex.ToString());
                ProgressText = "Error occurred";
                MessageBox.Show(ex.Message);
            }
        }

        private string OpenFileBrowser(string defaultExt, string filter)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = defaultExt;
            dlg.Filter = filter;

            var result = dlg.ShowDialog();

            return result ?? false ? dlg.FileName : string.Empty;
        }

        private void orgDropdown_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectedOrganization = (sender as ComboBox).SelectedItem as Organization;
        }

        private void StackPanel_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string file = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
                if (Path.GetExtension(file) == ".xlsx")
                    InputPath = file;
                Debug.WriteLine(file);
            }
        }

        private void StackPanel_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string file = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
                if(Path.GetExtension(file) != ".xlsx")
                    e.Effects = DragDropEffects.None;
            }
            else
                e.Effects = DragDropEffects.None;
        }

        private async void AddOrganization(object sender, RoutedEventArgs e)
        {
            // for now, create a new org object, serialize and save to file
            OrganizationList.Add(new Organization()
            {
                Name = "PS1",
                //ExcelTabNames = new List<string>(){ "501(c)(3)", "PS1" },
                TemplateFileName = "PS1_Template.docx",
                Mappings = new ObservableCollection<Mapping>
                {
                    new Mapping() { SheetName = "501(c)(3)", ExcelCell = "F4", WordDesignator = "[PART 1]" }
                }
            });

            string path = Path.Combine(AppContext.BaseDirectory, "Configs", "PS1.json");

            using (FileStream createStream = File.Create(path))
            {
                await JsonSerializer.SerializeAsync(createStream, OrganizationList[0], new JsonSerializerOptions { WriteIndented = true });
            }
        }
    }
}
