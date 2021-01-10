using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
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
using TextGrabber.Models;

namespace TextGrabber
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window //, INotifyPropertyChanged
    {
        public ObservableCollection<Relation> RelationsList { get; private set; } = new ObservableCollection<Relation>();

        //public event PropertyChangedEventHandler PropertyChanged;

        //protected void OnPropertyChanged([CallerMemberName] string name = null)
        //{
        //    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        //}

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
#if DEBUG
            InputPathTextBox.Text = @"C:\Users\denmo\Downloads\Book.xlsx";
            OutputPathTextBox.Text = @"C:\Users\denmo\Downloads\Document.docx";
#endif
        }

        private void InputBrowseBtnClick(object sender, RoutedEventArgs e)
        {
            // Set filter for file extension and default file extension 
            //dlg.DefaultExt = ".xlsx";
            //dlg.Filter = "JPEG Files (*.jpeg)|*.jpeg|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif";
            var filename = OpenFileBrowser(".xlsx", "EXCEL Files (*.xlsx)|*.xlsx");
            InputPathTextBox.Text = filename;
        }

        private void OutputBrowseBtnClick(object sender, RoutedEventArgs e)
        {
            // Set filter for file extension and default file extension 
            //dlg.DefaultExt = ".xlsx";
            //dlg.Filter = "JPEG Files (*.jpeg)|*.jpeg|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif";
            var filename = OpenFileBrowser(".docx", "WORD Files (*.docx)|*.docx");
            OutputPathTextBox.Text = filename;
        }

        private void AddRelation(object sender, RoutedEventArgs e)
        {
            RelationsList.Add(new Relation { excelCell = "F1", wordDesignator = "Organization" });
        }

        private void RemoveRelation(object sender, RoutedEventArgs e)
        {
            var btn = (Button)sender;
            var relation = (Relation)btn.DataContext;
            RelationsList.Remove(relation);
        }

        private void StartBtnClick(object sender, RoutedEventArgs e)
        {
            IWorkbook ExcelDoc;
            XWPFDocument WordDoc;
            using (FileStream file = new FileStream(InputPathTextBox.Text, FileMode.Open, FileAccess.Read))
            {
                ExcelDoc = WorkbookFactory.Create(file);
            }
            using (FileStream file = new FileStream(OutputPathTextBox.Text, FileMode.Open, FileAccess.ReadWrite))
            {
                WordDoc = new XWPFDocument(file);
            }

            StringBuilder aBuilder = new StringBuilder();
            StringBuilder dBuilder = new StringBuilder();

            foreach (IRow row in ExcelDoc.GetSheetAt(0))
            {
                string aValue = row.GetCell(0).StringCellValue;
                string dValue = row.GetCell(3).StringCellValue;
                
                if(!string.IsNullOrWhiteSpace(aValue))
                {
                    aBuilder.Append($" {aValue}");
                }
                if (!string.IsNullOrWhiteSpace(dValue))
                {
                    dBuilder.Append($" {dValue}");
                }
            }

            //IDictionary<int, string> RunPositions = new Dictionary<int, string>();

            for (int i = 0; i < WordDoc.BodyElements.Count; i++)
            {
                var element = WordDoc.BodyElements[i];
                if (element.ElementType == BodyElementType.PARAGRAPH)
                {
                    var paragraph = (XWPFParagraph)element;
                    if (paragraph.Style == "Heading1" && paragraph.Text == "Heading A")
                    {
                        //var run = paragraph.CreateRun();
                        //run.SetText(aBuilder.ToString());

                        var para = (XWPFParagraph)WordDoc.BodyElements[i + 1];
                        var run = para.CreateRun();
                        run.SetText(aBuilder.ToString());
                    }
                    else if (paragraph.Style == "Heading2" && paragraph.Text == "Heading D")
                    {
                        //var run = paragraph.CreateRun();
                        //var tempPara = WordDoc.CreateParagraph();
                        //var tempRun = tempPara.CreateRun();
                        //tempRun.SetText(dBuilder.ToString());
                        //paragraph.AddRun(tempRun);
                        //run.SetText(dBuilder.ToString());

                        var para = (XWPFParagraph)WordDoc.BodyElements[i + 1];
                        var run = para.CreateRun();
                        run.SetText(dBuilder.ToString());
                    }
                }
            }

            //foreach(var kvp in RunPositions)
            //{
            //    //var run = ((XWPFParagraph)WordDoc.BodyElements[kvp.Key]).InsertNewRun(kvp.Key + 1);                
            //    var para = WordDoc.CreateParagraph();
            //    var run = para.CreateRun();
            //    run.SetText(kvp.Value);
            //    WordDoc.SetParagraph(para, kvp.Key + 1);
            //    //WordDoc.RemoveBodyElement(WordDoc.GetPosOfParagraph(para));
            //}

            //foreach (var element in WordDoc.BodyElements)
            //{
            //    if (element.ElementType == BodyElementType.PARAGRAPH)
            //    {
            //        var paragraph = (XWPFParagraph)element;
            //        if (paragraph.Style == "Heading1" && paragraph.Text == "Heading A")
            //        {
            //            //var run = paragraph.CreateRun();
            //            //run.SetText(aBuilder.ToString());
            //        }
            //        else if (paragraph.Style == "Heading2" && paragraph.Text == "Heading D")
            //        {
            //            //var run = paragraph.CreateRun();
            //            //var tempPara = WordDoc.CreateParagraph();
            //            //var tempRun = tempPara.CreateRun();
            //            //tempRun.SetText(dBuilder.ToString());
            //            //paragraph.AddRun(tempRun);
            //            //run.SetText(dBuilder.ToString());
            //        }
            //    }
            //}

            FileStream sw = new FileStream(Path.Combine(Path.GetDirectoryName(OutputPathTextBox.Text), "sampleOutput.docx"), FileMode.Create);
            WordDoc.Write(sw);
            sw.Close();

            ExcelDoc.Close();
            WordDoc.Close();
        }

        private string OpenFileBrowser(string defaultExt, string filter)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = defaultExt;
            dlg.Filter = filter;

            // Display OpenFileDialog by calling ShowDialog method 
            var result = dlg.ShowDialog();

            return result ?? false ? dlg.FileName : string.Empty;
        }
    }
}
