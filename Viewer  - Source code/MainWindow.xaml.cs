//   #define SAUTINCOM

using System;
using System.Collections.Generic;
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
using System.IO;
using System.Diagnostics;
using System.Reflection;


using SautinSoft.Document;
using System.Windows.Markup;
using Microsoft.Win32;
using System.Security.Cryptography;
using System.Globalization;

namespace Viewer
{
    public sealed class PathNode
    {
        public int Type = 0;
        public string FullPath = "";
    }

    [ValueConversion(typeof(bool?), typeof(Visibility))]
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || ((bool?)value) == false) return Visibility.Hidden;
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }

    public partial class MainWindow : Window
    {
        private object _dummyNode = null;
        private string _lastLoadedDocument = null;

       


        public MainWindow()
        {
            InitializeComponent();

            string orderLink = @"https://sautinsoft.com/products/document/order.php";

#if (SAUTINCOM)
            orderLink = @"https://sautin.com/products/components/document/order.php";
#endif

            OrderLink.NavigateUri = new Uri(orderLink);
            OrderLink.ToolTip = "See the licenses and prices for Document .Net.";
        }

        private void Viewer_Loaded(object sender, RoutedEventArgs e)
        {
            foreach (string s in Directory.GetLogicalDrives())
            {
                TreeViewItem item = new TreeViewItem();
                item.Header = s;
                item.Tag = new PathNode()
                {
                    Type = 0,
                    FullPath = s
                };
                item.FontWeight = FontWeights.Normal;
                item.Items.Add(_dummyNode);
                item.Expanded += new RoutedEventHandler(folder_Expanded);

                FoldersTree.Items.Add(item);
            }
        }

        private static string[] SupportedFileExt = 
        {
            "docx", "docm", "dotx", "dotm", "html", "htm", "pdf", "rtf"
            //"docx", "docm", "dotx", "dotm", "pdf", "rtf"
        };

        void folder_Expanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = (TreeViewItem)sender;
            if (item.Items.Count == 1 && item.Items[0] == _dummyNode)
            {
                ExpandFolder(item);
                e.Handled = true;
            }
        }

        private void ExpandFolder(TreeViewItem folderNode)
        {
            folderNode.Items.Clear();

            try
            {
                string root = (folderNode.Tag as PathNode).FullPath;

                foreach (string s in Directory.GetDirectories(root))
                {
                    TreeViewItem subitem = new TreeViewItem();
                    subitem.Header = s.Substring(s.LastIndexOf("\\") + 1);
                    subitem.Tag = new PathNode()
                    {
                        Type = 1,
                        FullPath = System.IO.Path.Combine(root, s)
                    };
                    subitem.FontWeight = FontWeights.Normal;
                    subitem.Items.Add(_dummyNode);
                    subitem.Expanded += new RoutedEventHandler(folder_Expanded);
                    folderNode.Items.Add(subitem);
                }

                List<string> files = new List<string>();
                foreach (string ext in SupportedFileExt)
                {
                    files.AddRange(Directory.GetFiles(root, "*." + ext, SearchOption.TopDirectoryOnly));
                }

                foreach (string s in files)
                {
                    TreeViewItem subitem = new TreeViewItem();
                    subitem.Header = s.Substring(s.LastIndexOf("\\") + 1);
                    subitem.Tag = new PathNode()
                    {
                        Type = 2,
                        FullPath = System.IO.Path.Combine(root, s)
                    };
                    subitem.FontWeight = FontWeights.Normal;
                    folderNode.Items.Add(subitem);
                }
            }
            catch 
            {
            }
        }

        private PageContent CreatePageContent(SautinSoft.Document.DocumentPage page)
        {
            PageContent pageContent = new PageContent();
            FixedPage fixedPage = new FixedPage();

            UIElement visual = page.GetContent();

            FixedPage.SetLeft(visual, 0);
            FixedPage.SetTop(visual, 0);

            double pageWidth = page.Width;
            double pageHeight = page.Height;

            fixedPage.Width = pageWidth;
            fixedPage.Height = pageHeight;

            fixedPage.Children.Add((UIElement)visual);

            //Size sz = new Size(pageWidth, pageHeight);
            //fixedPage.Measure(sz);
            //fixedPage.Arrange(new Rect(new Point(), sz));
            //fixedPage.UpdateLayout();

            ((IAddChild)pageContent).AddChild(fixedPage);
            return pageContent;
        }

        private void ViewDocument(string fullPath)
        {
            try
            {
                //Tests.Test1();
                //return;

                Stopwatch sw1 = new Stopwatch();
                Stopwatch sw2 = new Stopwatch();
                int pages = 0;

                System.Windows.Input.Cursor cursor = Mouse.OverrideCursor;
                try
                {
                    Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

                    sw1.Start();

                    DocumentCore dc;

                    if (fullPath.ToLower().EndsWith(".pdf"))
                    {
                        dc = DocumentCore.Load(fullPath, new PdfLoadOptions() { DetectTables = false });
                    }
                    else
                    {
                        dc = DocumentCore.Load(fullPath);
                    }                        

                    sw1.Stop();

                    FixedDocument fd = new FixedDocument();
                    fd.Cursor = System.Windows.Input.Cursors.Arrow;

                    sw2.Start();

                    foreach (SautinSoft.Document.DocumentPage page in dc.GetPaginator().Pages)
                    {
                        fd.Pages.Add(CreatePageContent(page));
                        pages++;
                    }

                    sw2.Stop();

                    documentViewer.Document = fd;

                    dc = null;
                }
                finally
                {
                    Mouse.OverrideCursor = cursor;
                    GC.Collect();
                }

                documentViewer.UpdateLayout();

                _lastLoadedDocument = fullPath;
                Viewer.Title = "Document .Net Viewer - " + fullPath;

                LoadedCaption.Content = @"Loaded (ms): " + sw1.ElapsedMilliseconds;
                RenderedCaption.Content = @"Rendered (ms): " + sw2.ElapsedMilliseconds;
                UpdateStatus();
            }
            catch (Exception ex)
            {
                _lastLoadedDocument = null;
                documentViewer.Document = null;
                documentViewer.UpdateLayout();
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Folders_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            TreeViewItem temp = (FoldersTree.SelectedItem as TreeViewItem);

            if (temp == null) return;

            PathNode pn = (temp.Tag as PathNode);
            if (pn.Type != 2) return;

            ViewDocument(pn.FullPath);
        }

        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            TreeViewItem temp = (FoldersTree.SelectedItem as TreeViewItem);

            if (temp == null) return;

            PathNode pn = (temp.Tag as PathNode);
            if (pn.Type != 2) return;

            ViewDocument(pn.FullPath);
        }

        private void FoldersTree_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TreeViewItem temp = (FoldersTree.SelectedItem as TreeViewItem);

                if (temp == null) return;

                PathNode pn = (temp.Tag as PathNode);
                if (pn.Type != 2) return;

                ViewDocument(pn.FullPath);

                e.Handled = true;
            }
        }

        private void ButtonOpenDocument_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Microsoft Word (*.docx)|*.docx|Rich Text Format(*.rtf)|*.rtf|Portable Document Format (*.pdf)|*.pdf|HyperText Markup Language (*.html)|*.html|All supported|*.docx;*.rtf;*.html;*.txt;*.pdf";
            openFileDialog.FilterIndex = 5;
            openFileDialog.ValidateNames = true;
            if (openFileDialog.ShowDialog() == true)
            {
                ViewDocument(openFileDialog.FileName);
            }
        }

        private void ButtonSaveAs_Click(object sender, RoutedEventArgs e)
        {
            if (_lastLoadedDocument == null) return;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = System.IO.Path.GetFileNameWithoutExtension(_lastLoadedDocument);
            saveFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(_lastLoadedDocument);
            saveFileDialog.Filter = "Microsoft Word (DOCX)|*.docx|Rich Text Format (RTF)|*.rtf|Portable Document Format (PDF)|*.pdf|HyperText Markup Language (HTML)|*.html|HyperText Markup Language (HTML) Fixed|*.html|Images (Png)|*.png";
            saveFileDialog.FilterIndex = 
                (System.IO.Path.GetExtension(_lastLoadedDocument).ToLower() == ".pdf") ? 1 : 3;
            saveFileDialog.AddExtension = true;
            saveFileDialog.OverwritePrompt = true;
            saveFileDialog.ValidateNames = true;

            if (saveFileDialog.ShowDialog() == true)
                try
                {
                    System.Windows.Input.Cursor cursor = Mouse.OverrideCursor;
                    string openFile = "";

                    try
                    {
                        Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

                        DocumentCore dc = DocumentCore.Load(_lastLoadedDocument);

                        switch (saveFileDialog.FilterIndex)
                        {
                            case 4:
                                openFile = saveFileDialog.FileName;
                                dc.Save(saveFileDialog.FileName, new HtmlFlowingSaveOptions());
                                break;

                            case 6:
                                {
                                    SautinSoft.Document.DocumentPaginator dp = dc.GetPaginator();
                                    string path = System.IO.Path.GetDirectoryName(saveFileDialog.FileName);
                                    string fileName = System.IO.Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                                    for (int idx = 0; idx < dp.Pages.Count; ++idx)
                                    {
                                        string s = System.IO.Path.Combine(path, fileName + (idx + 1).ToString() + ".png");
                                        if (string.IsNullOrEmpty(openFile)) openFile = s;
                                        dp.Pages[idx].Rasterize(200).Save(s);
                                    }
                                }
                                break;

                            default:
                                openFile = saveFileDialog.FileName;
                                dc.Save(saveFileDialog.FileName);
                                break;
                        }

                        dc = null;
                    }
                    finally
                    {
                        Mouse.OverrideCursor = cursor;
                        GC.Collect();
                    }

                    System.Diagnostics.Process.Start(openFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            TreeViewItem temp = (FoldersTree.SelectedItem as TreeViewItem);

            if (temp == null) return;

            PathNode pn = (temp.Tag as PathNode);
            if (pn.Type != 2) return;

            System.Diagnostics.Process.Start(pn.FullPath);
        }

        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            TreeViewItem temp = (FoldersTree.SelectedItem as TreeViewItem);
            if (temp == null || (temp.Tag as PathNode).Type != 2) return;

            string fullPath = (temp.Tag as PathNode).FullPath;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = System.IO.Path.GetFileNameWithoutExtension(fullPath);
            saveFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(fullPath);
            saveFileDialog.Filter = "Microsoft Word (DOCX)|*.docx|Rich Text Format (RTF)|*.rtf|Portable Document Format (PDF)|*.pdf|HyperText Markup Language (HTML)|*.html|HyperText Markup Language (HTML) Fixed|*.html|Images (Png)|*.png";
            saveFileDialog.FilterIndex = (System.IO.Path.GetExtension(fullPath).ToLower() == ".pdf") ? 1 : 3;
            saveFileDialog.AddExtension = true;
            saveFileDialog.OverwritePrompt = true;
            saveFileDialog.ValidateNames = true;

            if (saveFileDialog.ShowDialog() == true)
                try
                {
                    System.Windows.Input.Cursor cursor = Mouse.OverrideCursor;
                    string openFile = "";
                    try
                    {
                        Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

                        DocumentCore dc = DocumentCore.Load(fullPath);

                        switch (saveFileDialog.FilterIndex)
                        {
                            case 4:
                                openFile = saveFileDialog.FileName;
                                dc.Save(saveFileDialog.FileName, new HtmlFlowingSaveOptions());
                                break;

                            case 6:
                                {
                                    SautinSoft.Document.DocumentPaginator dp = dc.GetPaginator();
                                    string path = System.IO.Path.GetDirectoryName(saveFileDialog.FileName);
                                    string fileName = System.IO.Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                                    for (int idx = 0; idx < dp.Pages.Count; ++idx)
                                    {
                                        string s = System.IO.Path.Combine(path, fileName + (idx + 1).ToString() + ".png");
                                        if (string.IsNullOrEmpty(openFile)) openFile = s;
                                        dp.Pages[idx].Rasterize(200).Save(s);
                                    }
                                }
                                break;

                            default:
                                openFile = saveFileDialog.FileName;
                                dc.Save(saveFileDialog.FileName);
                                break;
                        }

                        dc = null;
                    }
                    finally
                    {
                        Mouse.OverrideCursor = cursor;
                        GC.Collect();
                    }

                    System.Diagnostics.Process.Start(openFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
        }

        private void UpdateStatus()
        {
            if (documentViewer.PageCount == 0)
            {
                PagesCaption.Content = "";
            }
            else
            {
                PagesCaption.Content = documentViewer.MasterPageNumber.ToString() +
                    " of " + documentViewer.PageCount.ToString();
            }

            ZoomCaption.Content = ((int)documentViewer.Zoom).ToString() + @"%";
        }

        private void ButtonZoomIn_Click(object sender, RoutedEventArgs e)
        {
            documentViewer.IncreaseZoom();
            UpdateStatus();
        }

        private void ButtonZoomOut_Click(object sender, RoutedEventArgs e)
        {
            documentViewer.DecreaseZoom();
            UpdateStatus();
        }

        private void ButtonZoom100_Click(object sender, RoutedEventArgs e)
        {
            documentViewer.Zoom = 100;
            UpdateStatus();
        }

        private void ButtonFitToWidth_Click(object sender, RoutedEventArgs e)
        {
            documentViewer.FitToWidth();
            UpdateStatus();
        }

        private void ButtonFitToHeight_Click(object sender, RoutedEventArgs e)
        {
            documentViewer.FitToHeight();
            UpdateStatus();
        }

        private void ButtonTwoPages_Click(object sender, RoutedEventArgs e)
        {
            documentViewer.FitToMaxPagesAcross(2);
            UpdateStatus();
        }

        private void ButtonPreviousPage_Click(object sender, RoutedEventArgs e)
        {
            documentViewer.PreviousPage();
            UpdateStatus();
        }

        private void ButtonNextPage_Click(object sender, RoutedEventArgs e)
        {
            documentViewer.NextPage();
            UpdateStatus();
        }

        private void FoldersTree_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ColumnDefinition cf = MainGrid.ColumnDefinitions[0];

            if (FoldersTree.Visibility != System.Windows.Visibility.Visible)
            {
                cf.Tag = cf.Width;
                cf.Width = new GridLength(0);
            }
            else
            {
                cf.Width = cf.Tag != null ? (GridLength)cf.Tag : new GridLength(300);
            }
        }

        private void FolderSplitter_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ColumnDefinition cf = MainGrid.ColumnDefinitions[1];

            if (FolderSplitter.Visibility != System.Windows.Visibility.Visible)
            {
                cf.Tag = cf.Width;
                cf.Width = new GridLength(0);
            }
            else
            {
                cf.Width = cf.Tag != null ? (GridLength)cf.Tag : new GridLength(4);
            }
        }

        private void documentViewer_PageViewsChanged(object sender, EventArgs e)
        {
            UpdateStatus();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            TreeViewItem temp = (FoldersTree.SelectedItem as TreeViewItem);
            if (temp == null) return;

            if ((temp.Tag as PathNode).Type == 2) temp = temp.Parent as TreeViewItem;

            ExpandFolder(temp);
        }
    }
}
