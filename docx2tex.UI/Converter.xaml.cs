using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using docx2tex.Library;

namespace docx2tex.UI
{
    /// <summary>
    /// Interaction logic for Converter.xaml
    /// </summary>
    public partial class Converter : UserControl
    {

        #region Fields

        IRecentConversion _contentClosable;

        #endregion

        #region Lifecycle methods

        public Converter(IRecentConversion contentClosable)
        {
            _contentClosable = contentClosable;
            InitializeComponent();
        }

        #endregion

        #region Event handlers
        
        private void btnSelectWord2k7Doc_Click(object sender, RoutedEventArgs e)
        {
            SelectWord2K7Doc();
        }

        private void btnSelectLaTeXDoc_Click(object sender, RoutedEventArgs e)
        {
            SelectLaTeXDoc();
        }

        private void btnStartConversion_Click(object sender, RoutedEventArgs e)
        {
            StartConversion();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            _contentClosable.ContentClose();
        }

        #endregion

        #region Parameterization
        
        internal void SetParameters(string docxPath, string texPath)
        {
            txtSelectWord2k7Doc.Text = docxPath;
            txtSelectLaTeXDoc.Text = texPath;
        }

        #endregion

        #region Private operation methods
        
        private void SelectWord2K7Doc()
        {
            using (var ofd = new System.Windows.Forms.OpenFileDialog())
            {
                ofd.Multiselect = false;
                ofd.CheckFileExists = true;
                ofd.CheckPathExists = true;
                ofd.DereferenceLinks = true;
                ofd.Filter = "Word 2007 documents (*.docx;*.docm)|*.docx;*.docm";
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtSelectWord2k7Doc.Text = ofd.FileName;
                }
            }
        }

        private void SelectLaTeXDoc()
        {
            using (var ofd = new System.Windows.Forms.OpenFileDialog())
            {
                ofd.Multiselect = false;
                ofd.CheckPathExists = true;
                ofd.CheckFileExists = false;
                ofd.DereferenceLinks = true;
                ofd.Filter = "LaTeX documents (*.tex;*.ltx)|*.tex;*.ltx";
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtSelectLaTeXDoc.Text = ofd.FileName;
                }
            }
        }

        private void StartConversion()
        {
            var recentConversions = UserConfigHandler.GetRecentConversions();

            lblError.Content = string.Empty;

            string docxPath = txtSelectWord2k7Doc.Text;
            string texPath = txtSelectLaTeXDoc.Text;

            var statusInfo = new TextBoxOutput(txtScreen, scrLog);
            try
            {
                statusInfo.WriteLine(string.Empty);

                StaticConfigHelper.DocxPath = docxPath;

                var docx2TexWorker = new Docx2TexWorker();

                btnStartConversion.IsEnabled = false;
                statusInfo.WriteLine("Source: " + docxPath);
                statusInfo.WriteLine("Destination: " + texPath);
                statusInfo.WriteLine(string.Empty);

                if (docx2TexWorker.Process(docxPath, texPath, statusInfo))
                {
                    recentConversions.RemoveAll(rce => rce.From == docxPath && rce.To == texPath);
                    recentConversions.Insert(0, new FromToElement
                    {
                        Order = 0,
                        From = docxPath,
                        To = texPath
                    });
                    UserConfigHandler.UpdateRecentConversion(recentConversions);
                    _contentClosable.BuildRecentConversionMenus();
                }
            }
            catch (Exception ex)
            {
                lblError.Content = ex.Message.Replace(Environment.NewLine, string.Empty);
                statusInfo.Write(ex.ToString());
            }
            finally
            {
                statusInfo.WriteLine(string.Empty);
                btnStartConversion.IsEnabled = true;
            }
            statusInfo.Flush();
        }

        #endregion

        #region TextBoxOutput

		private class TextBoxOutput : IStatusInformation
        {
            private TextBox _txtScreen;
            private ScrollViewer _scrollViewer;
            private bool _lastCR;
            private DateTime _lastRefresh;

            public TextBoxOutput(TextBox txtScreen, ScrollViewer scrollViewer)
            {
                _txtScreen = txtScreen;
                _scrollViewer = scrollViewer;
                _lastCR = false;
                _lastRefresh = DateTime.Now.AddHours(-1.0);
            }

            public void Write(string data)
            {
                ResolveLastCR();
                _txtScreen.AppendText(data);
                EnsureText();
            }

            public void WriteCR(string data)
            {
                ResolveLastCR();
                _txtScreen.AppendText(data);
                _lastCR = true;
                EnsureText();
            }

            public void WriteLine(string data)
            {
                ResolveLastCR();
                _txtScreen.AppendText(data + Environment.NewLine);
                EnsureText();
            }

            void ResolveLastCR()
            {
                if (_lastCR)
                {
                    string data = _txtScreen.Text;
                    int lastNL = data.LastIndexOf(Environment.NewLine);
                    if (lastNL >= 0)
                    {
                        _txtScreen.Text = data.Substring(0, lastNL + 2);
                    }
//                    string []lines = _txtScreen.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
//                    _txtScreen.Text = string.Join(Environment.NewLine, lines, 0, lines.Length - 1) + Environment.NewLine;
                }
                _lastCR = false;
            }

            private void EnsureText()
            {
                DateTime lastRefresh = DateTime.Now;
                // for security reasons refresh screen only in every 200 milliseconds
                if (lastRefresh - _lastRefresh > new TimeSpan(0, 0, 0, 0, 200))
                {
                    _scrollViewer.ScrollToBottom();
                    System.Windows.Forms.Application.DoEvents();
                    _lastRefresh = lastRefresh;
                }
            }

            public void Flush()
            {
                _lastRefresh = _lastRefresh.AddHours(-1.0);
                EnsureText();
            }
        }

	    #endregion    
    }
}
