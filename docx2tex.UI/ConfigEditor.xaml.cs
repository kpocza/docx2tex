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
using docx2tex.Library.Data;
using docx2tex.Library;
using System.ComponentModel;

namespace docx2tex.UI
{
    /// <summary>
    /// Interaction logic for ConfigEditor.xaml
    /// </summary>
    public partial class ConfigEditor : UserControl
    {
        #region Fields

        private Docx2TexConfig _conf;
        private ConfigurationClassEnum _configurationClass;
        private string _documentFilePath;
        private IContentClosable _contentClosable;

        private List<InputEncInfo> AllEncodings
        {
            get { return InputEnc.Instance.InputEncs; }
        }

        public IEnumerable<InputEncInfo> AllEncodingsPlusEmpty
        {
            get 
            {
                yield return new InputEncInfo(null, null, "Not defined");
                foreach (var enc in AllEncodings)
                {
                    yield return enc;
                }
            }
        }

        public IEnumerable<string> AllDocumentClassPlusEmpty
        {
            get
            {
                yield return string.Empty;
                yield return "article";
                yield return "book";
                yield return "report";
            }
        }

        public IEnumerable<string> AllFontSizePlusEmpty
        {
            get
            {
                yield return string.Empty;
                yield return "10pt";
                yield return "11pt";
                yield return "12pt";
            }
        }

        public IEnumerable<string> AllPaperSizePlusEmpty
        {
            get
            {
                yield return string.Empty;
                yield return "letterpaper";
                yield return "legalpaper";
                yield return "executivepaper";
                yield return "a4paper";
                yield return "a5paper";
                yield return "b5paper";
            }
        }

        #endregion

        #region Lifecycle methods

        public ConfigEditor(ConfigurationClassEnum configClass, string documentFilePath, IContentClosable contentClosable)
        {
            InitializeComponent();

            _configurationClass = configClass;
            _documentFilePath = documentFilePath;
            _contentClosable = contentClosable;
        }

        #endregion

        #region Event handlers

        private void ConfigEditor_Loaded(object sender, RoutedEventArgs e)
        {
            
            if (!DesignerProperties.GetIsInDesignMode(this))
            {
                Initialize();
            }
        }

        private void btnSelectImgMgck_Click(object sender, RoutedEventArgs e)
        {
            SelectImageMagick();
        }

        private void txtLineLength_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = new List<char>(e.Text.ToCharArray()).Exists(c => !char.IsDigit(c));
        }

        private void btnCleanPage_Click(object sender, RoutedEventArgs e)
        {
            if (ConfirmClick())
            {
                CleanPage(tcConfig.SelectedItem as TabItem);
            }
        }

        private void btnCleanAll_Click(object sender, RoutedEventArgs e)
        {
            if (ConfirmClick())
            {
                CleanAllPages();
            }
        }

        private void btnRevertPage_Click(object sender, RoutedEventArgs e)
        {
            if (ConfirmClick())
            {
                RevertPage(tcConfig.SelectedItem as TabItem);
            }
        }

        private void btnRevertAll_Click(object sender, RoutedEventArgs e)
        {
            if (ConfirmClick())
            {
                RevertAll();
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (ConfirmClickSave())
            {
                Config.SaveConfig(_conf);
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            if (ConfirmClick())
            {
                _contentClosable.ContentClose();
            }
        }

        private void btnStyleSelectWord2k7Doc_Click(object sender, RoutedEventArgs e)
        {
            StyleSelectWord2K7Doc();
        }

        #endregion

        #region Private operation methods

		private void Initialize()
        {
            _conf = LoadConfig();
            tcConfig.DataContext = _conf;
        }

        private void SelectImageMagick()
        {
            using (var ofd = new System.Windows.Forms.OpenFileDialog())
            {
                var fullName = txtSelectImgMgck.Text;
                string dir = string.Empty;
                string fileName = string.Empty;
                try
                {
                    dir = System.IO.Path.GetDirectoryName(fullName);
                    if (string.IsNullOrEmpty(dir))
                    {
                        dir = System.IO.Path.GetPathRoot(fullName);
                    }
                    fileName = System.IO.Path.GetFileName(fullName);
                }
                catch
                {
                    dir = string.Empty;
                    fileName = string.Empty;
                }

                ofd.InitialDirectory = dir;
                ofd.FileName = fileName;
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtSelectImgMgck.Text = ofd.FileName;
                }
            }
        }

        private void CleanPage(TabItem currentTabPage)
        {
            if (currentTabPage != null)
            {
                if (currentTabPage == tabInfra)
                {
                    _conf.Infra = new Infra();
                }
                else if (currentTabPage == tabLtXTags)
                {
                    _conf.LaTeXTags = new LaTeXTags();
                }
                else if (currentTabPage == tabStyleMap)
                {
                    _conf.StyleMap = new StyleMap();
                }
                ReBind();
            }
        }

        private void CleanAllPages()
        {
            _conf.CleanProperties();
            ReBind();
        }

        private void RevertPage(TabItem currentTabPage)
        {
            var conf = LoadConfig();
            if (currentTabPage != null)
            {
                if (currentTabPage == tabInfra)
                {
                    _conf.Infra = conf.Infra;
                }
                else if (currentTabPage == tabLtXTags)
                {
                    _conf.LaTeXTags = conf.LaTeXTags;
                }
                else if (currentTabPage == tabStyleMap)
                {
                    _conf.StyleMap = conf.StyleMap;
                }
                ReBind();
            }
        }

        private void RevertAll()
        {
            _conf = LoadConfig();
            ReBind();
        }

        private void ReBind()
        {
            tcConfig.DataContext = null;
            tcConfig.DataContext = _conf;
        }

        private Docx2TexConfig LoadConfig()
        {
            Docx2TexConfig config = null;
            switch (_configurationClass)
            {
                case ConfigurationClassEnum.System:
                    config = Config.LoadSystemConfig();
                    txtConfigLevel.Text = "System level configuration";
                    break;
                case ConfigurationClassEnum.User:
                    config = Config.LoadUserConfig();
                    txtConfigLevel.Text = "User level configuration";
                    break;
                case ConfigurationClassEnum.Document:
                    StaticConfigHelper.DocxPath = _documentFilePath;
                    config = Config.LoadDocumentConfig();
                    txtConfigLevel.Text = "Document level configuration";
                    txtStyleSelectWord2k7Doc.Text = _documentFilePath;
                    EnumerateSyles();

                    break;
            }
            txtConfigLevelInfo.Text = config.ConfigurationFilePath;
            return config;
        }

        private void txtLineLength_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrEmpty((e.Source as TextBox).Text))
            {
                _conf.Infra.LineLength = null;
            }
        }

        private bool ConfirmClick()
        {
            return MessageBox.Show("All changes will be lost. Are you sure?", "Confirmation", MessageBoxButton.YesNo) == MessageBoxResult.Yes;
        }

        private bool ConfirmClickSave()
        {
            return MessageBox.Show("All modifications will be persisted. Are you sure?", "Confirmation", MessageBoxButton.YesNo) == MessageBoxResult.Yes;
        }

        private void StyleSelectWord2K7Doc()
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
                    txtStyleSelectWord2k7Doc.Text = ofd.FileName;
                    EnumerateSyles();
                }
            }
        }

        private void EnumerateSyles()
        {
            var styles = StyleEnumerator.Enumerate(txtStyleSelectWord2k7Doc.Text);
            lbStyles.Items.Clear();
            styles.ForEach(s => lbStyles.Items.Add(s));
        }

	    #endregion    

        #region Styles D&D

		private void lbStyles_MouseDown(object sender, MouseButtonEventArgs e)
        {
            ListBox lb = e.Source as ListBox;

            string data = GetObjectDataFromPoint(lb, e.GetPosition(lb));

            if (data != null)
            {
                DragDrop.DoDragDrop(lb, data, DragDropEffects.Copy);
            }
        }

        private string GetObjectDataFromPoint(ListBox lb, Point point)
        {
            UIElement element = lb.InputHitTest(point) as UIElement;
            if (element != null)
            {
                object data = DependencyProperty.UnsetValue;
                while (data == DependencyProperty.UnsetValue)
                {
                    data = lb.ItemContainerGenerator.ItemFromContainer(element);

                    if (data == DependencyProperty.UnsetValue)
                    {
                        element = VisualTreeHelper.GetParent(element) as UIElement;
                    }

                    if (element == lb)
                        return null;
                }
                if (data != DependencyProperty.UnsetValue)
                {
                    return data as string;
                }
            }
            return null;
        }

	    #endregion    
    }

    public enum ConfigurationClassEnum
    {
        System,
        User,
        Document
    }
}
