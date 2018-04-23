using System;
using System.Windows;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;

#region #usings
using DevExpress.Services;
#endregion #usings

namespace ProgressIndicator {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        public MainWindow() {
            InitializeComponent();
        }

        private void richEditControl1_Loaded(object sender, RoutedEventArgs e) {
            richEditControl1.ApplyTemplate();
            richEditControl1.LoadDocument("Docs\\invitation.docx");
            richEditControl1.Options.MailMerge.DataSource = new SampleData();
        }

        private void btnMailMerge_Click(object sender, RoutedEventArgs e) {
            MailMergeOptions myMergeOptions = richEditControl1.Document.CreateMailMergeOptions();
            myMergeOptions.MergeMode = MergeMode.NewSection;
            richEditControl1.Document.MailMerge(myMergeOptions, richEditControl2.Document);
            tabControl.SelectedIndex = 1;
        }

        private void richEditControl1_MailMergeStarted(object sender, MailMergeStartedEventArgs e) {
            #region #servicesubst
            richEditControl1.RemoveService(typeof(IProgressIndicationService));
            richEditControl1.AddService(typeof(IProgressIndicationService),
                new MyProgressIndicatorService(richEditControl1, this.progressBarControl1));
            #endregion #servicesubst
        }

        private void richEditControl1_MailMergeFinished(object sender, MailMergeFinishedEventArgs e) {
            richEditControl1.RemoveService(typeof(IProgressIndicationService));
        } 

        private void richEditControl1_MailMergeRecordStarted(object sender, MailMergeRecordStartedEventArgs e) {
            // Imitating slow data fetching
            System.Threading.Thread.Sleep(100);
        }

        private void richEditControl1_MailMergeRecordFinished(object sender, MailMergeRecordFinishedEventArgs e) {
            e.RecordDocument.AppendDocumentContent("Docs\\bungalow.docx", DocumentFormat.OpenXml);
        }

        private void tabControl_SelectionChanged(object sender, DevExpress.Xpf.Core.TabControlSelectionChangedEventArgs e) {
            switch (tabControl.SelectedIndex) {
                case 0:
                    this.btnMailMerge.IsEnabled = true;
                    break;
                case 1:
                    this.btnMailMerge.IsEnabled = false;
                    break;
            }
        }
    }
}
