// Developer Express Code Central Example:
// How to implement progress indicator for slow RichEditControl operations
// 
// This example demonstrates how to utilize the IProgressIndicationService service
// interface of the RichEditControl. The ProgressBar control on the form indicates
// how the mail merge operation proceeds.
// 
// You can find sample updates and versions for different programming languages here:
// http://www.devexpress.com/example=E3075

using System;
using System.Collections.Generic;
using System.Text;
using DevExpress.XtraEditors;
using DevExpress.Services;
using System.Windows.Forms;
using DevExpress.Xpf.Editors;
using System.Windows.Threading;

namespace ProgressIndicator
{
    #region #myprogressindicator
    class MyProgressIndicatorService : IProgressIndicationService
    {
        private ProgressBarEdit _Indicator;
        public ProgressBarEdit Indicator
        {
            get { return _Indicator; }
            set { _Indicator = value; }
        }

        public MyProgressIndicatorService(IServiceProvider provider, ProgressBarEdit indicator)
        {
            _Indicator = indicator;
        }
        
        #region IProgressIndicationService Members

        void IProgressIndicationService.Begin(string displayName, int minProgress, int maxProgress, int currentProgress)
        {
            _Indicator.Minimum = minProgress;
            _Indicator.Maximum = maxProgress;
            _Indicator.EditValue = currentProgress;
            _Indicator.Visibility = System.Windows.Visibility.Visible;
            Refresh();
        }

        void IProgressIndicationService.End()
        {
            _Indicator.Visibility = System.Windows.Visibility.Collapsed;
            Refresh();
        }

        void IProgressIndicationService.SetProgress(int currentProgress)
        {
            _Indicator.EditValue = currentProgress;
            Refresh();
        }
        #endregion

        void Refresh() {
            Action emptyDelegate = delegate() { };
            _Indicator.Dispatcher.Invoke(DispatcherPriority.Render, emptyDelegate);
        }
    }
#endregion #myprogressindicatorsindicator
}
