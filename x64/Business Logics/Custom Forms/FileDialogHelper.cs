using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using WinForms = System.Windows.Forms;
using System.Windows.Forms;

namespace Approved_BPCreation
{
    public class FileDialogHelper
{
        private static Thread _dialogThread;
        private static AutoResetEvent _threadReady = new AutoResetEvent(false);
        private static SynchronizationContext _syncContext;
        private static bool _initialized = false;
        private static readonly object _lock = new object();

        /// <summary>
        /// Initialize the STA thread. Call this once during add-on startup.
        /// </summary>
        public static void Initialize()
        {
            lock (_lock)
            {
                if (_initialized) return;

                _dialogThread = new Thread(() =>
                {
                    _syncContext = new WinForms.WindowsFormsSynchronizationContext();
                    _threadReady.Set();
                    WinForms.Application.Run(); // keep message loop alive
                });

                _dialogThread.SetApartmentState(ApartmentState.STA);
                _dialogThread.IsBackground = true;
                _dialogThread.Start();

                _threadReady.WaitOne();
                _initialized = true;
            }
        }

        /// <summary>
        /// Show a file dialog on top of SAP Business One
        /// </summary>
        public static string ShowFileDialog(string filter = "All files (*.*)|*.*", string title = "Select a file")
        {
            if (!_initialized)
                throw new InvalidOperationException("FileDialogHelper not initialized. Call FileDialogHelper.Initialize() during add-on startup.");

            string result = null;
            var done = new AutoResetEvent(false);

            _syncContext.Post(_ =>
            {
                try
                {
                    using (var dummy = new WinForms.Form())
                    {
                        dummy.TopMost = true;
                        dummy.StartPosition = WinForms.FormStartPosition.CenterScreen;
                        dummy.ShowInTaskbar = false;

                        using (var dialog = new WinForms.OpenFileDialog())
                        {
                            dialog.Filter = filter;
                            dialog.Title = title;
                            dialog.RestoreDirectory = true;

                            if (dialog.ShowDialog(dummy) == WinForms.DialogResult.OK)
                                result = dialog.FileName;
                        }
                    }
                }
                finally
                {
                    done.Set();
                }
            }, null);

            // Wait without freezing SAP
            while (!done.WaitOne(50))
                WinForms.Application.DoEvents();

            return result;
        }
       

    }
 }
