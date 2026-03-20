using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ASCOM.LunaticAstro.FilterWheel.FilterWheelDriver.Views
{
    public partial class SetupDialogControl : UserControl
    {
        public SetupDialogControl()
        {
            InitializeComponent();
            Loaded += SetupDialogControl_Loaded;
        }

        private void SetupDialogControl_Loaded(object sender, RoutedEventArgs e)
        {
            // Equivalent to SetupDialogForm_Load
            // Populate COM ports, load settings, etc.
        }

        private void CmdOK_Click(object sender, RoutedEventArgs e)
        {
            // Save settings here
            RaiseCloseRequest(true);
        }

        private void CmdCancel_Click(object sender, RoutedEventArgs e)
        {
            RaiseCloseRequest(false);
        }

        private void BrowseToAscom(object sender, MouseButtonEventArgs e)
        {
            System.Diagnostics.Process.Start("https://ascom-standards.org");
        }

        // Event to notify the WinForms host to close the dialog
        public event Action<bool> CloseRequested;

        private void RaiseCloseRequest(bool result)
        {
            CloseRequested?.Invoke(result);
        }
    }
}