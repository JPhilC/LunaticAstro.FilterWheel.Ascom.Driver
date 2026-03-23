using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.ObjectModel;
using System.IO.Ports;

namespace ASCOM.LunaticAstro.FilterWheel.FilterWheelDriver.ViewModel
{
    internal partial class SetupViewModel : ObservableObject
    {
        public event Action<bool>? CloseRequested;

        [ObservableProperty]
        private ObservableCollection<string> comPorts = new();

        [ObservableProperty]
        private string? selectedComPort;

        [ObservableProperty]
        private bool traceEnabled;

        public SetupViewModel()
        {
            if (FilterWheelHardware.IsInDesignMode)
                return; // Skip all ASCOM logic in the designer

            LoadComPorts();
            LoadSettings();
        }

        private void LoadComPorts()
        {
            ComPorts.Clear();
            foreach (var port in SerialPort.GetPortNames())
                ComPorts.Add(port);
        }

        public void LoadSettings()
        {
            // Load from ASCOM profile into static fields
            FilterWheelHardware.ReadProfile();

            // Copy static fields into ViewModel
            SelectedComPort = FilterWheelHardware.ComPort;
            TraceEnabled = FilterWheelHardware.Tl.Enabled;

            // Fallback if no saved COM port
            if (string.IsNullOrWhiteSpace(SelectedComPort) && ComPorts.Count > 0)
                SelectedComPort = ComPorts[0];
        }

        [RelayCommand]
        private void Ok()
        {
            SaveSettings();
            CloseRequested?.Invoke(true);
        }

        [RelayCommand]
        private void Cancel()
        {
            CloseRequested?.Invoke(false);
        }

        private void SaveSettings()
        {
            // Copy ViewModel values back into static fields
            FilterWheelHardware.ComPort = SelectedComPort ?? string.Empty;
            FilterWheelHardware.Tl.Enabled = TraceEnabled;

            // Persist to ASCOM Profile
            FilterWheelHardware.WriteProfile();
        }
    }
}