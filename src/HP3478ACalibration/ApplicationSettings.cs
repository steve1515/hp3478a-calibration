using System;
using System.IO;
using System.IO.Ports;
using System.Diagnostics;
using System.Xml.Linq;

namespace HP3478ACalibration
{
    class ApplicationSettings
    {
        // Serial Port Settings
        // Note: Serial port settings below are default values,
        //       but they can be changed in the configuration file.
        public string SerialPortName { get; } = "COM1";
        public int SerialBaudRate { get; } = 9600;
        public int SerialDataBits { get; } = 8;
        public Parity SerialParity { get; } = Parity.None;
        public StopBits SerialStopBits { get; } = StopBits.One;
        public Handshake SerialFlowControl { get; } = Handshake.None;

        // GPIB Adapter Settings
        // Note: GPIB adapter settings below are default values,
        //       but they can be changed in the configuration file.
        public int GPIBAdapterTimeout { get; } = 1000;
        public string GPIBAdapterVersionString { get; } = "GPIB";


        public ApplicationSettings(string configFileName)
        {
            string applicationPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            string configFile = Path.Combine(applicationPath, configFileName);

            // Use defaults if configuration file is not readable
            if (!File.Exists(configFile))
                return;

            // Read configuration file settings
            XElement xmlConfig = XElement.Load(configFile);

            // Get serial port settings
            foreach (XElement serialPort in xmlConfig.Elements("SerialPort"))
            {
                SerialPortName = (string)serialPort.Element("PortName") ?? SerialPortName;
                SerialBaudRate = (int?)serialPort.Element("BaudRate") ?? SerialBaudRate;
                SerialDataBits = (int?)serialPort.Element("DataBits") ?? SerialDataBits;

                Enum.TryParse((string)serialPort.Element("Parity") ?? SerialParity.ToString(), true, out Parity parseParity);
                SerialParity = parseParity;

                Enum.TryParse((string)serialPort.Element("StopBits") ?? SerialStopBits.ToString(), true, out StopBits parseStopBits);
                SerialStopBits = parseStopBits;

                Enum.TryParse((string)serialPort.Element("FlowControl") ?? SerialFlowControl.ToString(), true, out Handshake parseFlowControl);
                SerialFlowControl = parseFlowControl;
            }

            // Get GPIB adapter settings
            foreach (XElement gpibAdapter in xmlConfig.Elements("GPIBAdapter"))
            {
                GPIBAdapterTimeout = (int?)gpibAdapter.Element("Timeout") ?? GPIBAdapterTimeout;
                GPIBAdapterVersionString = (string)gpibAdapter.Element("VersionString") ?? GPIBAdapterVersionString;
            }
        }
    }
}
