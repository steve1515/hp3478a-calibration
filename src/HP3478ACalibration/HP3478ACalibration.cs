using System;
using System.Linq;
using System.Text;

namespace HP3478ACalibration
{
    public static class HP3478ACalibration
    {
        // The HP 3478A contains a 256x4 SRAM with the following format:
        //
        //   - The first nibble contains the CPU's calibration switch enable
        //     check value which alternates between 0x0 and 0xf when the switch
        //     is enabled. This first byte does not contain calibration data.
        //
        //   - The following nibbles contain 19 calibration entries each
        //     containing 13 nibbles for a total of 247 nibbles of calibration
        //     data.
        //
        //   - The remaining bytes are unused.
        //
        //   - Each 13 nibble calibration entry consists of 6 nibbles for
        //     offset, 5 for gain, and 2 for checksum.
        //
        // When the meter's SRAM is read via GPIB, an ASCII byte is returned
        // by adding 0x40 to the nibble.
        //
        // This library works with nibbles contained in bytes where the high
        // order nibble is discarded.
        // i.e. Functions accept byte arrays of either 13 bytes for a single
        //      calibration entry or 19 x 13 bytes for all calibration data.


        private const int _entryCount = 19;    // Number of calibration entries
        private const int _entryByteLen = 13;  // Number of bytes in a calibration entry

        // SRAM Calibration Entries
        private static readonly string[] _calibrationEntries =
        {                       // Calibration Entry Index
            "30 mV DC",         // 0
            "300 mV DC",        // 1
            "3 V DC",           // 2
            "30 V DC",          // 3
            "300 V DC",         // 4
            "<not used>",       // 5
            "AC V",             // 6
            "30 Ohm 2W/4W",     // 7
            "300 Ohm 2W/4W",    // 8
            "3 kOhm 2W/4W",     // 9
            "30 kOhm 2W/4W",    // 10
            "300 kOhm 2W/4W",   // 11
            "3 MOhm 2W/4W",     // 12
            "30 MOhm 2W/4W",    // 13
            "300 mA DC",        // 14
            "3A DC",            // 15
            "<not used>",       // 16
            "300 mA/3A AC",     // 17
            "<not used>"        // 18
        };

        // Unused Calibration Entry Indexes
        private static readonly int[] _unusedEntries = { 5, 16, 18 };

        /// <summary>
        /// Verifies all entries in calibration data contain a valid checksum.
        /// </summary>
        /// <param name="calibrationData">Calibration data.</param>
        /// <param name="skipUnusedEntries">True = Skip verifying unused calibration entries.</param>
        /// <returns>True on success, False otherwise.</returns>
        public static bool ValidateData(byte[,] calibrationData, bool skipUnusedEntries = true)
        {
            // Check for correct data length
            if (calibrationData.GetLength(0) != _entryCount
                || calibrationData.GetLength(1) != _entryByteLen)
                return false;

            for (int i = 0; i < _entryCount; i++)
            {
                if (skipUnusedEntries && _unusedEntries.Contains(i))
                    continue;

                if (!ValidateEntry(calibrationData.GetRow(i).ToArray()))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Verifies calibration entry data contains a valid checksum.
        /// </summary>
        /// <param name="entryData">Calibration entry data.</param>
        /// <returns>True on success, False otherwise.</returns>
        public static bool ValidateEntry(byte[] entryData)
        {
            // Check for correct data length
            if (entryData.Length != _entryByteLen)
                return false;

            byte checkSum = 0x00;
            for (int i = 0; i < 11; i++)
                checkSum += (byte)(entryData[i] & 0x0f);

            checkSum += (byte)(entryData[11] << 4);
            checkSum += (byte)(entryData[12] & 0x0f);

            return (checkSum == 0xff);
        }

        /// <summary>
        /// Prints table of all calibration data.
        /// </summary>
        /// <param name="calibrationData">Calibration data.</param>
        public static void PrintData(byte[,] calibrationData)
        {
            // Check for correct data length
            if (calibrationData.GetLength(0) != _entryCount
                || calibrationData.GetLength(1) != _entryByteLen)
                throw new ArgumentException("Invalid calibration data length.");

            // Get length of longest entry label string and create format string with correct alignment value.
            // Example: If max label length is 15, then format string becomes "{0,15}".
            int labelMaxLength = _calibrationEntries.OrderByDescending(s => s.Length).First().Length;
            string labelFormatString = String.Format("{{0,{0}}}", labelMaxLength);

            // Display table header
            Console.WriteLine("    {0}  Raw              Raw", "Calibration" + new string(' ', labelMaxLength - 11));
            Console.WriteLine("#   {0}  Offset  Offset   Gain   Gain      Checksum", "Entry" + new string(' ', labelMaxLength - 5));
            Console.WriteLine("--  {0}  ------  -------  -----  --------  --------", new string('-', labelMaxLength));

            for (int i = 0; i < _entryCount; i++)
            {
                byte[] entryData = calibrationData.GetRow(i).ToArray();

                Console.Write($"{i + 1,2:D2}  ");
                Console.Write(labelFormatString + "  ", _calibrationEntries[i]);
                Console.Write($"{GetOffsetRaw(entryData),6}  ");
                Console.Write($"{GetOffset(entryData),7}  ");
                Console.Write($"{GetGainRaw(entryData),5}  ");
                Console.Write($"{GetGain(entryData),8:F6}  ");
                Console.Write($"{entryData[11] & 0x0f:X}{entryData[12] & 0x0f:X}");
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Gets calibration offset value for a calibration entry.
        /// </summary>
        /// <param name="entryData">Calibration entry data.</param>
        /// <returns>Calibration offset value.</returns>
        public static int GetOffset(byte[] entryData)
        {
            // Check for correct data length
            if (entryData.Length != _entryByteLen)
                throw new ArgumentException("Invalid calibration entry data length.");

            // Convert BCD offset to integer
            int offsetValue = 0;
            for (int i = 0, j = 100000; i < 6; i++, j /= 10)
                offsetValue += (entryData[i] & 0x0f) * j;

            // Assume offset is negative if most significant BCD is >= 9
            // Example: 999999 = -1, 999998 = -2, etc.
            if (offsetValue >= 900000)
                offsetValue = offsetValue - 1000000;

            return offsetValue;
        }

        /// <summary>
        /// Gets calibration offset raw data string for a calibration entry.
        /// </summary>
        /// <param name="entryData">Calibration entry data.</param>
        /// <returns>Calibration offset raw data string.</returns>
        public static string GetOffsetRaw(byte[] entryData)
        {
            // Check for correct data length
            if (entryData.Length != _entryByteLen)
                throw new ArgumentException("Invalid calibration entry data length.");

            StringBuilder sb = new StringBuilder(6);
            for (int i = 0; i < 6; i++)
                sb.Append($"{entryData[i] & 0x0f:X}");

            return sb.ToString();
        }

        /// <summary>
        /// Gets calibration gain value for a calibration entry.
        /// </summary>
        /// <param name="entryData">Calibration entry data.</param>
        /// <returns>Calibration gain value.</returns>
        public static float GetGain(byte[] entryData)
        {
            // Check for correct data length
            if (entryData.Length != _entryByteLen)
                throw new ArgumentException("Invalid calibration entry data length.");

            // Convert gain data to gain value
            float gainValue = 1.0f;
            for (int i = 6, j = 100; i < 11; i++, j *= 10)
            {
                int gainNibble = entryData[i] & 0x0f;
                if (gainNibble >= 8)
                    gainNibble -= 16;

                gainValue += (float)gainNibble / j;
            }

            return gainValue;
        }

        /// <summary>
        /// Gets calibration gain raw data string for a calibration entry.
        /// </summary>
        /// <param name="entryData">Calibration entry data.</param>
        /// <returns>Calibration gain raw data string.</returns>
        public static string GetGainRaw(byte[] entryData)
        {
            // Check for correct data length
            if (entryData.Length != _entryByteLen)
                throw new ArgumentException("Invalid calibration entry data length.");

            StringBuilder sb = new StringBuilder(6);
            for (int i = 6; i < 11; i++)
                sb.Append($"{entryData[i] & 0x0f:X}");

            return sb.ToString();
        }
    }
}
