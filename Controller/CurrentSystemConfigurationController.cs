using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using static System.Convert;
using BootCamperMachineConfiguration.Helper;
using BootCamperMachineConfiguration.Model;
using Microsoft.Extensions.Configuration;

namespace BootCamperMachineConfiguration.Controller
{
    internal class CurrentSystemConfigurationController
    {
        #region Private Variables

        private readonly IConfiguration _configuration =
            new ConfigurationBuilder().AddJsonFile("appSettings.json", true, true).Build();
        private const string ObjectQuery = "SELECT * FROM Win32_OperatingSystem";
        private const string Win32Processor = "win32_processor";

        #endregion

        #region Get Current Machine Properties 

        /// <summary>
        /// Get Current Machine Properties
        /// </summary>
        /// <returns></returns>
        internal SystemProperties SystemProperties
        {
            get
            {
                var drivesList = DriveInfo.GetDrives().Where(e => e.IsReady).ToList();
                var systemProperties = new SystemProperties
                {
                    OperatingSystemName = OperatingSystemName,
                    OperatingSystemEdition = RuntimeInformation.OSDescription,
                    Architecture = RuntimeInformation.OSArchitecture.ToString(),
                    ProcessorName = ProcessorName,
                    UsableMemory = UsableMemory,
                    StorageSpace = GetGigaBytesFromBytes(drivesList.Sum(e => e.TotalSize)),
                    FreeDiskSpace = GetGigaBytesFromBytes(drivesList.Sum(e => e.AvailableFreeSpace))
                };
                var requiredProcessorName = GetProcessorName(ProcessorName);
                var currentCpuBenchMarkScore =
                    CpuBenchMarkScores.Where(e => e.CpuType.Contains(requiredProcessorName)).ToList();
                systemProperties.CpuBenchMarkScore = (currentCpuBenchMarkScore.Count > 0)
                    ? currentCpuBenchMarkScore[0].Score
                    : throw new Exception("IC/BootCamper CPU is not in our 'CPU Scores' list...");
                return systemProperties;
            }
        }

        /// <summary>
        /// Get the required processor name
        /// </summary>
        /// <param name="processorName"></param>
        /// <returns></returns>
        private static string GetProcessorName(string processorName)
        {
            processorName = processorName.Replace("Intel(R) Core(TM)", "");
            var startIndex = processorName.IndexOf("CPU", StringComparison.Ordinal);
            var endIndex = processorName.Length - startIndex - 3;
            processorName = processorName
                .Replace(processorName.Substring(startIndex, processorName.Length - endIndex), "").Trim();
            return processorName;
        }

        /// <summary>
        /// Get All CPUs bench mark scores
        /// </summary>
        /// <returns></returns>
        private IEnumerable<CpuBenchMarkScoreProperties> CpuBenchMarkScores
        {
            get
            {
                var range = _configuration["CpuScoresSheet"];
                var spreadSheetId = _configuration["SpreadSheetID"];
                var response = GoogleDriveHelper.GetValuesFromGoogleDriveSpreadSheet(spreadSheetId, range);
                return response.Values.Select(value => new CpuBenchMarkScoreProperties
                {
                    CpuType = value[0].ToString(),
                    Score = ToInt32(value[1])
                }).ToList();
            }
        }

        /// <summary>
        /// Get the operating system which is using in current machine
        /// </summary>
        /// <returns></returns>
        private static string OperatingSystemName
        {
            get
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return "Windows";
                if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return "MacOS";
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) return "Ubuntu";
                throw new Exception(
                    "Incompatible Operating System. Please use one of the Operating System from Windows,Mac or Linux");
            }
        }

        /// <summary>
        /// Get the current machine processor information
        /// </summary>
        /// <returns></returns>
        private static string ProcessorName
        {
            get
            {
                var managementClass = new ManagementClass(Win32Processor);
                var managementObjectCollection = managementClass.GetInstances();
                var processorInfo = string.Empty;
                foreach (var managementObject in managementObjectCollection)
                {
                    processorInfo = managementObject["Name"].ToString();
                }
                return processorInfo;
            }
        }

        /// <summary>
        /// Get the current machine usable memory information
        /// </summary>
        /// <returns></returns>
        private static double UsableMemory
        {
            get
            {
                var objectQuery = new ObjectQuery(ObjectQuery);
                var managementObjectCollection = new ManagementObjectSearcher(objectQuery).Get();
                double usableMemory = 0;
                foreach (var managementObject in managementObjectCollection)
                {
                    usableMemory = Math.Round((ToDouble(managementObject["TotalVisibleMemorySize"]) / (1024 * 1024)), 2);
                }
                return usableMemory;
            }
        }

        /// <summary>
        /// Converting bytes to Giga bytes
        /// </summary>
        /// <param name="memorySize"></param>
        /// <returns></returns>
        private static long GetGigaBytesFromBytes(long memorySize) => memorySize / 1024 / 1024 / 1024;

        #endregion
    }
}