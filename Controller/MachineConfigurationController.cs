using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static System.Console;
using static System.Convert;
using BootCamperMachineConfiguration.Helper;
using BootCamperMachineConfiguration.Model;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Configuration;

namespace BootCamperMachineConfiguration.Controller
{
    public class MachineConfigurationController
    {
        #region Private Variables

        private readonly IConfiguration _configuration =
            new ConfigurationBuilder().AddJsonFile("appSettings.json", true, true).Build();

        private static string _icEmail = string.Empty;

        #endregion

        #region Constructor

        public MachineConfigurationController(string icEmail) => _icEmail = icEmail;

        #endregion

        #region Update Google Sheet with Current Machine Details

        /// <summary>
        /// Updating Boot Camper Machine Details in Google Drive
        /// </summary>
        public void UpdateBootCamperMachineDetails()
        {
            var currentSystemProps = new CurrentSystemConfigurationController().SystemProperties;
            var validate = ValidateSystemProperties(currentSystemProps);
            var icList = GetIcsFromGoogleDrive();
            var index = (from ic in icList.Values
                where ic.Any(e => e.ToString() == _icEmail)
                select icList.Values.IndexOf(ic)).FirstOrDefault();
            // Why index + 3?
            // We started with retrieving data from third row. So +2
            // Retrieved index value starts from 0 so if it is 1 then actual value is 1+1
            // Finally it 2 + index + 1
            index = index + 3;
            //Update Enough Specifications
            var enoughSpecsStoredCell = _configuration["EnoughSpecsStoredCell"];
            var spreadSheetId = _configuration["SpreadSheetID"];
            GoogleDriveHelper.UpdateValuesInGoogleDriveSpreadSheet(spreadSheetId, enoughSpecsStoredCell + index,
                validate ? "Yes" : "No");
            //Update Actual Machine Specs
            var actualMachineSpecs = new StringBuilder();
            actualMachineSpecs
                .AppendLine(" - CPU : " + currentSystemProps.ProcessorName)
                .AppendLine(" - CPU Score : " + currentSystemProps.CpuBenchMarkScore)
                .AppendLine(" - Memory : " + currentSystemProps.UsableMemory + "G")
                .AppendLine(" - OS : " + currentSystemProps.OperatingSystemEdition)
                .AppendLine(" - Storage : " + currentSystemProps.StorageSpace + "G")
                .AppendLine(" - Free disk space : " + currentSystemProps.FreeDiskSpace + "G")
                .Append(" - Architecture : " + currentSystemProps.Architecture);
            var actualMachineSpecsStoredCell = _configuration["ActualMachineSpecsStoredCell"];
            GoogleDriveHelper.UpdateValuesInGoogleDriveSpreadSheet(spreadSheetId, actualMachineSpecsStoredCell + index,
                actualMachineSpecs.ToString());
            WriteLine("Machine Configuration Details Updated Successfully in Google Drive in ICs sheets");
        }

        /// <summary>
        /// Validate current system properties wit expected system properties
        /// </summary>
        /// <param name="currentSystemProps"></param>
        /// <returns></returns>
        private bool ValidateSystemProperties(SystemProperties currentSystemProps)
        {
            var expectedAllSystemProps = GetExpectedSystemProperties();
            if (expectedAllSystemProps == null)
            {
                WriteLine("Entered IC/BootCamper Email not found in our list...");
                WriteLine("Press any key to continue...");
                ReadKey(true);
                throw new Exception("Entered IC/BootCamper Details not found...");
            }
            foreach (var expectedSystemProps in expectedAllSystemProps)
            {
                var expectedOperatingSystemList = expectedSystemProps.OperatingSystem.Split("&");
                var expectedOperatingSystem = string.Empty;
                foreach (var operatingSystem in expectedOperatingSystemList)
                {
                    if (!operatingSystem.Contains(currentSystemProps.OperatingSystemName)) continue;
                    expectedOperatingSystem = operatingSystem;
                    break;
                }
                var expectedOsVersion = Regex.Replace(expectedOperatingSystem, @"[^0-9]+", string.Empty);
                var currentOperatingSystem =
                    Regex.Replace(currentSystemProps.OperatingSystemEdition, @"[^0-9-.]+", string.Empty);
                var currentOsVersion = currentOperatingSystem.Split(".")[0];
                if (ToInt32(currentOsVersion) < ToInt32(expectedOsVersion))
                    return false;
                if (currentSystemProps.CpuBenchMarkScore < expectedSystemProps.MinimalCpuBenchmarkScore)
                    return false;
                if (currentSystemProps.UsableMemory < expectedSystemProps.Memory)
                    return false;
                if (currentSystemProps.StorageSpace < expectedSystemProps.Storage)
                    return false;
                if (currentSystemProps.FreeDiskSpace < expectedSystemProps.FreeDiskSpace)
                    return false;
                if (!currentSystemProps.Architecture.Contains(expectedSystemProps.Architecture))
                    return false;
            }
            return true;
        }

        #endregion

        #region Get Expected System Properties

        /// <summary>
        /// Get expected system properties for the entered IC/BootCamper 
        /// </summary>
        /// <returns></returns>
        private IEnumerable<OutputProperties> GetExpectedSystemProperties()
        {
            var enteredIcProjectDetails = GetICsProjectList().Where(e => e.IcEmail == _icEmail).ToList();
            if (enteredIcProjectDetails.Count <= 0) return null;
            {
                var spreadSheetId = _configuration["SpreadSheetID"];
                var range = _configuration["ExpectedSystemPropertiesSheet"];
                var response = GoogleDriveHelper.GetValuesFromGoogleDriveSpreadSheet(spreadSheetId, range,
                    SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.COLUMNS);
                return (from value in response.Values
                    let projectName = value[0].ToString()
                    where enteredIcProjectDetails.Exists(e => e.Projects.Exists(d => d == projectName))
                    select new OutputProperties
                    {
                        OperatingSystem = value[1]?.ToString(),
                        MinimalCpuBenchmarkScore = ToInt32(value[2]),
                        Memory = ToDouble(value[3]),
                        Storage = ToInt64(value[4]),
                        FreeDiskSpace = ToInt64(value[5]),
                        Architecture = value[6]?.ToString()
                    }).ToList();
            }
        }

        /// <summary>
        /// Prepare ICs and ICs Projects list
        /// </summary>
        /// <returns></returns>
        private IEnumerable<IcProperties> GetICsProjectList()
        {
            var response = GetIcsFromGoogleDrive();
            return response.Values.Where(value => value.Count != 0)
                .Select(value => new IcProperties
                {
                    IcEmail = value[1]?.ToString(),
                    Projects = new List<string>
                        {
                            value[2]?.ToString(),
                            value[3]?.ToString(),
                            value[4]?.ToString(),
                            value[5]?.ToString()
                        }.Distinct()
                        .ToList()
                }).ToList();
        }

        #endregion

        #region Get ICs List

        /// <summary>
        /// Get ICs and ICs Projects list from Google Drive
        /// </summary>
        /// <returns></returns>
        private ValueRange GetIcsFromGoogleDrive()
        {
            var range = _configuration["IcsSheet"];
            var spreadSheetId = _configuration["SpreadSheetID"];
            var response = GoogleDriveHelper.GetValuesFromGoogleDriveSpreadSheet(spreadSheetId, range);
            return response;
        }

        #endregion
    }
}