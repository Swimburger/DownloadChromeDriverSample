using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Reflection;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace SeleniumConsoleFramework
{
    public class ChromeDriverInstaller
    {
        private static readonly HttpClient HttpClient = new HttpClient();

        public Task Install(ChromeDriverPlatform platform) => Install(platform, null, false);

        public Task Install(ChromeDriverPlatform platform, string chromeVersion) =>
            Install(platform, chromeVersion, false);

        public Task Install(ChromeDriverPlatform platform, bool forceDownload) =>
            Install(platform, null, forceDownload);

        public async Task Install(ChromeDriverPlatform platform, string chromeVersion, bool forceDownload)
        {
            var platformString = GetChromeDriverPlatformString(platform);

            if (string.IsNullOrEmpty(chromeVersion))
            {
                chromeVersion = GetChromeVersion();
            }

            //   Take the Chrome version number, remove the last part, 
            chromeVersion = chromeVersion.Substring(0, chromeVersion.LastIndexOf('.'));
            //   and append the result to URL "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_". 
            //   For example, with Chrome version 72.0.3626.81, you'd get a URL "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_72.0.3626".
            var chromeDriverVersionResponse = await HttpClient.GetAsync(
                $"https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_{chromeVersion}"
            );
            if (!chromeDriverVersionResponse.IsSuccessStatusCode)
            {
                if (chromeDriverVersionResponse.StatusCode == HttpStatusCode.NotFound)
                {
                    throw new Exception($"ChromeDriver version not found for Chrome version {chromeVersion}");
                }

                throw new Exception(
                    $"ChromeDriver version request failed with status code: {chromeDriverVersionResponse.StatusCode}, " +
                    $"reason phrase: {chromeDriverVersionResponse.ReasonPhrase}"
                );
            }

            var chromeDriverVersion = await chromeDriverVersionResponse.Content.ReadAsStringAsync();

            const string driverName = "chromedriver.exe";

            var targetPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            targetPath = Path.Combine(targetPath, driverName);
            if (!forceDownload && File.Exists(targetPath))
            {
                using (var process = Process.Start(
                           new ProcessStartInfo
                           {
                               FileName = targetPath,
                               Arguments = "--version",
                               UseShellExecute = false,
                               CreateNoWindow = true,
                               RedirectStandardOutput = true,
                               RedirectStandardError = true,
                           }
                       ))
                {
                    string existingChromeDriverVersion = await process.StandardOutput.ReadToEndAsync();
                    string error = await process.StandardError.ReadToEndAsync();
                    process.WaitForExit();

                    // expected output is something like "ChromeDriver 88.0.4324.96 (68dba2d8a0b149a1d3afac56fa74648032bcf46b-refs/branch-heads/4324@{#1784})"
                    // the following line will extract the version number and leave the rest
                    existingChromeDriverVersion = existingChromeDriverVersion.Split(' ')[1];
                    if (chromeDriverVersion == existingChromeDriverVersion)
                    {
                        return;
                    }

                    if (!string.IsNullOrEmpty(error))
                    {
                        throw new Exception($"Failed to execute {driverName} --version");
                    }
                }
            }

            var downloadUrl = (await HttpClient.GetFromJsonAsync<JsonNode>(
                    "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"
                ))
                ["versions"]?.AsArray()
                // filter down to matching version
                .FirstOrDefault(n => n["version"]?.GetValue<string>() == chromeDriverVersion)
                // get downloads for chromedriver
                ?["downloads"]?["chromedriver"]?.AsArray()
                // filter downloads by platform
                .SingleOrDefault(n => n["platform"]?.GetValue<string>() == platformString)
                // get URL for platform download
                ?["url"]?.GetValue<string>();

            var driverZipResponse = await HttpClient.GetAsync(downloadUrl);
            if (!driverZipResponse.IsSuccessStatusCode)
            {
                throw new Exception(
                    $"ChromeDriver download request failed with status code: {driverZipResponse.StatusCode}, reason phrase: {driverZipResponse.ReasonPhrase}");
            }

            // this reads the zipfile as a stream, opens the archive, 
            // and extracts the chromedriver executable to the targetPath without saving any intermediate files to disk
            using (var zipFileStream = await driverZipResponse.Content.ReadAsStreamAsync())
            using (var zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Read))
            using (var chromeDriverWriter = new FileStream(targetPath, FileMode.Create))
            {
                var entry = zipArchive.Entries.SingleOrDefault(e => e.Name == driverName);
                using (Stream chromeDriverStream = entry.Open())
                {
                    await chromeDriverStream.CopyToAsync(chromeDriverWriter);
                }
            }
        }

        public string GetChromeVersion()
        {
            var chromePath = (string)Registry.GetValue(
                @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", null,
                null);
            if (chromePath == null)
            {
                throw new Exception("Google Chrome not found in registry");
            }

            var fileVersionInfo = FileVersionInfo.GetVersionInfo(chromePath);
            return fileVersionInfo.FileVersion;
        }

        private static string GetChromeDriverPlatformString(ChromeDriverPlatform platform)
            => ChromeDriverPlatformMap[platform];

        private static readonly Dictionary<ChromeDriverPlatform, string> ChromeDriverPlatformMap =
            new Dictionary<ChromeDriverPlatform, string>
            {
                { ChromeDriverPlatform.Win32, "win32" },
                { ChromeDriverPlatform.Win64, "win64" },
            };
    }

    public enum ChromeDriverPlatform
    {
        Win32,
        Win64
    }
}