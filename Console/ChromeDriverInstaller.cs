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
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace SeleniumConsole
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

            chromeVersion ??= await GetChromeVersion();

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

            var driverName = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "chromedriver.exe" : "chromedriver";

            var targetPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            targetPath = Path.Combine(targetPath, driverName);
            if (!forceDownload && File.Exists(targetPath))
            {
                using var process = Process.Start(
                    new ProcessStartInfo
                    {
                        FileName = targetPath,
                        ArgumentList = { "--version" },
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                    }
                );
                string existingChromeDriverVersion = await process.StandardOutput.ReadToEndAsync();
                string error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();
                process.Kill(true);

                // expected output is something like "ChromeDriver 88.0.4324.96 (68dba2d8a0b149a1d3afac56fa74648032bcf46b-refs/branch-heads/4324@{#1784})"
                // the following line will extract the version number and leave the rest
                existingChromeDriverVersion = existingChromeDriverVersion.Split(" ")[1];
                if (chromeDriverVersion == existingChromeDriverVersion)
                {
                    return;
                }

                if (!string.IsNullOrEmpty(error))
                {
                    throw new Exception($"Failed to execute {driverName} --version");
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
                using Stream chromeDriverStream = entry.Open();
                await chromeDriverStream.CopyToAsync(chromeDriverWriter);
            }

            // on Linux/macOS, you need to add the executable permission (+x) to allow the execution of the chromedriver
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                using var process = Process.Start(
                    new ProcessStartInfo
                    {
                        FileName = "chmod",
                        ArgumentList = { "+x", targetPath },
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                    }
                );
                string error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();
                process.Kill(true);

                if (!string.IsNullOrEmpty(error))
                {
                    throw new Exception("Failed to make chromedriver executable");
                }
            }
        }

        public async Task<string> GetChromeVersion()
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
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

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                try
                {
                    using var process = Process.Start(
                        new ProcessStartInfo
                        {
                            FileName = "google-chrome",
                            ArgumentList = { "--product-version" },
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                        }
                    );
                    string output = await process.StandardOutput.ReadToEndAsync();
                    string error = await process.StandardError.ReadToEndAsync();
                    await process.WaitForExitAsync();
                    process.Kill(true);

                    if (!string.IsNullOrEmpty(error))
                    {
                        throw new Exception(error);
                    }

                    return output;
                }
                catch (Exception ex)
                {
                    throw new Exception("An error occurred trying to execute 'google-chrome --product-version'", ex);
                }
            }

            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                try
                {
                    using var process = Process.Start(
                        new ProcessStartInfo
                        {
                            FileName = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                            ArgumentList = { "--version" },
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                        }
                    );
                    string output = await process.StandardOutput.ReadToEndAsync();
                    string error = await process.StandardError.ReadToEndAsync();
                    await process.WaitForExitAsync();
                    process.Kill(true);

                    if (!string.IsNullOrEmpty(error))
                    {
                        throw new Exception(error);
                    }

                    output = output.Replace("Google Chrome ", "");
                    return output;
                }
                catch (Exception ex)
                {
                    throw new Exception(
                        $"An error occurred trying to execute '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome --version'",
                        ex
                    );
                }
            }

            throw new PlatformNotSupportedException("Your operating system is not supported.");
        }

        private static string GetChromeDriverPlatformString(ChromeDriverPlatform platform)
            => ChromeDriverPlatformMap[platform];

        private static readonly Dictionary<ChromeDriverPlatform, string> ChromeDriverPlatformMap = new()
        {
            { ChromeDriverPlatform.Linux64, "linux64" },
            { ChromeDriverPlatform.MacArm64, "mac-arm64" },
            { ChromeDriverPlatform.MacX64, "mac-x64" },
            { ChromeDriverPlatform.Win32, "win32" },
            { ChromeDriverPlatform.Win64, "win64" },
        };
    }

    public enum ChromeDriverPlatform
    {
        Linux64,
        MacArm64,
        MacX64,
        Win32,
        Win64
    }
}