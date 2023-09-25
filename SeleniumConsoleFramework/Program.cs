using System;
using OpenQA.Selenium.Chrome;

namespace SeleniumConsoleFramework
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Installing ChromeDriver");

            var chromeDriverInstaller = new ChromeDriverInstaller();

            // not necessary, but added for logging purposes
            var chromeVersion = chromeDriverInstaller.GetChromeVersion();
            Console.WriteLine($"Chrome version {chromeVersion} detected");

            chromeDriverInstaller.Install(ChromeDriverPlatform.Win64, chromeVersion).Wait();
            Console.WriteLine("ChromeDriver installed");

            Console.WriteLine("Enter URL to visit:");
            var url = Console.ReadLine();
            if (string.IsNullOrEmpty(url))
            {
                Console.WriteLine("No URL entered");
                Console.WriteLine("Press any key to exit");
                Console.ReadKey();
                return;
            }

            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");
            using (var chromeDriver = new ChromeDriver(chromeOptions))
            {
                chromeDriver.Navigate().GoToUrl(url);
                Console.WriteLine($"Page title: {chromeDriver.Title}");
            }
            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }
    }
}