using OpenQA.Selenium.Chrome;
using System;
using System.Threading.Tasks;

namespace SeleniumConsole
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            Console.WriteLine("Installing ChromeDriver");

            var chromeDriverInstaller = new ChromeDriverInstaller();

            // not necessary, but added for logging purposes
            var chromeVersion = await chromeDriverInstaller.GetChromeVersion();
            Console.WriteLine($"Chrome version {chromeVersion} detected");

            await chromeDriverInstaller.Install(chromeVersion);
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
