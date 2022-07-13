using IWshRuntimeLibrary;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.IO;

namespace Multiple_Tools_for_Firefox
{
    internal class Program
    {
        static string dataDir = AppDomain.CurrentDomain.BaseDirectory + @"data\";
        static string cookiesJsonDir = AppDomain.CurrentDomain.BaseDirectory + @"data\cookiesJson\";
        static string cookieShortcutDir = AppDomain.CurrentDomain.BaseDirectory + @"data\cookieShortcut\";
        static string webExtracted = "https://www.facebook.com";
        static void Main(string[] args)
        {
            //CookieExtractor();
            //CookieOpener(args);
            //CookieShortcut();
            //ProfileShortcut();
        }

        static void CookieExtractor()
        {
            FirefoxDriverService service = FirefoxDriverService.CreateDefaultService(dataDir);
            service.HideCommandPromptWindow = true;
            var profileFirefoxDirectories = Directory.GetDirectories(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Mozilla\Firefox\Profiles\");

            Console.WriteLine("Input profile name to get cookie (or leave blank): ");
            string action = Console.ReadLine();
            bool act = false;
            if (!string.IsNullOrEmpty(action))
            {
                profileFirefoxDirectories[0] = "p." + action;
                act = true;
            }
            foreach (var profileFirefox in profileFirefoxDirectories)
            {
                var profileFirefoxName = profileFirefox.Split('.')[1];
                Console.WriteLine("extracting profile: " + profileFirefoxName);
                FirefoxOptions options = new FirefoxOptions();
                options.Profile = new FirefoxProfileManager().GetProfile(profileFirefoxName);
                IWebDriver driver = new FirefoxDriver(service, options);
                driver.Url = webExtracted;
                System.IO.File.WriteAllText(cookiesJsonDir + @"\" + profileFirefoxName + ".json", JsonConvert.SerializeObject(driver.Manage().Cookies.AllCookies));
                driver.Quit();
                if (act)
                {
                    break;
                }

            };

            Console.WriteLine("Done");
            Console.ReadLine();
        }
        static void CookieOpener(string[] args)
        {
            if (args.Length == 0)
            {
                return;
            }
            var command = args[0];
            FirefoxDriverService service = FirefoxDriverService.CreateDefaultService(dataDir);
            service.HideCommandPromptWindow = true;
            IWebDriver driver = new FirefoxDriver(service);
            driver.Url = webExtracted;
            dynamic cookies;

            using (StreamReader r = new StreamReader(cookiesJsonDir + command))
            {
                cookies = JsonConvert.DeserializeObject(r.ReadToEnd());
            }

            foreach (var cookie in cookies)
            {
                var cookieTmpDictionary = new Dictionary<string, object>();
                cookieTmpDictionary.Add("name", cookie["Name"]);
                cookieTmpDictionary.Add("value", cookie["Value"]);
                cookieTmpDictionary.Add("domain", cookie["Domain"]);
                cookieTmpDictionary.Add("path", cookie["Path"]);
                cookieTmpDictionary.Add("expiry", cookie["Expiry"]);
                var cookieTmp = Cookie.FromDictionary(cookieTmpDictionary);
                driver.Manage().Cookies.AddCookie(cookieTmp);
            }
            driver.Url = webExtracted;
        }

        static void CookieShortcut()
        {
            foreach (var file in Directory.GetFiles(cookiesJsonDir, "*.json"))
            {
                string filename = cookieShortcutDir + Path.GetFileName(file) + ".bat";
                using (StreamWriter writer = System.IO.File.CreateText(filename))
                {
                    writer.WriteLine("taskkill /F /IM geckodriver.exe /T");
                    writer.WriteLine(AppDomain.CurrentDomain.BaseDirectory + "CookieOpener.exe " + '"' + Path.GetFileName(file) + '"');
                }
            }
        }

        static void ProfileShortcut()
        {
            object shDesktop = (object)"Desktop";
            WshShell shell = new WshShell();
            var profileFirefoxDirectories = Directory.GetDirectories(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Mozilla\Firefox\Profiles\");
            foreach (var profileFirefox in profileFirefoxDirectories)
            {
                string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + @"\" + profileFirefox.Split('.')[1] + ".lnk";
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
                shortcut.TargetPath = @"C:\Program Files\Mozilla Firefox\firefox.exe";
                shortcut.Arguments = "-p " + '"' + profileFirefox.Split('.')[1] + '"';
                shortcut.Save();
            };
        }
    }
}
