using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Utilities;

namespace RPAChallange
{
    public class ChromeBrowser
    {
        private IWebDriver _driver;
        private List<List<Dictionary<string, string>>> _extractedData;
        private static readonly Dictionary<string, string> XPathDictionary = new Dictionary<string, string>
        {
            { "FIRSTNAME", "//input[@ng-reflect-name=\"labelFirstName\"]" },
            { "LASTNAME", "//input[@ng-reflect-name=\"labelLastName\"]" },
            { "COMPANYNAME", "//input[@ng-reflect-name=\"labelCompanyName\"]" },
            { "ROLEINCOMPANY", "//input[@ng-reflect-name=\"labelRole\"]" },
            { "ADDRESS", "//input[@ng-reflect-name=\"labelAddress\"]" },
            { "EMAIL", "//input[@ng-reflect-name=\"labelEmail\"]" },
            { "PHONENUMBER", "//input[@ng-reflect-name=\"labelPhone\"]" },
        };

        public ChromeBrowser()
        {
            _driver = InitiateDriver();
        }

        public async Task Run()
        {
            // Navigate
            _driver.Navigate().GoToUrl(@"https://rpachallenge.com/");
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            _driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
            Thread.Sleep(2000);

            // Download Excel
            string filePath = await DownloadExcel();

            // Extract downloaded Excel data
            ExcelHandler handler = new ExcelHandler();
            _extractedData = handler.ReadAndExtractData(filePath);

            // Press Start Button to begin
            ButtonOnClick(By.XPath("//button[contains(@class, 'btn-large') and contains(@class, 'uiColorButton')]"));

            // Populate then submit
            foreach (var worksheetRows in _extractedData)
            {
                foreach (var rowDetails in worksheetRows)
                {
                    // Populate
                    PopulateInputForms(rowDetails);

                    // Submit
                    ButtonOnClick(By.XPath("//input[@value=\"Submit\"]"));
                    Thread.Sleep(500);
                }
            }

            // Close
            Thread.Sleep(5000);
            _driver.Close();
        }

        private static ChromeDriver InitiateDriver()
        {
            var options = new ChromeOptions();
            //options.AddArguments("--headless"); // Uncomment to enable headless
            options.AddArgument("no-sandbox");
            options.AddArgument("--ignore-certificate-errors");
            options.AddArgument("--ignore-ssl-errors");
            options.AcceptInsecureCertificates = true;
            return new ChromeDriver(options);
        }


        private async Task<string> DownloadExcel()
        {
            var elements = _driver.FindElements(By.XPath("//a[contains(@class, 'btn') and contains(@class, 'waves-effect') and contains(@class, 'uiColorPrimary')]"));

            if (elements.Count == 0)
            {
                Console.WriteLine("[E]: Unable to find element.");
                Environment.Exit(1);
            }

            string downloadUrl = elements[0].GetAttribute("href");
            //string downloadFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Path.GetFileName(downloadUrl));
            string downloadFilePath = Path.Combine(ProjectSourcePath.Value, Path.GetFileName(downloadUrl));

            using (var client = new HttpClient())
            {
                using (var response = await client.GetAsync(new Uri(downloadUrl)))
                {
                    response.EnsureSuccessStatusCode();
                    using (var s = await response.Content.ReadAsStreamAsync())
                    using (var f = File.OpenWrite(downloadFilePath))
                    {
                        await s.CopyToAsync(f);
                    }
                }
            }

            if (!File.Exists(downloadFilePath))
            {
                Console.WriteLine("[E]: File download failed.");
                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine("[I]: File download successfully at \'" + downloadFilePath + "\'.");
            }

            return downloadFilePath;
        }


        private void PopulateInputForms(Dictionary<string, string> rowDetails)
        {
            foreach (var cell in rowDetails)
            {
                Console.WriteLine("[I]: Populating inputs... " + cell.Key + ": " + cell.Value);
                _driver.FindElement(By.XPath(XPathDictionary[cell.Key])).SendKeys(cell.Value);
            }
        }


        private void ButtonOnClick(By by)
        {
            Console.WriteLine("[I]: Clicked on button.");
            try
            {
                _driver.FindElement(by).Click();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                Environment.Exit(1);
            }
        }

    }
}