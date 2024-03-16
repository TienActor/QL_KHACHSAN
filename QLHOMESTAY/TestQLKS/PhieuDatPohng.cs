using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Data;
using ExcelDataReader;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TestQLKS
{
    internal class PhieuDatPohng
    {
        private IWebDriver driver;
        private WebDriverWait wait;

        public IDictionary<string, object> vars { get; private set; }
        private IJavaScriptExecutor js;
        [SetUp]
        public void SetUp()
        {

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            driver = new ChromeDriver();
            js = (IJavaScriptExecutor)driver;
            vars = new Dictionary<string, object>();
            driver.Navigate().GoToUrl("http://localhost:49921/");
            driver.FindElement(By.CssSelector("#page > nav > div > div > div.col-xs-8.text-right.menu-1 > ul > li:nth-child(8) > a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Đăng Nhập")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("ma_kh")).SendKeys("binh");
            Thread.Sleep(1000);
            driver.FindElement(By.Id("mat_khau")).SendKeys("123456");
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector(".btn-primary")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[2]/nav/div/div/div[2]/ul/li[8]/a")).Click();
            Thread.Sleep(1000);
        }
        [Test]
        public void CancelRent()
        {
            // Đọc dữ liệu test từ file Excel
            //var testData = ReadTestData("C:\\Users\\dowif\\Documents\\DBCLPM\\DataTest.xlsx");
            int testCaseIndex = 1;
           
                
                string testCaseId = $"{testCaseIndex}";

                driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/div[2]/div/div[2]/table/tbody/tr[1]/td[6]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div/form/div/input")).Click();
                Thread.Sleep(1000);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.UrlContains("http://localhost:49921/Home/BookRoom")); 
                Assert.That(driver.Url, Does.Contain("http://localhost:49921/Home/BookRoom"));

                // Cập nhật kết quả thành công vào file test cases
                UpdateTestResult("C:\\Users\\dowif\\Documents\\DBCLPM\\Testcase.xlsx", testCaseId, "Pass");
            
            }
        
    

        private DataTable ReadTestData(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    return result.Tables[0]; // Lấy sheet đầu tiên trong file Excel
                }
            }
        }

        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(7); // Assuming that the first worksheet is where you want to write the results

            bool isTestCaseFound = false;
            foreach (IXLRow row in worksheet.Rows())
            {
                if (row.Cell(1).Value.ToString().Equals(testCaseID))
                {
                    isTestCaseFound = true;
                    row.Cell("F").Value = result;
                    break;
                }
            }
            if (!isTestCaseFound)
            {
                throw new Exception($"Test case ID '{testCaseID}' not found in the Excel file.");
            }

            workbook.Save();
        }
        [TearDown]
        public void Teardown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}
