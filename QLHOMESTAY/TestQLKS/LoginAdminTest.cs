using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using ExcelDataReader;
using System.Data;
using ClosedXML.Excel;
using SeleniumExtras.WaitHelpers;


namespace TestQLKS
{
    [TestFixture]
    public class LoginAdminTest
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        [SetUp]
        public void SetUp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Navigate().GoToUrl("http://localhost:49921/");
            driver.FindElement(By.XPath("/html/body/div[2]/nav/div/div/div[2]/ul/li[8]/a")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath("/html/body/div[2]/nav/div/div/div[2]/ul/li[8]/div/ul/li[1]/a")).Click();
        }
        private DataTable ReadTestData(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
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
                    return result.Tables[1];
                }
            }
        }
        private void UpdateTestResult(string filePath, string testCaseID, string result)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1); // Sửa lại số thứ tự worksheet nếu cần
            bool isTestCaseFound = false;
            foreach (IXLRow row in worksheet.RowsUsed())
            {
                if (row.Cell(1).Value.ToString() == testCaseID)
                {
                    isTestCaseFound = true;
                    row.Cell(6).SetValue(result);
                    break;
                }
            }
            if (!isTestCaseFound)
            {
                throw new Exception($"Test case ID '{testCaseID}' not found.");
            }
            workbook.Save();
        }
        [Test]
        public void Login()
        {
            var testData = ReadTestData("C:\\Users\\thong\\OneDrive\\Máy tính\\dataTest_Tho.xlsx");
            int testCaseIndex = 1;
            foreach (DataRow row in testData.Rows)
            {
                string testCaseId = $"LoginAdmin_{testCaseIndex}";
                string maAdmin = row["ma_admin"].ToString();
                string matKhau = row["mat_khau"].ToString();
                string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                string errorXPath = row["ErrorXPath"].ToString();
                try
                {
                    driver.FindElement(By.Id("ma_admin")).Click();
                    driver.FindElement(By.Id("ma_admin")).Clear();
                    driver.FindElement(By.Id("ma_admin")).SendKeys(maAdmin);
                    Thread.Sleep(100);
                    driver.FindElement(By.Id("mat_khau")).Click();
                    driver.FindElement(By.Id("mat_khau")).Clear();
                    driver.FindElement(By.Id("mat_khau")).SendKeys(matKhau);
                    Thread.Sleep(100);
                    driver.FindElement(By.CssSelector(".btn-primary")).Click();
                    Thread.Sleep(100);
                    if (!string.IsNullOrEmpty(expectedErrorMessage))
                    {
                        var errorElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(errorXPath)));
                        string actualErrorMessage = errorElement.Text;
                        Assert.That(actualErrorMessage, Is.EqualTo(expectedErrorMessage), $"Test case {testCaseId} failed. Expected error message: {expectedErrorMessage}, but got: {actualErrorMessage}");
                        UpdateTestResult("C:\\Users\\thong\\OneDrive\\Máy tính\\testCase_Tho.xlsx", testCaseId, actualErrorMessage == expectedErrorMessage ? "Pass" : "Failed");
                    }
                    else
                    {
                        // Trường hợp không có lỗi và chuyển trang dự kiến
                        wait.Until(ExpectedConditions.UrlContains("http://localhost:49921/"));
                        Assert.That(driver.Url, Does.Contain("http://localhost:49921/"), "Không quay lại trang chủ");
                        UpdateTestResult("C:\\Users\\thong\\OneDrive\\Máy tính\\testCase_Tho.xlsx", testCaseId, "Pass");
                    }
                }
                catch (Exception ex)
                {
                    UpdateTestResult("C:\\Users\\thong\\OneDrive\\Máy tính\\testCase_Tho.xlsx", testCaseId, "Fail");
                    Console.WriteLine($"Test failed for test case ID: {testCaseId} with error: {ex.Message}");
                }
                testCaseIndex++;
            }
        }
        [TearDown]
        protected void TearDown()
        {
            driver.Quit();
            driver.Dispose();
        }
    }
}