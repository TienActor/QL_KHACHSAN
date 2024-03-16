using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;


using OpenQA.Selenium.Chrome;

using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;


namespace TestQLKS
{
    [TestFixture]
    public class TestWithExcelData
    {
        public class Fixture
        {
            private IWebDriver driver;
            private WebDriverWait wait;
            private IXLWorkbook workbook;
            private IXLWorksheet worksheet;
            private string testDataFilePath = "C:\\Users\\TIEN\\Downloads\\Calculator_Testing.xlsx";

            // private string testResultFilePath = "C:\\path_to_your_test_result.xlsx";

            [SetUp]
            public void SetUp()
            {
                driver = new ChromeDriver();
                workbook = new XLWorkbook(testDataFilePath);
                worksheet = workbook.Worksheet(1);
                driver.Manage().Window.Maximize();
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            }

            [Test]
            public void TestCalculatorWithDataFromExcel()
            {
                foreach (var row in worksheet.RangeUsed().RowsUsed().Skip(1)) // Skip header row
                {
                    // Đọc dữ liệu test từ Excel
                    string id = row.Cell(1).GetValue<string>();
                    string description = row.Cell(2).GetValue<string>();
                    double number1 = row.Cell(3).GetValue<double>();
                    double number2 = row.Cell(4).GetValue<double>();
                    double expected = row.Cell(7).GetValue<double>();

                    // Chạy test với dữ liệu đó
                    driver.Navigate().GoToUrl("https://testsheepnz.github.io/BasicCalculator.html");
                    driver.FindElement(By.Id("number1Field")).Clear();
                    driver.FindElement(By.Id("number1Field")).SendKeys(number1.ToString());
                    driver.FindElement(By.Id("number2Field")).Clear();
                    driver.FindElement(By.Id("number2Field")).SendKeys(number2.ToString());
                    driver.FindElement(By.Id("calculateButton")).Click();
                    double actual = Convert.ToDouble(driver.FindElement(By.Id("numberAnswerField")).GetAttribute("value"));

                    // Wait for result to be displayed
                    Thread.Sleep(1000); // It's better to use WebDriverWait


                    // Lấy kết quả thực tế
                    var actualResult = Convert.ToDouble(driver.FindElement(By.Id("numberAnswerField")).GetAttribute("value"));

                    // So sánh kết quả thực tế với kết quả mong đ
                    // Check the result and write back to Excel
                    // Check the result and write back to Excel
                    bool passed = Math.Abs(actual - expected) < 0.0001; // Use a tolerance for floating point comparison
                    row.Cell(9).Value = actual; // Actual result
                    row.Cell(10).Value = passed ? "Passed" : "Failed"; // Test result

                    Assert.AreEqual(expected, actual, 0.0001, id + " failed");
                }
            }


            private void UpdateTestResult(string testCaseID, string result)
            {
                // Mở file Excel
                using (var workbook = new XLWorkbook(testDataFilePath))
                {
                    // Lấy sheet chứa kết quả test
                    var worksheet = workbook.Worksheet("TestSheetName"); // Thay đổi "TestSheetName" thành tên sheet của bạn

                    // Tìm hàng với testCaseID
                    var row = worksheet.RowsUsed().FirstOrDefault(r => r.Cell(1).Value.ToString() == testCaseID);
                    if (row != null)
                    {
                        // Cập nhật kết quả test
                        row.Cell("G").Value = result; // Thay đổi "G" thành số cột chứa kết quả Pass/Fail
                        workbook.Save();
                    }
                    else
                    {
                        throw new Exception($"Test case ID '{testCaseID}' not found.");
                    }
                }

               

            }
            [TearDown] // Correct position for TearDown
            public void Teardown()
            {
                driver.Quit();
               driver.Dispose();
                workbook.Save();
                workbook.Dispose();
                
            }
        }
    }
}
