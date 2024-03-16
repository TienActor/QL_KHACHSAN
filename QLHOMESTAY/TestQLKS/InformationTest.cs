﻿using ClosedXML.Excel;
using ExcelDataReader;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestQLKS
{
    public class InformationTest
    {

        [TestFixture]
        public class ContactLoginTest
        {
            private IWebDriver driver;
            private WebDriverWait wait;

            // private string initialUrl = "http://localhost:49921/Account/Login";


            [SetUp]
            public void SetUp()
            {
                // Register the code page provider to ensure encoding 1252 is available
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                driver = new ChromeDriver();
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10)); // Adjust the time as necessary.
                driver.Navigate().GoToUrl("http://localhost:49921/Home/Contact");
                Thread.Sleep(1000);


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
                // Cập nhật kết quả test trong file Excel
                var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheet(5); // Giả sử kết quả test nằm ở sheet đầu tiên

                bool isTestCaseFound = false;

                foreach (IXLRow row in worksheet.RowsUsed())
                {
                    // Giả sử cột 'A' chứa ID của test case
                    if (row.Cell("A").Value.ToString() == testCaseID)
                    {
                        isTestCaseFound = true;
                        // Cập nhật cột 'G' với kết quả, đảm bảo rằng đây là cột đúng trong file của bạn
                        row.Cell("F").SetValue(result);
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
            public void contactLogin()
            {

                // Đọc dữ liệu test từ file Excel
                var testData = ReadTestData("C:\\Users\\TIEN\\Documents\\DBCL\\DataTestWeb.xlsx");
                int testCaseIndex = 1;


                foreach (DataRow row in testData.Rows)
                {
                    // Define testCaseId at the beginning of the loop
                    string testCaseId = $"{testCaseIndex}";

                    // Lấy thông tin từ datatest
                    string username = row["UserName"].ToString();
                    string email = row["Email"].ToString();
                    string message = row["TestMessenger"].ToString();
                    string rating = row["Rating"].ToString() ;
                    string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();

                    int ratingValue = int.Parse(rating); // Giả định rating là số và hợp lệ
                                                         //string expectedErrorMessage = row["ExpectedErrorMessage"].ToString();
                    int starIndex = 6 - ratingValue;
                    try
                    {
                        wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ho_ten")));
                        driver.FindElement(By.Id("ho_ten")).Clear();
                        driver.FindElement(By.Id("ho_ten")).Click();
                        driver.FindElement(By.Id("ho_ten")).SendKeys(username);


                        driver.FindElement(By.Id("mail")).Clear();
                        driver.FindElement(By.Id("mail")).SendKeys(email);
                        Thread.Sleep(1000);

                        driver.FindElement(By.CssSelector(".btn-primary")).Click();
                        wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText("Phản Hồi"))).Click();

                        var noiDungInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("noi_dung")));
                        noiDungInput.Click();
                        noiDungInput.SendKeys(message);
                        Thread.Sleep(1000);
                        // Sử dụng nth-child trong CSS Selector có thể không đúng, nên kiểm tra lại hoặc cập nhật nếu cần

                        var RatingInput = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[3]/div/div/div/div[1]/form/div[2]/div")));
                        RatingInput.Click();

                        string cssSelector = $".star:nth-child({starIndex})";

                        noiDungInput.SendKeys(cssSelector);
                        // driver.FindElement(By.CssSelector(cssSelector)).Click();
                        //driver.FindElement(By.CssSelector(".star:nth-child({starIndex})")).Click();
                        SubmitFormAndHandleAlert(expectedErrorMessage, testCaseId);
                    }
                    catch (Exception e)
                    {
                        Assert.Fail("Test failed with exception: " + e.Message);
                    }
                }
            }


            private void SubmitFormAndHandleAlert(string expectedErrorMessage, string testCaseId)
            {
                // Click on the submit button and handle the alert
                driver.FindElement(By.CssSelector(".btn")).Click();
                IAlert alert = wait.Until(ExpectedConditions.AlertIsPresent());
                string alertText = alert.Text;
                alert.Accept();


                if (!string.IsNullOrEmpty(expectedErrorMessage))
                {
                    // If an error message is expected, wait for the alert and verify the message
                  

                    Assert.AreEqual(expectedErrorMessage, alertText, $"Test case {testCaseId} failed. Expected error message: {expectedErrorMessage}, but got: {alertText}");
                    UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Failed");
                }

                else
                {
                    // If no error message is expected, check the page URL or other success criteria
                    wait.Until(ExpectedConditions.UrlContains("http://localhost:49921/Home/Contact"));
                    UpdateTestResult("C:\\Users\\TIEN\\Documents\\DBCL\\TestCaseTien.xlsx", testCaseId, "Passed");
                }
            }
            [TearDown]
            public void Teardown()
            {
                driver.Quit();
                driver.Dispose();
            }
        }

    }
}
