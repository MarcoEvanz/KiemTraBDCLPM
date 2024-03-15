using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using ExcelDataReader;
using OfficeOpenXml;
using NUnit.Framework.Interfaces;

namespace Test2
{
    internal class Demo
    {
        IWebDriver driver;
        public static IEnumerable<TestCaseData> GetTestCaseDatasFromExcel(string sheetName)
        {
            var testData = new List<TestCaseData>();
            using (var stream = File.Open("TestCaseData.xlsx", FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var table = result.Tables[sheetName];
                    for (int i= 0; i < table.Rows.Count; i++)
                    {
                        string username = Convert.ToString(table.Rows[i][0]);
                        string password = Convert.ToString(table.Rows[i][1]);
                        string expected = Convert.ToString(table.Rows[i][2]);
                        testData.Add(new TestCaseData(username, password, expected));
                    }
                }
            }
            return testData;
        }
        [TearDown]
        public void TearDown()
        {

            driver.Quit();
        }
        public void WriteDataToExcel(String Actual, string Result, String SheetName)
        {
            // Đường dẫn của tệp Excel đích
            string excelFilePath = "TestCaseData.xlsx";

            // Tạo một tệp Excel mới
            using (var excelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                // Lấy hoặc tạo một Sheet có tên được truyền vào
                var worksheet = excelPackage.Workbook.Worksheets[SheetName];

                // Ghi dữ liệu vào các ô trong Sheet
                int lastRow = 1;
                while (worksheet.Cells[lastRow, 4].Value != null)
                {
                    lastRow++;
                }

                // Ghi dữ liệu vào ô ở dòng mới sau dòng cuối cùng
                worksheet.Cells[lastRow, 4].Value = Actual;

                lastRow = 1;
                while (worksheet.Cells[lastRow, 5].Value != null)
                {
                    lastRow++;
                }

                // Ghi dữ liệu vào ô ở dòng mới sau dòng cuối cùng
                worksheet.Cells[lastRow, 5].Value = Result;

                // Lưu tệp Excel
                excelPackage.Save();
            }

            Console.WriteLine("Đã lưu dữ liệu thành công");
        }
        [SetUp]
        public void Setup()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            ChromeDriverService service = ChromeDriverService.CreateDefaultService("D:\\chromedriver-win64");
            driver = new ChromeDriver(service, options);


        }
        [Test]
        public void Test()
        {
            driver.Navigate().GoToUrl("https://www.saucedemo.com/");
            Thread.Sleep(1000);
        }

        [Test]
        [TestCaseSource(nameof(GetTestCaseDatasFromExcel), new object[] { "login" })]
        public void TestLogin(string username, string password, string expected)
        {
            string sheetname = "login";
            string result = "Pass";
            Test();
            driver.FindElement(By.CssSelector("*[data-test='username']")).SendKeys(username);
            driver.FindElement(By.CssSelector("*[data-test='password']")).SendKeys(password);
            driver.FindElement(By.CssSelector("*[data-test='login-button']")).Click();
            Thread.Sleep(1000);
            string actual = driver.Url;
            if (actual.Equals(expected))
            {
                Console.WriteLine("Actual: " + actual);
                Console.WriteLine( "Expected " +expected);
                result = "Pass";
                dangXuat();
            }
            else
            {
                Console.WriteLine("Actual: " + actual);
                Console.WriteLine("Expected " + expected);
                result = "False";
            }
            WriteDataToExcel(actual,result, sheetname);
        }
        public void loginAsStrd()
        {
            Test();
            driver.FindElement(By.CssSelector("*[data-test=\"username\"]")).Click();
            driver.FindElement(By.CssSelector("*[data-test=\"username\"]")).SendKeys("standard_user");
            driver.FindElement(By.CssSelector("*[data-test=\"password\"]")).Click();
            driver.FindElement(By.CssSelector("*[data-test=\"password\"]")).SendKeys("secret_sauce");
            driver.FindElement(By.CssSelector("*[data-test=\"login-button\"]")).Click();
            Thread.Sleep(1000);
        }
        [Test]
        public void xemThongTinSanPham()
        {
            loginAsStrd();
            driver.FindElement(By.CssSelector("#item_4_title_link > .inventory_item_name")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("*[data-test=\"back-to-products\"]")).Click();
            Thread.Sleep(1000);
            driver.Close();
        }
        [Test]
        public void themSanPhamVaXoaSanPham()
        {
            loginAsStrd();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("*[data-test=\"add-to-cart-sauce-labs-backpack\"]")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("1")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("*[data-test=\"remove-sauce-labs-backpack\"]")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("*[data-test=\"continue-shopping\"]")).Click();
            Thread.Sleep(1000);
            driver.Close();
        }
        [Test]
        public void testMuaHang()
        {
            loginAsStrd();
            driver.FindElement(By.CssSelector("*[data-test=\"add-to-cart-sauce-labs-backpack\"]")).Click();
            driver.FindElement(By.LinkText("1")).Click();
            driver.FindElement(By.CssSelector("*[data-test=\"checkout\"]")).Click();
            driver.FindElement(By.CssSelector("*[data-test=\"firstName\"]")).Click();
            driver.FindElement(By.CssSelector("*[data-test=\"firstName\"]")).SendKeys("Marco");
            driver.FindElement(By.CssSelector("*[data-test=\"lastName\"]")).Click();
            driver.FindElement(By.CssSelector("*[data-test=\"lastName\"]")).SendKeys("Chronos");
            driver.FindElement(By.CssSelector("*[data-test=\"postalCode\"]")).Click();
            driver.FindElement(By.CssSelector("*[data-test=\"postalCode\"]")).SendKeys("72000");
            driver.FindElement(By.CssSelector("*[data-test=\"continue\"]")).Click();
            driver.FindElement(By.CssSelector("*[data-test=\"finish\"]")).Click();
            driver.FindElement(By.CssSelector("*[data-test=\"back-to-products\"]")).Click();
            driver.Close();
        }
        [Test]
        public void dangXuat()
        {
            Thread.Sleep(1000);
            driver.FindElement(By.Id("react-burger-menu-btn")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("logout_sidebar_link")).Click();
            Thread.Sleep(1000);
            driver.Close();
        }
        [Test]
        public void reset()
        {
            loginAsStrd();
            Thread.Sleep(2000);
            driver.FindElement(By.Id("react-burger-menu-btn")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.Id("reset_sidebar_link")).Click();
            Thread.Sleep(2000);
            driver.Close();
        }
    }
}
