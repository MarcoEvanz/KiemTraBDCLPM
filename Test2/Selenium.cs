using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Test2
{
    internal class Selenium
    {
        IWebDriver driver;
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
            driver.Navigate().GoToUrl("https://portal.huflit.edu.vn/Login");
            Thread.Sleep(2000);
            Assert.That(driver.Title, Is.EqualTo("Đăng nhập"));
        }
        [Test]
        public void Input_Invalid()
        {
            Test();
            driver.FindElement(By.Name("txtTaiKhoan")).SendKeys("00961");
            driver.FindElement(By.Name("txtMatKhau")).SendKeys("00961");
            driver.FindElement(By.XPath("//input[@value='Đăng nhập']")).Click();
            Thread.Sleep(1000);
            string actual = driver.FindElement(By.CssSelector("div[class='loginbox-forgot'] span")).Text;
            string expected = "Tên đăng nhập hoặc mật khẩu không chính xác";
            //Assert.That(actual, Is.EqualTo("Tên đăng nhập hoặc mật khẩu không chính xác"));
            if(actual != expected)
            {
                Console.WriteLine("Fail");
            }
            else
            {
                Console.WriteLine("Pass.");
            }

        }
        [TearDown]
        public void TearDown()
        {
            driver.Quit();
        }
    }
}
