using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using System;
using System.Threading;
using System.Linq;
using System.Diagnostics;

namespace NotepadTest
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class UnitTest1
    {

        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        private const string ExcelAppID = @"C:\Users\vgrk2\Desktop\Excel.lnk";

        protected static WindowsDriver<WindowsElement> session;

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
          

            if (session == null)
            {
                Process.Start(@"C:\Users\vgrk2\Desktop\Excel.lnk");
                               
                // Create a new session to launch Notepad application
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", ExcelAppID);
                

                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);
                Assert.IsNotNull(session.SessionId);

                // Verify that Notepad is started with untitled new file
                Assert.AreEqual("Excel", session.Title);

                // Set implicit timeout to 1.5 seconds to make element search to retry every 500 ms for at most three times
                session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);

                

            }

        }

        [TestMethod]
        public void TestMethod1()
        {

            Thread.Sleep(TimeSpan.FromSeconds(3));

            // LeftClick on ListItem "Blank workbook" at (156,96)
            Console.WriteLine("LeftClick on ListItem \"Blank workbook\" at (156,96)");
            // string xpath_LeftClickListItemBlankworkb_156_96 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"XLMAIN\"][@Name=\"Excel\"]/Pane[@ClassName=\"FullpageUIHost\"]/Pane[@ClassName=\"NetUIFullpageUIWindow\"]/Pane[@Name=\"Backstage view\"][@AutomationId=\"BackstageView\"]/Group[@ClassName=\"NetUIElement\"][@Name=\"Home\"]/Pane[@ClassName=\"NetUIScrollViewer\"][@Name=\"Home\"]/Group[@ClassName=\"NetUISlabContainer\"][@Name=\"New\"]/List[@ClassName=\"NetUIListView\"]/ListItem[@ClassName=\"NetUIListViewItem\"][@Name=\"Blank workbook\"]";
            var winElem_LeftClickListItemBlankworkb_156_96 = session.FindElementByName("Blank workbook");
            if (winElem_LeftClickListItemBlankworkb_156_96 != null)
            {
                winElem_LeftClickListItemBlankworkb_156_96.Click();
            }
            else
            {
                Console.WriteLine($"Failed to find element using xpath: ");
                return;
            }

            var FormulasRibbon = session.FindElementByName("Formulas");
            if (FormulasRibbon != null)
            {
                FormulasRibbon.Click();
                Thread.Sleep(TimeSpan.FromSeconds(3));
            }
            else
            {
                Console.WriteLine($"Failed to find element using xpath:");
                return;
            }



        }

        
        public void DataSelection() {

            // DO NOT USE - IN PROGRESS

            Thread.Sleep(TimeSpan.FromSeconds(3));

            // LeftClick on ListItem "Blank workbook" at (156,96)
            Console.WriteLine("LeftClick on ListItem \"Blank workbook\" at (156,96)");
            // string xpath_LeftClickListItemBlankworkb_156_96 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"XLMAIN\"][@Name=\"Excel\"]/Pane[@ClassName=\"FullpageUIHost\"]/Pane[@ClassName=\"NetUIFullpageUIWindow\"]/Pane[@Name=\"Backstage view\"][@AutomationId=\"BackstageView\"]/Group[@ClassName=\"NetUIElement\"][@Name=\"Home\"]/Pane[@ClassName=\"NetUIScrollViewer\"][@Name=\"Home\"]/Group[@ClassName=\"NetUISlabContainer\"][@Name=\"New\"]/List[@ClassName=\"NetUIListView\"]/ListItem[@ClassName=\"NetUIListViewItem\"][@Name=\"Blank workbook\"]";
            var winElem_LeftClickListItemBlankworkb_156_96 = session.FindElementByName("Blank workbook");
            if (winElem_LeftClickListItemBlankworkb_156_96 != null)
            {
                winElem_LeftClickListItemBlankworkb_156_96.Click();
            }
            else
            {
                Console.WriteLine($"Failed to find element using xpath: ");
                return;
            }

            // LeftClick on TabItem "Data" at (31,24)
            Console.WriteLine("LeftClick on TabItem \"Data\" at (31,24)");
            var DataTab = session.FindElementByName("Data");
            if (DataTab != null)
            {
                DataTab.Click();
            }
            else
            {
                Console.WriteLine($"Failed to find element using xpath:");
                return;
            }

            var DataToolsRibbon = session.FindElementByName("Remove Duplicates");
            if (DataToolsRibbon != null)
            {
                DataToolsRibbon.Click();
            }
            else
            {
                Console.WriteLine($"Failed to find element using xpath:");
                return;
            }



            /*// LeftClick on ComboBox "Allow:" at (172,16)
            Console.WriteLine("LeftClick on ComboBox \"Allow:\" at (172,16)");
            string xpath_LeftClickComboBoxAllow_172_16 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"bosa_sdm_XL9\"][@Name=\"Data Validation\"]/ComboBox[@Name=\"Allow:\"]";
            var winElem_LeftClickComboBoxAllow_172_16 = session.FindElementByName("Allow:");
            if (winElem_LeftClickComboBoxAllow_172_16 != null)
            {
                winElem_LeftClickComboBoxAllow_172_16.Click();
            }
            else
            {
                Console.WriteLine($"Failed to find element using xpath: {xpath_LeftClickComboBoxAllow_172_16}");
                return;
            }


            // LeftClick on Edit "Source:" at (350,14)
            Console.WriteLine("LeftClick on Edit \"Source:\" at (350,14)");
            string xpath_LeftClickEditSource_350_14 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"bosa_sdm_XL9\"][@Name=\"Data Validation\"]/Edit[@Name=\"Source:\"]";
            var winElem_LeftClickEditSource_350_14 = session.FindElementByName("Source:");
            if (winElem_LeftClickEditSource_350_14 != null)
            {
                winElem_LeftClickEditSource_350_14.Click();
            }
            else
            {
                Console.WriteLine($"Failed to find element using xpath: {xpath_LeftClickEditSource_350_14}");
                return;
            }


            // LeftClick on Button "OK" at (61,25)
            Console.WriteLine("LeftClick on Button \"OK\" at (61,25)");
            string xpath_LeftClickButtonOK_61_25 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"bosa_sdm_XL9\"][@Name=\"Data Validation\"]/Button[@Name=\"OK\"]";
            var winElem_LeftClickButtonOK_61_25 = session.FindElementByName("OK");
            if (winElem_LeftClickButtonOK_61_25 != null)
            {
                winElem_LeftClickButtonOK_61_25.Click();
            }
            else
            {
                Console.WriteLine($"Failed to find element using xpath: {xpath_LeftClickButtonOK_61_25}");
                return;
            }



*/


        }

        [ClassCleanup]
        public static void TearDown()
        {
            // Close the application and delete the session
            if (session != null)
            {
                session.Close();
                      
                session.Quit();
                session = null;
            }
        }
    } 
}
