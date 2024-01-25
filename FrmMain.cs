using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using OpenQA.Selenium.Support.UI;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Menu;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.InteropServices;

//
//NuGet\Install-Package Selenium.Support -Version 4.16.2
namespace NIRFQuality
{
    using Excel = Microsoft.Office.Interop.Excel;

    public partial class FrmMain : Form
    {
        ChromeDriver driver;
        Excel.Application excelApp;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        int startRow;
        int endRow;

        public FrmMain()
        {
            InitializeComponent();            
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            driver = new ChromeDriver();
            startRow = 711;
            endRow = 720;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            btnLogin.Enabled = false;
            driver.Url = "https://login.nirfindia.org/";
            //driver.FindElement(By.Name("q")).SendKeys("webdriver");
            //Console.WriteLine(driver.Title);
            excelApp = new Excel.Application();
            workbook = excelApp.Workbooks.Open(@"E:\projects\NIRFQuality\FDP_26_12_23.xlsx");
            worksheet = workbook.Worksheets["FDP"];
            btnStart.Enabled = true;            
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            //  Allow main UI thread to properly display please wait form.
            System.Windows.Forms.Application.DoEvents();
            driver.Url = "https://login.nirfindia.org/Innovation/PII/FDP";
            Console.WriteLine(driver.Title);
            Range usedRange = worksheet.UsedRange;
            driver.FindElement(By.XPath("//*[@id='tbodyNoofconstitutedcollege']/tr[1]/td[7]/a")).Click();
            for (int row = startRow; row <= endRow; row++) //usedRange.Rows.Count
            {
                Console.WriteLine("Writing Row " + row);                
                SelectElement yearDropDown = new SelectElement(driver.FindElement(By.XPath("//*[@id='tbodyNoofconstitutedcollege']/tr[" +row + "]//td[1]/select")));
                yearDropDown.SelectByText(worksheet.Cells[row, 2].Value2.ToString());
                driver.FindElement(By.XPath("//*[@id='tbodyNoofconstitutedcollege']/tr[" +row + "]/td[2]/input")).SendKeys(worksheet.Cells[row, 3].Value2.ToString());
                driver.FindElement(By.XPath("//*[@id='tbodyNoofconstitutedcollege']/tr[" +row + "]/td[3]/input")).SendKeys(worksheet.Cells[row, 4].Value2.ToString());
                driver.FindElement(By.XPath("//*[@id='tbodyNoofconstitutedcollege']/tr[" +row + "]/td[4]/input")).SendKeys(worksheet.Cells[row, 5].Value2.ToString());
                
                String cellValue = worksheet.Cells[row, 6].Value2.ToString();
                double dateDouble = double.Parse(cellValue);
                DateTime dateValue = DateTime.FromOADate(dateDouble);
                String dateString = dateValue.ToString("dd-MM-yyyy");
                driver.FindElement(By.XPath("//*[@id='tbodyNoofconstitutedcollege']/tr[" +row + "]/td[5]/input")).SendKeys(dateString);
                cellValue = worksheet.Cells[row, 7].Value2.ToString();
                dateDouble = double.Parse(cellValue);
                dateValue = DateTime.FromOADate(dateDouble);
                dateString = dateValue.ToString("dd-MM-yyyy");
                driver.FindElement(By.XPath("//*[@id='tbodyNoofconstitutedcollege']/tr[" +row + "]/td[6]/input")).SendKeys(dateString);
                driver.FindElement(By.XPath("//*[@id='tbodyNoofconstitutedcollege']/tr[1]/td[7]/a")).Click();
                // Optional: Pause between rows for visual clarity
                //System.Threading.Thread.Sleep(100); // Adjust delay as needed
            }
            Console.WriteLine("Completed");
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            driver.Quit();
            workbook.Close();
            excelApp.Quit();

            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
        }        
    }
}
