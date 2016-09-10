using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using ScrapySharp.Network;
using ScrapySharp.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;

namespace Crawler
{
    public partial class Form1 : Form
    {
        string connectionString = string.Empty;
        //string connectionString = "Server=tcp:zjding.database.windows.net,1433;Initial Catalog=Costco;Persist Security Info=False;User ID=zjding;Password=G4indigo;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";
        //string connectionString = "Data Source=DESKTOP-ABEPKAT;Initial Catalog=Costco;Persist Security Info=True;User ID=sa;Password=G4indigo";
        ScrapingBrowser Browser = new ScrapingBrowser();
        IWebDriver driver;

        List<string> categoryArray = new List<string>();
        List<string> subCategoryArray = new List<string>();
        List<string> productUrlArray = new List<string>();

        List<string> categoryUrlArray = new List<string>();
        List<string> subCategoryUrlArray = new List<string>();
        List<string> productListPages = new List<string>();

        List<string> newProductArray = new List<string>();
        List<string> discontinueddProductArray = new List<string>();
        List<string> priceUpProductArray = new List<string>();
        List<string> priceDownProductArray = new List<string>();
        List<string> stockChangeProductArray = new List<string>();

        List<string> eBayListingDiscontinueddProductArray = new List<string>();
        List<string> eBayListingPriceUpProductArray = new List<string>();
        List<string> eBayListingPriceDownProductArray = new List<string>();
        List<string> eBayListingOptionsChangeProductArray = new List<string>();


        int nScanProducts = 0;
        int nImportProducts = 0;
        int nSkipProducts = 0;
        int nImportErrors = 0;

        string emailMessage;

        string destinFileName;

        DateTime startDT;
        DateTime productListEndDT;
        DateTime endDT;

        List<string> firstTry = new List<string>();
        List<string> secondTry = new List<string>();
        List<string> firstTryResult = new List<string>();
        List<string> secondTryResult = new List<string>();

        int nProductListPages;
        int nProductUrlArray;
        int nCategoryUrlArray;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetConnectionString();

            runCrawl();

            //int nEBayListingChangePriceUp = 0;
            //int nEBayListingChangePriceDown = 0;
            //int nEBayListingChangeDiscontinue = 0;
            //int nEBayListingChangeOptions = 0;
            //CheckEBayListing(out nEBayListingChangePriceUp, out nEBayListingChangePriceDown, out nEBayListingChangeDiscontinue, out nEBayListingChangeOptions);

            this.Close();
        }

        public void SetConnectionString()
        {
            string azureConnectionString = "Server=tcp:zjding.database.windows.net,1433;Initial Catalog=Costco;Persist Security Info=False;User ID=zjding;Password=G4indigo;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";

            SqlConnection cn = new SqlConnection(azureConnectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            cn.Open();
            string sqlString = "SELECT ConnectionString FROM DatabaseToUse WHERE bUse = 1 and ApplicationName = 'Crawler'";
            cmd.CommandText = sqlString;
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    connectionString = (reader.GetString(0)).ToString();
                }
            }
            reader.Close();
            cn.Close();
        }

        public void CheckEBayListing(out int nEBayListingChangePriceUp, out int nEBayListingChangePriceDown, out int nEBayListingChangeDiscontinue, out int nEBayListingChangeOptions)
        {
            if (string.IsNullOrEmpty(connectionString))
                SetConnectionString();

            GetEBayListingProudctUrls();

            GetProductInfo(true, false);

            PopulateTables();

            CompareEBayListings(out nEBayListingChangePriceUp, out nEBayListingChangePriceDown, out nEBayListingChangeDiscontinue, out nEBayListingChangeOptions);
        }

        private void GetEBayListingProudctUrls()
        {
            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr;
            cmd.Connection = cn;
            cn.Open();

            string sqlString = @"SELECT CostcoUrl FROM eBay_CurrentListings WHERE DeleteDT is NULL";

            cmd.CommandText = sqlString;

            rdr = cmd.ExecuteReader();

            productUrlArray.Clear();

            while (rdr.Read())
            {
                productUrlArray.Add(rdr["CostcoUrl"].ToString());
            }

            rdr.Close();
            cn.Close();
        }

        public void runCrawl()
        {
            startDT = DateTime.Now;

            // test
            productUrlArray.Clear();
            productUrlArray.Add("http://www.costco.com/Steve-Madden-Leather-Passcase-Wallet.product.100242459.html");
            //productUrlArray.Add(@"file:///C:/Users/Jason%20Ding/Desktop/Jura%20Impressa%20F7%20Automatic%20Coffee%20Center.html");
            GetProductInfo(false);
            //end test

            if (string.IsNullOrEmpty(connectionString))
                SetConnectionString();

            //GetDepartmentArray();

            //GetProductUrls_New();

            //GetProductInfo();

            SecondTry(1);

            GetProductInfo(false);

            SecondTry(2);

            GetProductInfo(false);

            PopulateTables();

            CompareProducts();

            ArchieveProducts();

            endDT = DateTime.Now;

            SendEmail();
        }

        private void GetDepartmentArray()
        {
            string sqlString;

            categoryUrlArray.Clear();

            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            cn.Open();
            sqlString = "SELECT CategoryName FROM Costco_Departments WHERE bInclude = 1";
            cmd.CommandText = sqlString;
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    categoryUrlArray.Add(reader.GetString(0));
                }
            }
            reader.Close();
            cn.Close();
        }

        private void GetProductUrls_New()
        {
            productListPages.Clear();
            productUrlArray.Clear();

            driver = new FirefoxDriver(new FirefoxBinary(), new FirefoxProfile(), TimeSpan.FromSeconds(180));

            List<string> subCategory = new List<string>();

            int i = 0;

            while (i < categoryUrlArray.Count)
            {
                string url;
                if (categoryUrlArray[i].Contains("http"))
                    url = categoryUrlArray[i];
                else
                    url = "http://www.costco.com" + categoryUrlArray[i];

                driver.Navigate().GoToUrl(url);
                if (hasElement(driver, By.ClassName("categoryclist")))
                {
                    var categoryList = driver.FindElement(By.ClassName("categoryclist"));

                    var subCategoryList = categoryList.FindElements(By.ClassName("col-xs-6"));
                    foreach (var s in subCategoryList)
                    {
                        categoryUrlArray.Add(s.FindElement(By.TagName("a")).GetAttribute("href"));
                    }
                }

                if (hasElement(driver, By.ClassName("product-list")))
                {
                    var productList = driver.FindElement(By.ClassName("product-list"));

                    if (hasElement(productList, By.ClassName("paging")))
                    {
                        if (hasElement(productList, By.ClassName("page")))
                        {
                            foreach (var pg in productList.FindElements(By.ClassName("page")))
                            {
                                productListPages.Add(pg.FindElement(By.TagName("a")).GetAttribute("href"));
                            }
                        }
                        else
                        {
                            productListPages.Add(url);
                        }
                    }
                }

                i++;
            }

            foreach (var pl in productListPages)
            { 
                AddProductUrls(pl);
            }

            nScanProducts = productUrlArray.Count;

            driver.Close();

            productListEndDT = DateTime.Now;

        }

        private void AddProductUrls(string url)
        {
            driver.Navigate().GoToUrl(url);

            if (hasElement(driver, By.ClassName("product-list")))
            {
                var productList = driver.FindElement(By.ClassName("product-list"));

                foreach (var p in productList.FindElements(By.ClassName("product")))
                {
                    productUrlArray.Add(p.FindElement(By.TagName("a")).GetAttribute("href"));
                }
            }
        }

        private bool hasElement(IWebElement webElement, By by)
        {
            try
            {
                webElement.FindElement(by);
                return true;
            }
            catch (NoSuchElementException e)
            {
                return false;
            }
        }

        private bool hasElement(IWebDriver webDriver, By by)
        {
            try
            {
                webDriver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException e)
            {
                return false;
            }
        }

        private void GetProductInfo(bool bTruncate = true, bool bTruncateCostcoCategoriesTable = true)
        {
            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cn.Open();

            string sqlString;

            if (bTruncate)
            {
                sqlString = "TRUNCATE TABLE Raw_ProductInfo";
                cmd.CommandText = sqlString;
                cmd.ExecuteNonQuery();

                sqlString = "TRUNCATE TABLE Import_Skips";
                cmd.CommandText = sqlString;
                cmd.ExecuteNonQuery();

                if (bTruncateCostcoCategoriesTable)
                {
                    sqlString = "TRUNCATE TABLE Costco_Categories";
                    cmd.CommandText = sqlString;
                    cmd.ExecuteNonQuery();
                }

                sqlString = "TRUNCATE TABLE Import_Errors";
                cmd.CommandText = sqlString;
                cmd.ExecuteNonQuery();

                nScanProducts = 0;
                nImportProducts = 0;
                nSkipProducts = 0;
                nImportErrors = 0;


            }

            driver = new FirefoxDriver(new FirefoxBinary(), new FirefoxProfile(), TimeSpan.FromSeconds(180));
            driver.Navigate().GoToUrl("https://www.costco.com/LogonForm");
            IWebElement logonForm = driver.FindElement(By.Id("LogonForm"));
            logonForm.FindElement(By.Id("logonId")).SendKeys("zjding@gmail.com");
            logonForm.FindElement(By.Id("logonPassword")).SendKeys("721123");
            logonForm.FindElement(By.ClassName("submit")).Click();

            //productUrlArray.Clear();
            //productUrlArray.Add("http://www.costco.com/Orgain%c2%ae-Healthy-Kids-Organic-Shake-18ct--8.25oz-Chocolate.product.100083891.html");

            //IWebDriver driver = new FirefoxDriver();
            WebPage PageResult;

            int i = 1;

            foreach (string pu in productUrlArray)
            {
                try
                {
                    i++;
                    nScanProducts++;

                    string productUrl = HttpUtility.HtmlDecode(pu);
                    productUrl = productUrl.Replace("%2c", ",");

                    //string UrlNum = productUrl.Substring(0, productUrl.LastIndexOf('.'));
                    //UrlNum = UrlNum.Substring(UrlNum.LastIndexOf('.') + 1);

                    //PageResult = Browser.NavigateToPage(new Uri(productUrl));

                    //HtmlNode html = PageResult.Html;

                    //if (html.InnerText.Contains("Product Not Found"))
                    //{
                    //    sqlString = "INSERT INTO Import_Skips (Url, SkipPoint) VALUES ('" + pu.Replace(@"'", @"''") + "','" + "Product not found" + "')";
                    //    cmd.CommandText = sqlString;
                    //    cmd.ExecuteNonQuery();
                    //    nSkipProducts++;
                    //    continue;
                    //}

                    driver.Navigate().GoToUrl(productUrl);

                    // category
                    IWebElement eCrumbs = driver.FindElement(By.ClassName("crumbs"));
                    List<IWebElement> eCrumbArray = eCrumbs.FindElements(By.TagName("li")).ToList();


                    string stSubCategories = "";
                    int iCategory = 1;
                    string columns = "";
                    string values = "";
                    foreach (IWebElement crumb in eCrumbArray)
                    {
                        string temp = crumb.Text.Replace("\n", "");
                        temp = temp.Replace("\t", "");
                        temp = temp.Replace("'", "");
                        stSubCategories += temp + "|";

                        columns += "Category" + iCategory.ToString() + ",";
                        values += "'" + temp + "',";

                        iCategory++;
                    }
                    stSubCategories = stSubCategories.Substring(0, stSubCategories.Length - 1);
                    columns = columns.Substring(0, columns.Length - 1);
                    values = values.Substring(0, values.Length - 1);

                    IWebElement eProductDetails = driver.FindElement(By.Id("product-details"));

                    IWebElement eTitleAndItemNumber = eProductDetails.FindElement(By.ClassName("visible-xl"));

                    // name
                    IWebElement eTitle = eTitleAndItemNumber.FindElement(By.TagName("h1"));
                    string productName = eTitle.Text;
                    productName = productName.Replace("???", "");
                    productName = productName.Replace("??", "");
                    productName = productName.Trim();

                    // item number
                    IWebElement eItemNumber = eTitleAndItemNumber.FindElement(By.ClassName("item-number")).FindElement(By.TagName("span"));
                    string itemNumber = eItemNumber.Text;

                    // variants
                    string optionsString = string.Empty;
                    string imageOptions = string.Empty;
                    string imageLink = string.Empty;

                    if (hasElement(eProductDetails, By.Id("variants")))
                    {
                        var eVariants = eProductDetails.FindElement(By.Id("variants"));

                        var productOptions = eVariants.FindElements(By.ClassName("swatchDropdown"));

                        List<string> selectList = new List<string>();

                        foreach (var productOption in productOptions)
                        {
                            selectList.Add(productOption.FindElement(By.TagName("select")).GetAttribute("id").ToString());
                        }

                        if (selectList.Count == 2)
                        {
                            IWebElement selectElement0 = driver.FindElement(By.Id(selectList[0]));
                            var options0 = selectElement0.FindElements(By.TagName("option"));

                            foreach (IWebElement option0 in options0)
                            {
                                if (option0.GetAttribute("value").ToString().ToUpper() != "UNSELECTED")
                                {
                                    // optionsString
                                    string option0String = option0.Text;
                                    //string swatch0 = option0.GetAttribute("swatch") == string.Empty ? string.Empty : "(" + option0.GetAttribute("swatch") + ")";

                                    option0.Click();

                                    IWebElement selectElement1 = driver.FindElement(By.Id(selectList[1]));
                                    var options1 = selectElement1.FindElements(By.TagName("option"));

                                    optionsString += option0String + /*swatch0 +*/ ":";

                                    foreach (IWebElement option1 in options1)
                                    {
                                        if (option1.GetAttribute("value").ToString().ToUpper() != "UNSELECTED")
                                        {
                                            if (option1.Text.Contains("$"))
                                            {
                                                optionsString += option1.Text.Substring(0, option1.Text.LastIndexOf("- $") - 1) + ";";
                                            }
                                            else
                                            {
                                                optionsString += option1.Text + ";";
                                            }
                                        }
                                    }

                                    optionsString = optionsString.Substring(0, optionsString.Length - 1);
                                    optionsString += "|";

                                    // imagesString
                                    IWebElement thumb_holder = driver.FindElement(By.Id("thumbnails"));
                                    var thumblis = thumb_holder.FindElement(By.ClassName("slick-track")).FindElements(By.TagName("a"));

                                    imageOptions += option0String + /*swatch0 +*/ "=";

                                    foreach (IWebElement li in thumblis)
                                    {
                                        string imgUrl = li.FindElement(By.TagName("img")).GetAttribute("src");
                                        imgUrl = imgUrl.Replace(@"648", @"649");
                                        imageOptions += imgUrl + "|";
                                    }

                                    imageOptions = imageOptions.Substring(0, imageOptions.Length - 1);
                                    imageOptions += "~";
                                }
                            }

                            optionsString = optionsString.Substring(0, optionsString.Length - 1);
                            imageOptions = imageOptions.Substring(0, imageOptions.Length - 1);
                        }
                        else if (selectList.Count == 1)
                        {
                            IWebElement selectElement0 = driver.FindElement(By.Id(selectList[0]));
                            var options0 = selectElement0.FindElements(By.TagName("option"));
                            foreach (IWebElement option0 in options0)
                            {
                                if (option0.GetAttribute("value").ToString().ToUpper() != "UNSELECTED")
                                {
                                    if (option0.Text.Contains("$"))
                                    {
                                        int index = option0.Text.LastIndexOf("- $");
                                        optionsString += option0.Text.Substring(0, index - 1) + ";";
                                    }
                                    else
                                    {
                                        optionsString += option0.Text + ";";
                                    }
                                }
                            }
                            optionsString = optionsString.Substring(0, optionsString.Length - 1);

                            // imagesString
                            IWebElement thumb_holder = driver.FindElement(By.Id("thumbnails"));
                            var thumblis = thumb_holder.FindElement(By.ClassName("slick-track")).FindElements(By.TagName("a"));

                            foreach (IWebElement li in thumblis)
                            {
                                string imgUrl = li.FindElement(By.TagName("img")).GetAttribute("src");
                                imgUrl = imgUrl.Replace(@"648", @"649");
                                imageOptions += imgUrl + ";";
                            }

                            imageOptions = imageOptions.Substring(0, imageOptions.Length - 1);
                        }
                    }

                    // price
                    IWebElement ePrice = eProductDetails.FindElements(By.ClassName("form-group"))[0];
                    IWebElement eYourPrice = ePrice.FindElement(By.ClassName("your-price")).FindElement(By.ClassName("value"));
                    string price = eYourPrice.Text.Replace(",", "");

                    // marketing
                    IWebElement eMarketing = eProductDetails.FindElement(By.ClassName("marketing-container"));
                    string merchandisingText = hasElement(eMarketing, By.ClassName("merchandisingText")) ?
                                                eMarketing.FindElement(By.ClassName("merchandisingText")).Text : "";
                    string promotionalText = hasElement(eMarketing, By.ClassName("PromotionalText")) ?
                                                eMarketing.FindElement(By.ClassName("PromotionalText")).Text : "";

                    string discount = merchandisingText + " | " + promotionalText;

                    // feature
                    IWebElement eFeatures = eProductDetails.FindElement(By.ClassName("features-container"));

                    IWebElement eShipping = eFeatures.FindElement(By.Id("shipping-statement"));

                    string shipping = "0";

                    if (eShipping.Text.Contains("Included") || eShipping.Text.Contains("Free"))
                        shipping = "0";

                    int a = 1;
                    //HtmlNode productInfo = html.SelectSingleNode("//div[@id='product-details']");

                    //string productName = ((productInfo).SelectNodes("h1"))[0].InnerText;
                    //productName = productName.Replace("???", "");
                    //productName = productName.Replace("??", "");
                    //productName = productName.Trim();

                    //List<HtmlNode> topReviewPanelNode = productInfo.CssSelect(".top_review_panel").ToList<HtmlNode>();

                    //string discount = "";

                    //HtmlNode discountNote = topReviewPanelNode[0].SelectSingleNode("//p[@class='merchandisingText']");

                    //if (discountNote != null)
                    //{
                    //    discount = discountNote.InnerText.Replace("?", "");
                    //    discount = discountNote.InnerText.Replace("'", "");
                    //}



                    //List<HtmlNode> col1Node = productInfo.CssSelect(".col1").ToList<HtmlNode>();
                    //string itemNumber = (col1Node[0].SelectNodes("p")[0]).InnerText;
                    //if (itemNumber.ToUpper().Contains("ITEM") && itemNumber.Length > 6)
                    //    itemNumber = itemNumber.Substring(6);
                    //else
                    //    itemNumber = "";

                    //discountNote = col1Node[0].CssSelect(".merchandisingText").FirstOrDefault();

                    //if (discountNote != null)
                    //{
                    //    discount = discount.Length == 0 ? discountNote.InnerText.Replace("?", "") : discount + "; " + discountNote.InnerText.Replace("?", "");
                    //    discount = discount.Replace("?", "");
                    //    discount = discount.Replace("'", "");
                    //}

                    //discount = discount.Replace("Free Shipping", "");

                    //string price;
                    //List<HtmlNode> yourPriceNode = col1Node.CssSelect(".your-price").ToList<HtmlNode>();
                    //if (yourPriceNode.Count > 0)
                    //{
                    //    List<HtmlNode> priceNode = yourPriceNode[0].CssSelect(".currency").ToList<HtmlNode>();
                    //    price = priceNode[0].InnerText;
                    //    price = price.Replace("$", "");
                    //    price = price.Replace(",", "");

                    //    if (price == "- -")
                    //        price = "-2";
                    //}
                    //else
                    //{
                    //    price = "-1";
                    //}

                    //var productOptionsNode = col1Node.CssSelect(".product-option").FirstOrDefault();

                    //string shipping = "0";

                    //var productSHNode = col1Node[0].SelectSingleNode("//li[@class='product']");

                    //string optionsString = string.Empty;
                    //string imageOptions = string.Empty;
                    //string imageLink = string.Empty;

                    //if (productSHNode != null)
                    //{
                    //    if (productSHNode.InnerText.ToUpper().Contains("INCLUDED") || productSHNode.InnerText.ToUpper().Contains("INLCUDED"))
                    //    {
                    //        shipping = "0";
                    //    }
                    //    else
                    //    {
                    //        string shString = productSHNode.InnerText;
                    //        int nDollar = shString.IndexOf("$");
                    //        if (nDollar > 0)
                    //        {
                    //            shString = shString.Substring(nDollar + 1);
                    //            int nStar = shString.IndexOf("*");
                    //            if (nStar == -1)
                    //                nStar = shString.IndexOf(" ");
                    //            shString = shString.Substring(0, nStar);
                    //            shString = shString.Replace(" ", "");
                    //            shipping = shString;
                    //        }
                    //        else
                    //        {
                    //            int nShipping = shString.IndexOf("Shipping");
                    //            int nQuantity = shString.ToUpper().IndexOf("QUANTITY");

                    //            if (nShipping == -1 || nQuantity == -1)
                    //            {
                    //                sqlString = "INSERT INTO Import_Skips (Url, SkipPoint) VALUES ('" + pu.Replace(@"'", @"''") + "','" + "Shipping and Quantity" + "')";
                    //                cmd.CommandText = sqlString;
                    //                cmd.ExecuteNonQuery();
                    //                nSkipProducts++;
                    //                continue;
                    //            }

                    //            shString = shString.Substring(nShipping, nQuantity);
                    //            Char[] strarr = shString.ToCharArray().Where(c => Char.IsDigit(c) || c.Equals('.')).ToArray();
                    //            decimal number = Convert.ToDecimal(new string(strarr));
                    //            shipping = number.ToString();
                    //        }
                    //    }
                    //}

                    //if (string.IsNullOrEmpty(string.Empty))
                    //{
                    //    HtmlNode imageColumnNode = html.CssSelect(".image-column").ToList<HtmlNode>().First();

                    //    HtmlNode imageNode = imageColumnNode.SelectSingleNode("//img[@itemprop='image']");

                    //    imageLink = (imageNode.Attributes["src"]).Value;
                    //}

                    //#region
                    //if (productSHNode.InnerText.ToUpper().Contains("OPTIONS"))
                    //{

                    //    driver.Navigate().GoToUrl(productUrl);
                    //    var productOptions = driver.FindElements(By.ClassName("product-option"));

                    //    List<string> selectList = new List<string>();

                    //    foreach (var productOption in productOptions)
                    //    {
                    //        selectList.Add(productOption.FindElement(By.TagName("select")).GetAttribute("id").ToString());
                    //    }

                    //    if (selectList.Count == 2)
                    //    {
                    //        IWebElement selectElement0 = driver.FindElement(By.Id(selectList[0]));
                    //        var options0 = selectElement0.FindElements(By.TagName("option"));
                    //        foreach (IWebElement option0 in options0)
                    //        {
                    //            if (option0.GetAttribute("value").ToString().ToUpper() != "UNSELECTED")
                    //            {
                    //                // optionsString
                    //                string option0String = option0.Text;
                    //                //string swatch0 = option0.GetAttribute("swatch") == string.Empty ? string.Empty : "(" + option0.GetAttribute("swatch") + ")";

                    //                option0.Click();

                    //                IWebElement selectElement1 = driver.FindElement(By.Id(selectList[1]));
                    //                var options1 = selectElement1.FindElements(By.TagName("option"));

                    //                optionsString += option0String + /*swatch0 +*/ ":";

                    //                foreach (IWebElement option1 in options1)
                    //                {
                    //                    if (option1.GetAttribute("value").ToString().ToUpper() != "UNSELECTED")
                    //                    {
                    //                        if (option1.Text.Contains("$"))
                    //                        {
                    //                            optionsString += option1.Text.Substring(0, option1.Text.LastIndexOf("- $") - 1) + ";";
                    //                        }
                    //                        else
                    //                        {
                    //                            optionsString += option1.Text + ";";
                    //                        }
                    //                    }
                    //                }

                    //                optionsString = optionsString.Substring(0, optionsString.Length - 1);
                    //                optionsString += "|";

                    //                // imagesString
                    //                IWebElement thumb_holder = driver.FindElement(By.Id("thumb_holder"));
                    //                var thumblis = thumb_holder.FindElements(By.TagName("li"));

                    //                imageOptions += option0String + /*swatch0 +*/ "=";

                    //                foreach (IWebElement li in thumblis)
                    //                {
                    //                    string imgUrl = li.FindElement(By.TagName("img")).GetAttribute("src");
                    //                    imgUrl = imgUrl.Replace(@"/50-", @"/500-");
                    //                    imageOptions += imgUrl + "|";
                    //                }

                    //                imageOptions = imageOptions.Substring(0, imageOptions.Length - 1);
                    //                imageOptions += "~";
                    //            }
                    //        }

                    //        optionsString = optionsString.Substring(0, optionsString.Length - 1);
                    //        imageOptions = imageOptions.Substring(0, imageOptions.Length - 1);

                    //    }
                    //    else if (selectList.Count == 1)
                    //    {
                    //        IWebElement selectElement0 = driver.FindElement(By.Id(selectList[0]));
                    //        var options0 = selectElement0.FindElements(By.TagName("option"));
                    //        foreach (IWebElement option0 in options0)
                    //        {
                    //            if (option0.GetAttribute("value").ToString().ToUpper() != "UNSELECTED")
                    //            {
                    //                if (option0.Text.Contains("$"))
                    //                {
                    //                    optionsString += option0.Text.Substring(0, option0.Text.LastIndexOf("- $") - 1) + ";";
                    //                }
                    //                else
                    //                {
                    //                    optionsString += option0.Text + ";";
                    //                }
                    //            }
                    //        }
                    //        optionsString = optionsString.Substring(0, optionsString.Length - 1);

                    //        // imagesString
                    //        IWebElement thumb_holder = driver.FindElement(By.Id("thumb_holder"));
                    //        var thumblis = thumb_holder.FindElements(By.TagName("li"));

                    //        foreach (IWebElement li in thumblis)
                    //        {
                    //            string imgUrl = li.FindElement(By.TagName("img")).GetAttribute("src");
                    //            imgUrl = imgUrl.Replace(@"/50-", @"/500-");
                    //            imageOptions += imgUrl + ";";
                    //        }

                    //        imageOptions = imageOptions.Substring(0, imageOptions.Length - 1);
                    //    }
                    //}
                    //#endregion

                    //if (firstTry.Contains(pu))
                    //    firstTryResult.Add(pu);

                    //if (secondTry.Contains(pu))
                    //    secondTryResult.Add(pu);

                    //sqlString = "INSERT INTO Raw_ProductInfo (Name, UrlNumber, ItemNumber, Category, Price, Shipping, Discount, ImageLink, ImageOptions, Url, Options) VALUES ('" + productName.Replace("'", "''") + "','" + UrlNum + "','" + itemNumber + "','" + stSubCategories + "'," + price + "," + shipping + "," + "'" + discount + "','" + imageLink.Replace("'", "''") + "','" + imageOptions.Replace("'", "''") + "','" + productUrl.Replace("'", "''") + "','" + optionsString + "')";
                    //cmd.CommandText = sqlString;
                    //cmd.ExecuteNonQuery();
                    //nImportProducts++;

                    //sqlString = @"IF NOT EXISTS (SELECT * FROM Costco_Categories WHERE ";
                    //int j = 1;
                    //foreach (var c in stSubCategories.Split('|'))
                    //{
                    //    sqlString += "Category" + j.ToString() + "='" + c + "'";
                    //    sqlString += " AND ";

                    //    j++;
                    //}
                    //for (int k = j; k <= 8; k++)
                    //{
                    //    sqlString += "Category" + k.ToString() + " is NULL";
                    //    if (k < 8)
                    //    {
                    //        sqlString += " AND ";
                    //    }
                    //}
                    //sqlString += @") BEGIN
                    //                INSERT INTO Costco_Categories (" + columns + ") VALUES (" + values + ") END";
                    //cmd.CommandText = sqlString;
                    //cmd.ExecuteNonQuery();

                    //sqlString = @"IF NOT EXISTS (SELECT * FROM Costco_eBay_Categories WHERE ";
                    //j = 1;
                    //foreach (var c in stSubCategories.Split('|'))
                    //{
                    //    sqlString += "Category" + j.ToString() + "='" + c + "'";
                    //    sqlString += " AND ";

                    //    j++;
                    //}
                    //for (int k = j; k <= 8; k++)
                    //{
                    //    sqlString += "Category" + k.ToString() + " is NULL";
                    //    if (k < 8)
                    //    {
                    //        sqlString += " AND ";
                    //    }
                    //}
                    //sqlString += @") BEGIN
                    //                INSERT INTO Costco_eBay_Categories (" + columns + ") VALUES (" + values + ") END";
                    //cmd.CommandText = sqlString;
                    //cmd.ExecuteNonQuery();
                }
                catch (Exception exception)
                {
                    string productUrl = HttpUtility.HtmlDecode(pu);
                    productUrl = productUrl.Replace("%2c", ",");
                    productUrl = productUrl.Replace(@"'", @"''");
                    sqlString = "INSERT INTO Import_Errors (Url, Exception) VALUES ('" + productUrl + "','" + exception.Message.Replace(@"'", @"''") + "')";
                    cmd.CommandText = sqlString;
                    cmd.ExecuteNonQuery();
                    nImportErrors++;

                    continue;
                }
            }

            cn.Close();

            driver.Close();
        }

        //private void GetProductInfo(bool bTruncate = true, bool bTruncateCostcoCategoriesTable = true)
        //{
        //    SqlConnection cn = new SqlConnection(connectionString);
        //    SqlCommand cmd = new SqlCommand();
        //    cmd.Connection = cn;
        //    cn.Open();

        //    string sqlString;

        //    if (bTruncate)
        //    {
        //        sqlString = "TRUNCATE TABLE Raw_ProductInfo";
        //        cmd.CommandText = sqlString;
        //        cmd.ExecuteNonQuery();

        //        sqlString = "TRUNCATE TABLE Import_Skips";
        //        cmd.CommandText = sqlString;
        //        cmd.ExecuteNonQuery();

        //        if (bTruncateCostcoCategoriesTable)
        //        {
        //            sqlString = "TRUNCATE TABLE Costco_Categories";
        //            cmd.CommandText = sqlString;
        //            cmd.ExecuteNonQuery();
        //        }

        //        sqlString = "TRUNCATE TABLE Import_Errors";
        //        cmd.CommandText = sqlString;
        //        cmd.ExecuteNonQuery();

        //        nScanProducts = 0;
        //        nImportProducts = 0;
        //        nSkipProducts = 0;
        //        nImportErrors = 0;


        //    }

        //    driver = new FirefoxDriver(new FirefoxBinary(), new FirefoxProfile(), TimeSpan.FromSeconds(180));
        //    driver.Navigate().GoToUrl("https://www.costco.com/LogonForm");
        //    IWebElement logonForm = driver.FindElement(By.Id("LogonForm"));
        //    logonForm.FindElement(By.Id("logonId")).SendKeys("zjding@gmail.com");
        //    logonForm.FindElement(By.Id("logonPassword")).SendKeys("721123");
        //    logonForm.FindElement(By.ClassName("submit")).Click();

        //    //productUrlArray.Clear();
        //    //productUrlArray.Add("http://www.costco.com/Orgain%c2%ae-Healthy-Kids-Organic-Shake-18ct--8.25oz-Chocolate.product.100083891.html");

        //    //IWebDriver driver = new FirefoxDriver();
        //    WebPage PageResult;

        //    int i = 1;

        //    foreach (string pu in productUrlArray)
        //    {
        //        try
        //        {
        //            i++;
        //            nScanProducts++;

        //            string productUrl = HttpUtility.HtmlDecode(pu);
        //            productUrl = productUrl.Replace("%2c", ",");

        //            string UrlNum = productUrl.Substring(0, productUrl.LastIndexOf('.'));
        //            UrlNum = UrlNum.Substring(UrlNum.LastIndexOf('.') + 1);

        //            PageResult = Browser.NavigateToPage(new Uri(productUrl));

        //            HtmlNode html = PageResult.Html;

        //            if (html.InnerText.Contains("Product Not Found"))
        //            {
        //                sqlString = "INSERT INTO Import_Skips (Url, SkipPoint) VALUES ('" + pu.Replace(@"'", @"''") + "','" + "Product not found" + "')";
        //                cmd.CommandText = sqlString;
        //                cmd.ExecuteNonQuery();
        //                nSkipProducts++;
        //                continue;
        //            }

        //            string stSubCategories = "";

        //            HtmlNode category = html.SelectSingleNode("//ul[@id='breadcrumbs']");

        //            List<HtmlNode> subCategories = category.SelectNodes("li").ToList();

        //            int iCategory = 1;
        //            string columns = "";
        //            string values = "";
        //            foreach (HtmlNode subCategory in subCategories)
        //            {
        //                string temp = subCategory.InnerText.Replace("\n", "");
        //                temp = temp.Replace("\t", "");
        //                temp = temp.Replace("'", "");
        //                stSubCategories += temp + "|";

        //                columns += "Category" + iCategory.ToString() + ",";
        //                values += "'" + temp + "',";

        //                iCategory++;
        //            }
        //            stSubCategories = stSubCategories.Substring(0, stSubCategories.Length - 1);
        //            columns = columns.Substring(0, columns.Length - 1);
        //            values = values.Substring(0, values.Length - 1);

        //            HtmlNode productInfo = html.SelectSingleNode("//div[@class='product-info']");

        //            List<HtmlNode> topReviewPanelNode = productInfo.CssSelect(".top_review_panel").ToList<HtmlNode>();

        //            string productName = (topReviewPanelNode[0].SelectNodes("h1"))[0].InnerText;
        //            productName = productName.Replace("???", "");
        //            productName = productName.Replace("??", "");
        //            productName = productName.Trim();



        //            string discount = "";

        //            HtmlNode discountNote = topReviewPanelNode[0].SelectSingleNode("//p[@class='merchandisingText']");

        //            if (discountNote != null)
        //            {
        //                discount = discountNote.InnerText.Replace("?", "");
        //                discount = discountNote.InnerText.Replace("'", "");
        //            }



        //            List<HtmlNode> col1Node = productInfo.CssSelect(".col1").ToList<HtmlNode>();
        //            string itemNumber = (col1Node[0].SelectNodes("p")[0]).InnerText;
        //            if (itemNumber.ToUpper().Contains("ITEM") && itemNumber.Length > 6)
        //                itemNumber = itemNumber.Substring(6);
        //            else
        //                itemNumber = "";

        //            discountNote = col1Node[0].CssSelect(".merchandisingText").FirstOrDefault();

        //            if (discountNote != null)
        //            {
        //                discount = discount.Length == 0 ? discountNote.InnerText.Replace("?", "") : discount + "; " + discountNote.InnerText.Replace("?", "");
        //                discount = discount.Replace("?", "");
        //                discount = discount.Replace("'", "");
        //            }

        //            discount = discount.Replace("Free Shipping", "");

        //            string price;
        //            List<HtmlNode> yourPriceNode = col1Node.CssSelect(".your-price").ToList<HtmlNode>();
        //            if (yourPriceNode.Count > 0)
        //            {
        //                List<HtmlNode> priceNode = yourPriceNode[0].CssSelect(".currency").ToList<HtmlNode>();
        //                price = priceNode[0].InnerText;
        //                price = price.Replace("$", "");
        //                price = price.Replace(",", "");

        //                if (price == "- -")
        //                    price = "-2";
        //            }
        //            else
        //            {
        //                price = "-1";
        //            }

        //            var productOptionsNode = col1Node.CssSelect(".product-option").FirstOrDefault();

        //            string shipping = "0";

        //            var productSHNode = col1Node[0].SelectSingleNode("//li[@class='product']");

        //            string optionsString = string.Empty;
        //            string imageOptions = string.Empty;
        //            string imageLink = string.Empty;

        //            if (productSHNode != null)
        //            {
        //                if (productSHNode.InnerText.ToUpper().Contains("INCLUDED") || productSHNode.InnerText.ToUpper().Contains("INLCUDED"))
        //                {
        //                    shipping = "0";
        //                }
        //                else
        //                {
        //                    string shString = productSHNode.InnerText;
        //                    int nDollar = shString.IndexOf("$");
        //                    if (nDollar > 0)
        //                    {
        //                        shString = shString.Substring(nDollar + 1);
        //                        int nStar = shString.IndexOf("*");
        //                        if (nStar == -1)
        //                            nStar = shString.IndexOf(" ");
        //                        shString = shString.Substring(0, nStar);
        //                        shString = shString.Replace(" ", "");
        //                        shipping = shString;
        //                    }
        //                    else
        //                    {
        //                        int nShipping = shString.IndexOf("Shipping");
        //                        int nQuantity = shString.ToUpper().IndexOf("QUANTITY");

        //                        if (nShipping == -1 || nQuantity == -1)
        //                        {
        //                            sqlString = "INSERT INTO Import_Skips (Url, SkipPoint) VALUES ('" + pu.Replace(@"'", @"''") + "','" + "Shipping and Quantity" + "')";
        //                            cmd.CommandText = sqlString;
        //                            cmd.ExecuteNonQuery();
        //                            nSkipProducts++;
        //                            continue;
        //                        }

        //                        shString = shString.Substring(nShipping, nQuantity);
        //                        Char[] strarr = shString.ToCharArray().Where(c => Char.IsDigit(c) || c.Equals('.')).ToArray();
        //                        decimal number = Convert.ToDecimal(new string(strarr));
        //                        shipping = number.ToString();
        //                    }
        //                }
        //            }

        //            if (string.IsNullOrEmpty(string.Empty))
        //            {
        //                HtmlNode imageColumnNode = html.CssSelect(".image-column").ToList<HtmlNode>().First();

        //                HtmlNode imageNode = imageColumnNode.SelectSingleNode("//img[@itemprop='image']");

        //                imageLink = (imageNode.Attributes["src"]).Value;
        //            }

        //            #region
        //            if (productSHNode.InnerText.ToUpper().Contains("OPTIONS"))
        //            {

        //                driver.Navigate().GoToUrl(productUrl);
        //                var productOptions = driver.FindElements(By.ClassName("product-option"));

        //                List<string> selectList = new List<string>();

        //                foreach (var productOption in productOptions)
        //                {
        //                    selectList.Add(productOption.FindElement(By.TagName("select")).GetAttribute("id").ToString());
        //                }

        //                if (selectList.Count == 2)
        //                {
        //                    IWebElement selectElement0 = driver.FindElement(By.Id(selectList[0]));
        //                    var options0 = selectElement0.FindElements(By.TagName("option"));
        //                    foreach (IWebElement option0 in options0)
        //                    {
        //                        if (option0.GetAttribute("value").ToString().ToUpper() != "UNSELECTED")
        //                        {
        //                            // optionsString
        //                            string option0String = option0.Text;
        //                            //string swatch0 = option0.GetAttribute("swatch") == string.Empty ? string.Empty : "(" + option0.GetAttribute("swatch") + ")";

        //                            option0.Click();

        //                            IWebElement selectElement1 = driver.FindElement(By.Id(selectList[1]));
        //                            var options1 = selectElement1.FindElements(By.TagName("option"));

        //                            optionsString += option0String + /*swatch0 +*/ ":";

        //                            foreach (IWebElement option1 in options1)
        //                            {
        //                                if (option1.GetAttribute("value").ToString().ToUpper() != "UNSELECTED")
        //                                {
        //                                    if (option1.Text.Contains("$"))
        //                                    {
        //                                        optionsString += option1.Text.Substring(0, option1.Text.LastIndexOf("- $") - 1) + ";";
        //                                    }
        //                                    else
        //                                    {
        //                                        optionsString += option1.Text + ";";
        //                                    }
        //                                }
        //                            }

        //                            optionsString = optionsString.Substring(0, optionsString.Length - 1);
        //                            optionsString += "|";

        //                            // imagesString
        //                            IWebElement thumb_holder = driver.FindElement(By.Id("thumb_holder"));
        //                            var thumblis = thumb_holder.FindElements(By.TagName("li"));

        //                            imageOptions += option0String + /*swatch0 +*/ "=";

        //                            foreach (IWebElement li in thumblis)
        //                            {
        //                                string imgUrl = li.FindElement(By.TagName("img")).GetAttribute("src");
        //                                imgUrl = imgUrl.Replace(@"/50-", @"/500-");
        //                                imageOptions += imgUrl + "|";
        //                            }

        //                            imageOptions = imageOptions.Substring(0, imageOptions.Length - 1);
        //                            imageOptions += "~";
        //                        }
        //                    }

        //                    optionsString = optionsString.Substring(0, optionsString.Length - 1);
        //                    imageOptions = imageOptions.Substring(0, imageOptions.Length - 1);

        //                }
        //                else if (selectList.Count == 1)
        //                {
        //                    IWebElement selectElement0 = driver.FindElement(By.Id(selectList[0]));
        //                    var options0 = selectElement0.FindElements(By.TagName("option"));
        //                    foreach (IWebElement option0 in options0)
        //                    {
        //                        if (option0.GetAttribute("value").ToString().ToUpper() != "UNSELECTED")
        //                        {
        //                            if (option0.Text.Contains("$"))
        //                            {
        //                                optionsString += option0.Text.Substring(0, option0.Text.LastIndexOf("- $") - 1) + ";";
        //                            }
        //                            else
        //                            {
        //                                optionsString += option0.Text + ";";
        //                            }
        //                        }
        //                    }
        //                    optionsString = optionsString.Substring(0, optionsString.Length - 1);

        //                    // imagesString
        //                    IWebElement thumb_holder = driver.FindElement(By.Id("thumb_holder"));
        //                    var thumblis = thumb_holder.FindElements(By.TagName("li"));

        //                    foreach (IWebElement li in thumblis)
        //                    {
        //                        string imgUrl = li.FindElement(By.TagName("img")).GetAttribute("src");
        //                        imgUrl = imgUrl.Replace(@"/50-", @"/500-");
        //                        imageOptions += imgUrl + ";";
        //                    }

        //                    imageOptions = imageOptions.Substring(0, imageOptions.Length - 1);
        //                }
        //            }
        //            #endregion

        //            //if (firstTry.Contains(pu))
        //            //    firstTryResult.Add(pu);

        //            //if (secondTry.Contains(pu))
        //            //    secondTryResult.Add(pu);

        //            sqlString = "INSERT INTO Raw_ProductInfo (Name, UrlNumber, ItemNumber, Category, Price, Shipping, Discount, ImageLink, ImageOptions, Url, Options) VALUES ('" + productName.Replace("'", "''") + "','" + UrlNum + "','" + itemNumber + "','" + stSubCategories + "'," + price + "," + shipping + "," + "'" + discount + "','" + imageLink.Replace("'", "''") + "','" + imageOptions.Replace("'", "''") + "','" + productUrl.Replace("'", "''") + "','" + optionsString + "')";
        //            cmd.CommandText = sqlString;
        //            cmd.ExecuteNonQuery();
        //            nImportProducts++;

        //            sqlString = @"IF NOT EXISTS (SELECT * FROM Costco_Categories WHERE ";
        //            int j = 1;
        //            foreach (var c in stSubCategories.Split('|'))
        //            {
        //                sqlString += "Category" + j.ToString() + "='" + c + "'";
        //                sqlString += " AND ";

        //                j++;
        //            }
        //            for (int k = j; k <= 8; k++)
        //            {
        //                sqlString += "Category" + k.ToString() + " is NULL";
        //                if (k < 8)
        //                {
        //                    sqlString += " AND ";
        //                }
        //            }
        //            sqlString += @") BEGIN
        //                            INSERT INTO Costco_Categories (" + columns + ") VALUES (" + values + ") END";
        //            cmd.CommandText = sqlString;
        //            cmd.ExecuteNonQuery();

        //            sqlString = @"IF NOT EXISTS (SELECT * FROM Costco_eBay_Categories WHERE ";
        //            j = 1;
        //            foreach (var c in stSubCategories.Split('|'))
        //            {
        //                sqlString += "Category" + j.ToString() + "='" + c + "'";
        //                sqlString += " AND ";

        //                j++;
        //            }
        //            for (int k = j; k <= 8; k++)
        //            {
        //                sqlString += "Category" + k.ToString() + " is NULL";
        //                if (k < 8)
        //                {
        //                    sqlString += " AND ";
        //                }
        //            }
        //            sqlString += @") BEGIN
        //                            INSERT INTO Costco_eBay_Categories (" + columns + ") VALUES (" + values + ") END";
        //            cmd.CommandText = sqlString;
        //            cmd.ExecuteNonQuery();
        //        }
        //        catch (Exception exception)
        //        {
        //            string productUrl = HttpUtility.HtmlDecode(pu);
        //            productUrl = productUrl.Replace("%2c", ",");
        //            productUrl = productUrl.Replace(@"'", @"''");
        //            sqlString = "INSERT INTO Import_Errors (Url, Exception) VALUES ('" + productUrl + "','" + exception.Message.Replace(@"'", @"''") + "')";
        //            cmd.CommandText = sqlString;
        //            cmd.ExecuteNonQuery();
        //            nImportErrors++;

        //            continue;
        //        }
        //    }

        //    cn.Close();

        //    driver.Close();
        //}

        private void SecondTry(int i = 0)
        {
            productUrlArray.Clear();

            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            cn.Open();

            //string sqlString = @"select * from ProductInfo p 
            //            where 
            //            not exists
            //            (select 1 from Raw_ProductInfo sp where sp.UrlNumber = p.UrlNumber)";

            string sqlString = @"select Url from Import_Skips 
                        where SkipPoint = 'Product not found'";

            cmd.CommandText = sqlString;
            SqlDataReader rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                productUrlArray.Add(rdr["Url"].ToString());
                //if (i == 1)
                //{
                //    firstTry.Add(rdr["Url"].ToString());
                //}
                //else if (i == 2)
                //{
                //    secondTry.Add(rdr["Url"].ToString());
                //}staging_productInfo
            }

            rdr.Close();

            sqlString = @"select Url from Import_Errors";

            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                productUrlArray.Add(rdr["Url"].ToString());
            }

            rdr.Close();

            sqlString = @"delete from Import_Skips 
                        where SkipPoint = 'Product not found'";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            cn.Close();
        }

        private void PopulateTables()
        {
            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            // copy to staging_productInfo
            cn.Open();
            string sqlString = "TRUNCATE TABLE Staging_ProductInfo";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"insert into dbo.staging_productInfo (Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, imageOptions, url, options)
                        select distinct Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, imageOptions, url, options
                        from dbo.Raw_ProductInfo
                        where Price > 0
                        order by UrlNumber";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            //// copy to staging_productInfo_filtered
            //sqlString = "TRUNCATE TABLE Staging_ProductInfo_Filtered";
            //cmd.CommandText = sqlString;
            //cmd.ExecuteNonQuery();

            //sqlString = @"insert into dbo.Staging_ProductInfo_Filtered(Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url)
            //            select distinct Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url 
            //            from dbo.Raw_ProductInfo 
            //            where Price > 0 and Price < 100 and Shipping = 0
            //            order by UrlNumber";

            //cmd.CommandText = sqlString;
            //cmd.ExecuteNonQuery();


        }

        private void CompareProducts()
        {
            CompareCostcoInventory();

            int nEBayListingChangePriceUp = 0;
            int nEBayListingChangePriceDown = 0;
            int nEBayListingChangeDiscontinue = 0;
            int nEBayListingChangeOptions = 0;

            CompareEBayListings(out nEBayListingChangePriceUp, out nEBayListingChangePriceDown, out nEBayListingChangeDiscontinue, out nEBayListingChangeOptions );
        }

        private void CompareCostcoInventory()
        {
            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            SqlDataReader rdr;

            cn.Open();

            string sqlString = string.Empty;

            sqlString = @"  TRUNCATE TABLE CostcoInventoryChange_Discontinue; 
                            DELETE FROM CostcoInventoryChange_New where InsertTime < DATEADD(day, -1, GETDATE()); 
                            TRUNCATE TABLE CostcoInventoryChange_PriceUp;
                            TRUNCATE TABLE CostcoInventoryChange_PriceDown;   ";

            //sqlString = @"  TRUNCATE TABLE CostcoInventoryChange_Discontinue; 
            //                TRUNCATE TABLE CostcoInventoryChange_New; 
            //                TRUNCATE TABLE CostcoInventoryChange_PriceUp;
            //                TRUNCATE TABLE CostcoInventoryChange_PriceDown;   ";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            // price up
            sqlString = @"select s.Name, s.Price as newPrice, p.Price as oldPrice, s.Url 
                                from [dbo].[Staging_ProductInfo] s, [dbo].[ProductInfo] p
                                where s.UrlNumber = p.UrlNumber
                                and s.Price > p.Price";
            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            priceUpProductArray.Clear();

            while (rdr.Read())
            {
                priceUpProductArray.Add("<a href='" + rdr["Url"].ToString().Replace("'", "&#39;") + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["newPrice"].ToString() + "|(" + rdr["oldPrice"].ToString() + ")");
            }

            rdr.Close();

            sqlString = @"  INSERT INTO CostcoInventoryChange_PriceUp (Name, CostcoUrl, UrlNumber, CostcoOldPrice, CostcoNewPrice, ImageLink) 
                            SELECT s.Name, s.Url, s.UrlNumber, p.Price, s.Price, s.ImageLink
                            FROM [dbo].[Staging_ProductInfo] s, [dbo].[ProductInfo] p
                            WHERE s.UrlNumber = p.UrlNumber
                            AND s.Price > p.Price ";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            // price down
            sqlString = @"select s.Name, s.Price as newPrice, p.Price as oldPrice, s.Url from 
                        [dbo].[Staging_ProductInfo] s, [dbo].[ProductInfo] p
                        where s.UrlNumber = p.UrlNumber
                        and s.Price < p.Price";
            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            priceDownProductArray.Clear();

            while (rdr.Read())
            {
                priceDownProductArray.Add("<a href='" + rdr["Url"].ToString().Replace("'", "&#39;") + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["newPrice"].ToString() + "|(" + rdr["oldPrice"].ToString() + ")");
            }

            rdr.Close();

            sqlString = @"  INSERT INTO CostcoInventoryChange_PriceDown (Name, CostcoUrl, UrlNumber, CostcoOldPrice, CostcoNewPrice, ImageLink) 
                            SELECT s.Name, s.Url, s.UrlNumber, p.Price, s.Price, s.ImageLink
                            FROM [dbo].[Staging_ProductInfo] s, [dbo].[ProductInfo] p
                            WHERE s.UrlNumber = p.UrlNumber
                            AND s.Price < p.Price ";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            // new products
            sqlString = @"select * from Staging_ProductInfo sp
                        where 
                        not exists
                        (select 1 from ProductInfo p  where sp.UrlNumber = p.UrlNumber)";
            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            newProductArray.Clear();

            while (rdr.Read())
            {
                newProductArray.Add("<a href='" + rdr["Url"].ToString().Replace("'", "&#39;") + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["Price"].ToString());
            }

            rdr.Close();

            sqlString = @"  INSERT INTO CostcoInventoryChange_New ([Name]
                                                                  ,[UrlNumber]
                                                                  ,[ItemNumber]
                                                                  ,[Category]
                                                                  ,[Price]
                                                                  ,[Shipping]
                                                                  ,[Limit]
                                                                  ,[Discount]
                                                                  ,[Details]
                                                                  ,[Specification]
                                                                  ,[CostcoUrl]
                                                                  ,[Options]
                                                                  ,[ImageLink]
                                                                  ,[ImageOptions]
                                                                  ,[NumberOfImage]
                                                                  ,InsertTime)
                            SELECT [Name]
                                  ,[UrlNumber]
                                  ,[ItemNumber]
                                  ,[Category]
                                  ,[Price]
                                  ,[Shipping]
                                  ,[Limit]
                                  ,[Discount]
                                  ,[Details]
                                  ,[Specification]
                                  ,[Url]
                                  ,[Options]
                                  ,[ImageLink]
                                  ,[ImageOptions]
                                  ,[NumberOfImage] 
                                  ,GETDATE()
                            from Staging_ProductInfo sp
                            WHERE 
                            NOT EXISTS
                            (SELECT 1 FROM ProductInfo p  WHERE sp.UrlNumber = p.UrlNumber)
                            AND
                            NOT EXISTS 
                            (SELECT 1 FROM CostcoInventoryChange_New n WHERE n.CostcoUrl = sp.UrlNumber)";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            // discontinued products
            sqlString = @"select * from ProductInfo p 
                        where 
                        not exists
                        (select 1 from Staging_ProductInfo sp where sp.UrlNumber = p.UrlNumber)";
            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            discontinueddProductArray.Clear();

            while (rdr.Read())
            {
                discontinueddProductArray.Add("<a href='" + rdr["Url"].ToString().Replace("'", "&#39;") + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["Price"].ToString());
            }

            rdr.Close();

            sqlString = @"  INSERT INTO CostcoInventoryChange_Discontinue (Name, CostcoUrl, UrlNumber)
                            SELECT p.Name, p.Url, p.UrlNumber 
                            FROM ProductInfo p 
                            WHERE 
                            NOT EXISTS
                            (SELECT 1 FROM Staging_ProductInfo sp WHERE sp.UrlNumber = p.UrlNumber)";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            // stockChange products
            sqlString = @"select s.Name, s.Url from 
                        [dbo].[Staging_ProductInfo] s, [dbo].[ProductInfo] p
                        where s.UrlNumber = p.UrlNumber
                        and s.Options <> p.Options";
            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            stockChangeProductArray.Clear();

            while (rdr.Read())
            {
                stockChangeProductArray.Add("<a href='" + rdr["Url"].ToString().Replace("'", "&#39;") + "'>" + rdr["Name"].ToString() + "</a>");
            }

            rdr.Close();
        }

        private void CompareEBayListings(out int nEBayListingChangePriceUp, out int nEBayListingChangePriceDown, out int nEBayListingChangeDiscontinue, out int nEBayListingChangeOptions)
        {
            nEBayListingChangePriceUp = 0;

            string sqlString;

            SqlConnection cn = new SqlConnection(connectionString);

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            cn.Open();

            sqlString = "TRUNCATE TABLE eBayListingChange_PriceUp";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = "TRUNCATE TABLE eBayListingChange_PriceDown";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = "TRUNCATE TABLE eBayListingChange_Discontinue";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = "TRUNCATE TABLE eBayListingChange_OptionChange";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            // eBay listing price up
            sqlString = @"INSERT INTO [dbo].[eBayListingChange_PriceUp] (Name, CostcoUrl, UrlNumber, eBayItemNumber, CostcoOldPrice, CostcoNewPrice, eBayListingOldPrice, ImageLink)
                                select p.name as Name, p.Url as CostcoUrl, p.UrlNumber as UrlNumber, l.eBayItemNumber as eBayItemNumber, p.Price as CostcoOldPrice, s.Price as CostcoNewPrice, l.eBayListingPrice as eBayListingOldPrice, p.ImageLink as ImageLink
                                from [dbo].[Staging_ProductInfo] s, [dbo].[ProductInfo] p, [dbo].[eBay_CurrentListings] l
                                where s.UrlNumber = p.UrlNumber
                                and s.UrlNumber = l.CostcoUrlNumber
                                and l.DeleteDT is NULL
                                and s.Price > p.Price";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"select COUNT(1) from [dbo].[eBayListingChange_PriceUp]";
            cmd.CommandText = sqlString;
            nEBayListingChangePriceUp = Convert.ToInt16(cmd.ExecuteScalar());

            //     sqlString = @"INSERT INTO [dbo].[eBay_ToChange] (Name, CostcoUrlNumber, eBayItemNumber, eBayOldListingPrice, 
            //eBayNewListingPrice, eBayReferencePrice, 
            //CostcoOldPrice, CostcoNewPrice, PriceChange)
            //                     SELECT l.Name, l.CostcoUrlNumber, l.eBayItemNumber, l.eBayListingPrice, l.eBayListingPrice, 
            //                     l.eBayReferencePrice, l.CostcoPrice, r.Price, 'up'
            //                     FROM [dbo].[eBay_CurrentListings] l, [dbo].[Staging_ProductInfo] r
            //                     WHERE l.CostcoPrice < r.Price 
            //                     AND l.CostcoUrlNumber = r.UrlNumber";

            //     cmd.CommandText = sqlString;
            //     cmd.ExecuteNonQuery();

            // eBay listing price down
            sqlString = @"INSERT INTO [dbo].[eBayListingChange_PriceDown] (Name, CostcoUrl, UrlNumber, eBayItemNumber, CostcoOldPrice, CostcoNewPrice, eBayListingOldPrice, ImageLink)
                                select p.name as Name, p.Url as CostcoUrl, p.UrlNumber as UrlNumber, l.eBayItemNumber as eBayItemNumber, p.Price as CostcoOldPrice, s.Price as CostcoNewPrice, l.eBayListingPrice as eBayListingOldPrice, p.ImageLink as ImageLink
                                from [dbo].[Staging_ProductInfo] s, [dbo].[ProductInfo] p, [dbo].[eBay_CurrentListings] l
                                where s.UrlNumber = p.UrlNumber
                                and s.UrlNumber = l.CostcoUrlNumber
                                and l.DeleteDT is NULL
                                and s.Price < p.Price";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"select COUNT(1) from [dbo].[eBayListingChange_PriceDown]";
            cmd.CommandText = sqlString;
            nEBayListingChangePriceDown = Convert.ToInt16(cmd.ExecuteScalar());

       //     sqlString = @"INSERT INTO [dbo].[eBay_ToChange] (Name, CostcoUrlNumber, eBayItemNumber, eBayOldListingPrice, 
							//eBayNewListingPrice, eBayReferencePrice, 
							//CostcoOldPrice, CostcoNewPrice, PriceChange)
       //                     SELECT l.Name, l.CostcoUrlNumber, l.eBayItemNumber, l.eBayListingPrice, l.eBayListingPrice, 
       //                     l.eBayReferencePrice, l.CostcoPrice, r.Price, 'down'
       //                     FROM [dbo].[eBay_CurrentListings] l, [dbo].[Staging_ProductInfo] r
       //                     WHERE l.CostcoPrice < r.Price
       //                     AND l.CostcoUrlNumber = r.UrlNumber";

       //     cmd.CommandText = sqlString;
       //     cmd.ExecuteNonQuery();

            // eBay listing discontinused 
            sqlString = @"INSERT INTO [dbo].[eBayListingChange_Discontinue] (Name, CostcoUrl, UrlNumber, eBayItemNumber)
                        SELECT p.name, p.CostcoUrl, p.CostcoUrlNumber, p.CostcoUrlNumber
                        FROM [dbo].[eBay_CurrentListings] p
                        WHERE not exists (SELECT 1 FROM Staging_ProductInfo sp where sp.UrlNumber = p.CostcoUrlNumber)
                        AND p.DeleteDT is NULL";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"select COUNT(1) from [dbo].[eBayListingChange_Discontinue]";
            cmd.CommandText = sqlString;
            nEBayListingChangeDiscontinue = Convert.ToInt16(cmd.ExecuteScalar());

            //rdr = cmd.ExecuteReader();

            //eBayListingDiscontinueddProductArray.Clear();


            //while (rdr.Read())
            //{
            //    eBayListingDiscontinueddProductArray.Add("<a href='" + rdr["CostcoUrl"].ToString().Replace("'", "&#39;") + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["CostcoPrice"].ToString());
            //}

            //rdr.Close();

            //sqlString = @"INSERT INTO [dbo].[eBay_ToRemove] (Name, CostcoUrlNumber, eBayItemNumber)
            //                SELECT l.Name, l.CostcoUrlNumber, l.eBayItemNumber
            //                FROM [dbo].[eBay_CurrentListings] l
            //                WHERE not exists 
	           //             (SELECT 1 FROM [dbo].[Staging_ProductInfo] r where r.UrlNumber = l.CostcoUrlNumber)";

            //cmd.CommandText = sqlString;
            //cmd.ExecuteNonQuery();

            // options change 
            sqlString = @"INSERT INTO [dbo].[eBayListingChange_OptionChange] (Name, CostcoUrl, UrlNumber, eBayItemNumber, CostcoOldOptions, CostcoNewOptions, CostcoNewImageOptions, ImageLink)
                        select s.name as Name, s.Url as CostcoUrl, s.UrlNumber as UrlNumber, l.eBayItemNumber as eBayItemNumber, p.Options as CostcoOldOptions, s.Options as CostcoNewOptions, s.ImageOptions as CostcoNewImageOptions, p.ImageLink as ImageLink
                        from [dbo].[Staging_ProductInfo] s, [dbo].[eBay_CurrentListings] l, [dbo].[ProductInfo] p
                        where s.UrlNumber = l.CostcoUrlNumber
                        and s.UrlNumber = p.UrlNumber
                        and l.DeleteDT is NULL
                        and s.Options <> p.Options";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"select COUNT(1) from [dbo].[eBayListingChange_OptionChange]";
            cmd.CommandText = sqlString;
            nEBayListingChangeOptions = Convert.ToInt16(cmd.ExecuteScalar());

            //rdr = cmd.ExecuteReader();

            //eBayListingDiscontinueddProductArray.Clear();


            //while (rdr.Read())
            //{
            //    eBayListingDiscontinueddProductArray.Add("<a href='" + rdr["CostcoUrl"].ToString().Replace("'", "&#39;") + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["CostcoPrice"].ToString());
            //}

            //rdr.Close();

            cn.Close();
        }

        private void ArchieveProducts()
        {
            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            cn.Open();

            // Archieve
            string sqlString = @"insert into [dbo].[Archieve] (Name, urlNumber, itemnumber, Category, price, shipping, limit, discount, details, specification, imageLink, imageOptions, url, ImportedDT, NumberOfImage, Options)
                                select distinct Name, urlNumber, itemnumber, Category, price, shipping, limit, discount, details, specification, imageLink, imageOptions, url, GETDATE(), NumberOfImage, Options
                                from  [dbo].[ProductInfo]";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"TRUNCATE TABLE ProductInfo";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"insert into [dbo].[ProductInfo] (Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, imageOptions, url, options)
                        select distinct Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, imageOptions, url, options
                        from  dbo.staging_productInfo";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"TRUNCATE TABLE Staging_ProductInfo";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"DELETE FROM Archieve WHERE ImportedDT < DATEADD(day, -30, GETDATE())";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            cn.Close();
        }

        private void SendEmail()
        {
            emailMessage = "<p>Start: " + startDT.ToLongTimeString() + "</p></br>";
            //emailMessage += "<p>Productlist End: " + productListEndDT.ToLongTimeString() + "</p></br>";
            emailMessage += "<p>End: " + endDT.ToLongTimeString() + "</p></br>";

            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<p>Product scanned: " + nScanProducts.ToString() + "</p></br>";
            emailMessage += "<p>Product imported: " + nImportProducts.ToString() + "</p></br>";
            emailMessage += "</br>";
            emailMessage += "</br>";

            //emailMessage += "<p>nCategoryUrlArray: " + nCategoryUrlArray.ToString() + "</p></br>";
            //emailMessage += "<p>nProductListPages: " + nProductListPages.ToString() + "</p></br>";
            //emailMessage += "<p>nProductUrlArray: " + nProductUrlArray.ToString() + "</p></br>";

            //emailMessage += "</br>";
            //emailMessage += "</br>";

            //emailMessage += "<p>Product Scanned: " + nScanProducts.ToString() + "</p></br>";
            //emailMessage += "<p>Product Imported: " + nImportProducts.ToString() + "</p></br>";
            //emailMessage += "<p>Product Skipped: " + nSkipProducts.ToString() + "</p></br>";
            //emailMessage += "<p>Product Errored: " + nImportErrors.ToString() + "</p></br>";

            //emailMessage += "</br>";
            //emailMessage += "</br>";

            //emailMessage += "<h3>First try fix products: (" + firstTryResult.Count.ToString() + ")</h3>" + "</br>";
            //emailMessage += "</br>";

            //foreach (string a in firstTryResult)
            //{
            //    emailMessage += "<p>" + a + "</p></br>";
            //}

            //emailMessage += "</br>";
            //emailMessage += "</br>";

            //emailMessage += "<h3>Second try fix products: (" + secondTryResult.Count.ToString() + ")</h3>" + "</br>";
            //emailMessage += "</br>";

            //foreach (string a in secondTryResult)
            //{
            //    emailMessage += "<p>" + a + "</p></br>";
            //}

            //emailMessage += "</br>";
            //emailMessage += "</br>";

            emailMessage += "<h3>Price up products: (" + priceUpProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>Price down products: (" + priceDownProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>New products: (" + newProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>Discontinued products: (" + discontinueddProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>Stock changed products: (" + stockChangeProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>eBay listing price up products: (" + eBayListingPriceUpProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>eBay listing price down products: (" + eBayListingPriceDownProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>eBay listing discontinued products: (" + eBayListingDiscontinueddProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<h3>Price up products: (" + priceUpProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            if (priceUpProductArray.Count == 0)
                emailMessage += "<p>No price up product</p>" + "</br>";
            else
            {
                foreach (string priceUpProduct in priceUpProductArray)
                {
                    emailMessage += "<p>" + priceUpProduct + "</p></br>";
                }
            }
            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<h3>Price down products: (" + priceDownProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            if (priceDownProductArray.Count == 0)
                emailMessage += "<p>No price down product" + "</p></br>";
            else
            {
                foreach (string priceDownProduct in priceDownProductArray)
                {
                    emailMessage += "<p>" + priceDownProduct + "</p></br>";
                }
            }
            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<h3>New products: (" + newProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            if (newProductArray.Count == 0)
                emailMessage += "<p>No new product</p>" + "</br>";
            else
            {
                foreach (string newProduct in newProductArray)
                {
                    emailMessage += "<p>" + newProduct + "</P></br>";
                }
            }
            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<h3>Discontinued products: (" + discontinueddProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            if (this.discontinueddProductArray.Count == 0)
                emailMessage += "<p>No discontinued product</p>" + "</br>";
            else
            {
                foreach (string discontinueddProduct in discontinueddProductArray)
                {
                    emailMessage += "<p>" + discontinueddProduct + "</P></br>";
                }
            }

            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<h3>Stock changed products: (" + stockChangeProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            if (this.stockChangeProductArray.Count == 0)
                emailMessage += "<p>No stock changed product</p>" + "</br>";
            else
            {
                foreach (string stockChangeProduct in stockChangeProductArray)
                {
                    emailMessage += "<p>" + stockChangeProduct + "</P></br>";
                }
            }

            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<h3>eBay listing price up products: (" + eBayListingPriceUpProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            if (eBayListingPriceUpProductArray.Count == 0)
                emailMessage += "<p>No eBay listing price up product</p>" + "</br>";
            else
            {
                foreach (string priceUpProduct in eBayListingPriceUpProductArray)
                {
                    emailMessage += "<p>" + priceUpProduct + "</p></br>";
                }
            }
            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<h3>eBay listing price down products: (" + eBayListingPriceDownProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            if (eBayListingPriceDownProductArray.Count == 0)
                emailMessage += "<p>No eBay listing price down product</p>" + "</br>";
            else
            {
                foreach (string priceDownProduct in eBayListingPriceDownProductArray)
                {
                    emailMessage += "<p>" + priceDownProduct + "</p></br>";
                }
            }
            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<h3>eBay listing discontinued products: (" + eBayListingDiscontinueddProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            if (eBayListingDiscontinueddProductArray.Count == 0)
                emailMessage += "<p>No eBay listing discontinued product</p>" + "</br>";
            else
            {
                foreach (string priceDiscontinuedProduct in eBayListingDiscontinueddProductArray)
                {
                    emailMessage += "<p>" + priceDiscontinuedProduct + "</p></br>";
                }
            }
            emailMessage += "</br>";
            emailMessage += "</br>";

            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress("zjding@gmail.com");
                mail.To.Add("zjding@gmail.com");
                mail.Subject = DateTime.Now.ToLongDateString();
                mail.Body = emailMessage;
                mail.IsBodyHtml = true;
                //mail.Attachments.Add(new Attachment("C:\\file.zip"));

                using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtp.Credentials = new NetworkCredential("zjding@gmail.com", "G4indigo");
                    smtp.EnableSsl = true;
                    smtp.Send(mail);
                }
            }

        }
    }
}
