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
        string connectionString = "Server=tcp:zjding.database.windows.net,1433;Initial Catalog=Costco;Persist Security Info=False;User ID=zjding;Password=G4indigo;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";

        ScrapingBrowser Browser = new ScrapingBrowser();
        IWebDriver driver;

        List<string> categoryArray = new List<string>();
        List<string> subCategoryArray = new List<string>();
        List<string> productUrlArray = new List<string>();

        List<String> categoryUrlArray = new List<string>();
        List<String> subCategoryUrlArray = new List<string>();
        List<string> productListPages = new List<string>();

        List<String> newProductArray = new List<string>();
        List<String> discontinueddProductArray = new List<string>();
        List<String> priceUpProductArray = new List<string>();
        List<String> priceDownProductArray = new List<string>();

        List<String> eBayListingDiscontinueddProductArray = new List<string>();
        List<String> eBayListingPriceUpProductArray = new List<string>();
        List<String> eBayListingPriceDownProductArray = new List<string>();


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
            runCrawl();
        }

        public void runCrawl()
        {
            startDT = DateTime.Now;

            GetDepartmentArray();

            GetProductUrls_New();

            GetProductInfo();

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

        private void GetProductInfo(bool bTruncate = true)
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

                sqlString = "TRUNCATE TABLE Costco_Categories";
                cmd.CommandText = sqlString;
                cmd.ExecuteNonQuery();

                sqlString = "TRUNCATE TABLE Import_Errors";
                cmd.CommandText = sqlString;
                cmd.ExecuteNonQuery();
            }

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

                    string productUrl = HttpUtility.HtmlDecode(pu);
                    productUrl = productUrl.Replace("%2c", ",");

                    string UrlNum = productUrl.Substring(0, productUrl.LastIndexOf('.'));
                    UrlNum = UrlNum.Substring(UrlNum.LastIndexOf('.') + 1);

                    PageResult = Browser.NavigateToPage(new Uri(productUrl));

                    HtmlNode html = PageResult.Html;

                    if (html.InnerText.Contains("Product Not Found"))
                    {
                        sqlString = "INSERT INTO Import_Skips (Url, SkipPoint) VALUES ('" + pu.Replace(@"'", @"''") + "','" + "Product not found" + "')";
                        cmd.CommandText = sqlString;
                        cmd.ExecuteNonQuery();
                        continue;
                    }

                    string stSubCategories = "";

                    HtmlNode category = html.SelectSingleNode("//ul[@itemprop='breadcrumb']");

                    List<HtmlNode> subCategories = category.SelectNodes("li").ToList();

                    int iCategory = 1;
                    string columns = "";
                    string values = "";
                    foreach (HtmlNode subCategory in subCategories)
                    {
                        string temp = subCategory.InnerText.Replace("\n", "");
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

                    HtmlNode productInfo = html.CssSelect(".product-info").ToList<HtmlNode>().First();

                    List<HtmlNode> topReviewPanelNode = productInfo.CssSelect(".top_review_panel").ToList<HtmlNode>();

                    string discount = "";

                    HtmlNode discountNote = topReviewPanelNode[0].SelectSingleNode("//p[@class='merchandisingText']");

                    if (discountNote != null)
                    {
                        discount = discountNote.InnerText.Replace("?", "");
                        discount = discountNote.InnerText.Replace("'", "");
                    }

                    string productName = ((topReviewPanelNode[0]).SelectNodes("h1"))[0].InnerText;
                    productName = productName.Replace("???", "");
                    productName = productName.Replace("??", "");
                    productName = productName.Trim();

                    List<HtmlNode> col1Node = productInfo.CssSelect(".col1").ToList<HtmlNode>();
                    string itemNumber = (col1Node[0].SelectNodes("p")[0]).InnerText;
                    if (itemNumber.ToUpper().Contains("ITEM") && itemNumber.Length > 6)
                        itemNumber = itemNumber.Substring(6);
                    else
                        itemNumber = "";

                    discountNote = col1Node[0].CssSelect(".merchandisingText").FirstOrDefault();

                    if (discountNote != null)
                    {
                        discount = discount.Length == 0 ? discountNote.InnerText.Replace("?", "") : discount + "; " + discountNote.InnerText.Replace("?", "");
                        discount = discount.Replace("?", "");
                        discount = discount.Replace("'", "");
                    }

                    string price;
                    List<HtmlNode> yourPriceNode = col1Node.CssSelect(".your-price").ToList<HtmlNode>();
                    if (yourPriceNode.Count > 0)
                    {
                        List<HtmlNode> priceNode = yourPriceNode[0].CssSelect(".currency").ToList<HtmlNode>();
                        price = priceNode[0].InnerText;
                        price = price.Replace("$", "");
                        price = price.Replace(",", "");

                        if (price == "- -")
                            price = "-2";
                    }
                    else
                    {
                        price = "-1";
                    }

                    var productOptionsNode = col1Node.CssSelect(".product-option").FirstOrDefault();

                    string shipping = "0";

                    var productSHNode = col1Node[0].SelectSingleNode("//li[@class='product']");

                    if (productSHNode != null)
                    {
                        if (productSHNode.InnerText.ToUpper().Contains("OPTIONS"))
                        {
                            sqlString = "INSERT INTO Import_Skips (Url, SkipPoint) VALUES ('" + pu.Replace(@"'", @"''") + "','" + "Options" + "')";
                            cmd.CommandText = sqlString;
                            cmd.ExecuteNonQuery();
                            continue;
                        }
                        else if (productSHNode.InnerText.ToUpper().Contains("INCLUDED") || productSHNode.InnerText.ToUpper().Contains("INLCUDED"))
                        {
                            shipping = "0";
                        }
                        else
                        {
                            string shString = productSHNode.InnerText;
                            int nDollar = shString.IndexOf("$");
                            if (nDollar > 0)
                            {
                                shString = shString.Substring(nDollar + 1);
                                int nStar = shString.IndexOf("*");
                                if (nStar == -1)
                                    nStar = shString.IndexOf(" ");
                                shString = shString.Substring(0, nStar);
                                shString = shString.Replace(" ", "");
                                shipping = shString;
                            }
                            else
                            {
                                int nShipping = shString.IndexOf("Shipping");
                                int nQuantity = shString.ToUpper().IndexOf("QUANTITY");

                                if (nShipping == -1 || nQuantity == -1)
                                {
                                    sqlString = "INSERT INTO Import_Skips (Url, SkipPoint) VALUES ('" + pu.Replace(@"'", @"''") + "','" + "Shipping and Quantity" + "')";
                                    cmd.CommandText = sqlString;
                                    cmd.ExecuteNonQuery();
                                    continue;
                                }

                                shString = shString.Substring(nShipping, nQuantity);
                                Char[] strarr = shString.ToCharArray().Where(c => Char.IsDigit(c) || c.Equals('.')).ToArray();
                                decimal number = Convert.ToDecimal(new string(strarr));
                                shipping = number.ToString();
                            }
                        }
                    }

                    HtmlNode imageColumnNode = html.CssSelect(".image-column").ToList<HtmlNode>().First();

                    HtmlNode imageNode = imageColumnNode.SelectSingleNode("//img[@itemprop='image']");

                    string imageUrl = (imageNode.Attributes["src"]).Value;

                    sqlString = "INSERT INTO Raw_ProductInfo (Name, UrlNumber, ItemNumber, Category, Price, Shipping, Discount,  ImageLink, Url) VALUES ('" + productName.Replace("'", "''") + "','" + UrlNum + "','" + itemNumber + "','" + stSubCategories + "'," + price + "," + shipping + "," + "'" + discount + "','" + imageUrl.Replace("'", "''") + "','" + productUrl.Replace("'", "''") + "')";
                    cmd.CommandText = sqlString;
                    cmd.ExecuteNonQuery();
                    nImportProducts++;

                    sqlString = "INSERT INTO Costco_Categories (" + columns + ") VALUES (" + values + ")";
                    cmd.CommandText = sqlString;
                    cmd.ExecuteNonQuery();
                }
                catch (Exception exception)
                {
                    string productUrl = HttpUtility.HtmlDecode(pu);
                    productUrl = productUrl.Replace("%2c", ",");
                    productUrl = productUrl.Replace(@"'", @"''");
                    sqlString = "INSERT INTO Import_Errors (Url, Exception) VALUES ('" + productUrl + "','" + exception.Message.Replace(@"'", @"''") + "')";
                    cmd.CommandText = sqlString;
                    cmd.ExecuteNonQuery();

                    continue;
                }
            }

            cn.Close();



            //driver.Dispose();

            //MessageBox.Show("Start: " + startDT.ToLongTimeString() + "; End: " + endDT.ToLongTimeString());
        }

        private void SecondTry(int i = 0)
        {
            productUrlArray.Clear();

            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            cn.Open();

            string sqlString = @"select * from ProductInfo p 
                        where 
                        not exists
                        (select 1 from Raw_ProductInfo sp where sp.UrlNumber = p.UrlNumber)";

            //string sqlString = @"select Url from Import_Skips 
            //            where SkipPoint = 'Product not found'";

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
                //}
            }

            rdr.Close();

            //sqlString = @"select Url from Import_Errors";

            //cmd.CommandText = sqlString;
            //rdr = cmd.ExecuteReader();

            //while (rdr.Read())
            //{
            //    productUrlArray.Add(rdr["Url"].ToString());
            //}

            //rdr.Close();

            //sqlString = @"delete from Import_Skips 
            //            where SkipPoint = 'Product not found'";

            //cmd.CommandText = sqlString;
            //cmd.ExecuteNonQuery();

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

            sqlString = @"insert into dbo.staging_productInfo (Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url)
                        select distinct Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url
                        from dbo.Raw_ProductInfo
                        where Price > 0
                        order by UrlNumber";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            // copy to staging_productInfo_filtered
            sqlString = "TRUNCATE TABLE Staging_ProductInfo_Filtered";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"insert into dbo.Staging_ProductInfo_Filtered(Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url)
                        select distinct Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url 
                        from dbo.Raw_ProductInfo 
                        where Price > 0 and Price < 100 and Shipping = 0
                        order by UrlNumber";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();


        }

        private void CompareProducts()
        {
            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            SqlDataReader rdr;

            cn.Open();

            // price up
            string sqlString = @"select s.Name, s.Price as newPrice, p.Price as oldPrice, s.Url 
                                from [dbo].[Staging_ProductInfo] s, [dbo].[ProductInfo] p
                                where s.UrlNumber = p.UrlNumber
                                and s.Price > p.Price";
            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            priceUpProductArray.Clear();

            while (rdr.Read())
            {
                priceUpProductArray.Add("<a href='" + rdr["Url"].ToString() + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["newPrice"].ToString() + "|(" + rdr["oldPrice"].ToString() + ")");
            }

            rdr.Close();

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
                priceDownProductArray.Add("<a href='" + rdr["Url"].ToString() + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["newPrice"].ToString() + "|(" + rdr["oldPrice"].ToString() + ")");
            }

            rdr.Close();

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
                newProductArray.Add("<a href='" + rdr["Url"].ToString() + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["Price"].ToString());
            }

            rdr.Close();

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
                discontinueddProductArray.Add("<a href='" + rdr["Url"].ToString() + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["Price"].ToString());
            }

            rdr.Close();

            // eBay listing price up
            sqlString = @"select s.Name, s.CostcoPrice as OldBasePrice, s.eBayListingPrice as eBayListingPrice, p.Price as NewBasePrice, p.Url as CostcoUrl, s.eBayItemNumber as eBayItemNumber
                            from [dbo].[eBay_CurrentListings] s, [dbo].[Staging_ProductInfo] p
                            where s.CostcoUrlNumber = p.UrlNumber
                            and s.CostcoPrice < p.Price";
            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            eBayListingPriceUpProductArray.Clear();

            while (rdr.Read())
            {
                eBayListingPriceUpProductArray.Add("<a href='" + rdr["CostcoUrl"].ToString() + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["NewBasePrice"].ToString() + "|(" + rdr["OldBasePrice"].ToString() + ")");
            }

            rdr.Close();

            sqlString = @"INSERT INTO [dbo].[eBay_ToChange] (Name, CostcoUrlNumber, eBayItemNumber, eBayOldListingPrice, 
							eBayNewListingPrice, eBayReferencePrice, 
							CostcoOldPrice, CostcoNewPrice, PriceChange)
                            SELECT l.Name, l.CostcoUrlNumber, l.eBayItemNumber, l.eBayListingPrice, l.eBayListingPrice, 
                            l.eBayReferencePrice, l.CostcoPrice, r.Price, 'up'
                            FROM [dbo].[eBay_CurrentListings] l, [dbo].[Staging_ProductInfo] r
                            WHERE l.CostcoPrice < r.Price 
                            AND l.CostcoUrlNumber = r.UrlNumber";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            // eBay listing price down
            sqlString = @"select s.Name, s.CostcoPrice as OldBasePrice, s.eBayListingPrice as eBayListingPrice, p.Price as NewBasePrice, p.Url as CostcoUrl, s.eBayItemNumber as eBayItemNumber
                            from [dbo].[eBay_CurrentListings] s, [dbo].[Staging_ProductInfo] p
                            where s.CostcoUrlNumber = p.UrlNumber
                            and s.CostcoPrice > p.Price";
            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            eBayListingPriceDownProductArray.Clear();

            while (rdr.Read())
            {
                eBayListingPriceDownProductArray.Add("<a href='" + rdr["CostcoUrl"].ToString() + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["NewBasePrice"].ToString() + "|(" + rdr["OldBasePrice"].ToString() + ")");
            }

            rdr.Close();

            sqlString = @"INSERT INTO [dbo].[eBay_ToChange] (Name, CostcoUrlNumber, eBayItemNumber, eBayOldListingPrice, 
							eBayNewListingPrice, eBayReferencePrice, 
							CostcoOldPrice, CostcoNewPrice, PriceChange)
                            SELECT l.Name, l.CostcoUrlNumber, l.eBayItemNumber, l.eBayListingPrice, l.eBayListingPrice, 
                            l.eBayReferencePrice, l.CostcoPrice, r.Price, 'down'
                            FROM [dbo].[eBay_CurrentListings] l, [dbo].[Staging_ProductInfo] r
                            WHERE l.CostcoPrice < r.Price
                            AND l.CostcoUrlNumber = r.UrlNumber";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            // eBay listing discontinused 
            sqlString = @"select * from eBay_CurrentListings p 
                        where 
                        not exists
                        (select 1 from Staging_ProductInfo sp where sp.UrlNumber = p.CostcoUrlNumber)";
            cmd.CommandText = sqlString;
            rdr = cmd.ExecuteReader();

            eBayListingDiscontinueddProductArray.Clear();


            while (rdr.Read())
            {
                eBayListingDiscontinueddProductArray.Add("<a href='" + rdr["CostcoUrl"].ToString() + "'>" + rdr["Name"].ToString() + "</a>|" + rdr["CostcoPrice"].ToString());
            }

            rdr.Close();

            sqlString = @"INSERT INTO [dbo].[eBay_ToRemove] (Name, CostcoUrlNumber, eBayItemNumber)
                            SELECT l.Name, l.CostcoUrlNumber, l.eBayItemNumber
                            FROM [dbo].[eBay_CurrentListings] l
                            WHERE not exists 
	                        (SELECT 1 FROM [dbo].[Staging_ProductInfo] r where r.UrlNumber = l.CostcoUrlNumber)";

            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            cn.Close();
        }

        private void ArchieveProducts()
        {
            SqlConnection cn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;

            cn.Open();

            // price up
            string sqlString = @"insert into [dbo].[Archieve] (Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url, ImportedDT)
                                select distinct Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url, GETDATE()
                                from  [dbo].[ProductInfo]";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = "TRUNCATE TABLE ProductInfo";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = @"insert into [dbo].[ProductInfo] (Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url)
                        select distinct Name, urlNumber, itemnumber, Category, price, shipping, discount, details, specification, imageLink, url
                        from  dbo.staging_productInfo";
            cmd.CommandText = sqlString;
            cmd.ExecuteNonQuery();

            sqlString = "TRUNCATE TABLE Staging_ProductInfo";
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

            emailMessage += "<p>Product scanned: " + productUrlArray.Count.ToString() + "</p></br>";
            emailMessage += "<p>Product imported: " + nImportProducts.ToString() + "</p></br>";
            emailMessage += "</br>";
            emailMessage += "</br>";

            emailMessage += "<h3>Price up products: (" + priceUpProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>Price down products: (" + priceDownProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>New products: (" + newProductArray.Count.ToString() + ")</h3>" + "</br>";
            emailMessage += "</br>";
            emailMessage += "<h3>Discontinued products: (" + discontinueddProductArray.Count.ToString() + ")</h3>" + "</br>";
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
                emailMessage += "<p>No new product</p>" + "</br>";
            else
            {
                foreach (string discontinueddProduct in discontinueddProductArray)
                {
                    emailMessage += "<p>" + discontinueddProduct + "</P></br>";
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
                    smtp.Credentials = new NetworkCredential("zjding@gmail.com", "yueding00");
                    smtp.EnableSsl = true;
                    smtp.Send(mail);
                }
            }

        }
    }
}
