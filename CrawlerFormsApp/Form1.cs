using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace CrawlerFormsApp
{
    public partial class Form1 : Form
    {

        //tring HomePage = "https://www.adayroi.com/";

       // string HomePage = "https://tiki.vn/";

        //string HomePage = "https://vinabook.com/";

        string HomePage = "https://www.fahasa.com/";

        HttpClient httpClient;
        HttpClientHandler handler;
        CookieContainer cookie = new CookieContainer();
        public Form1()
        {
            InitializeComponent();
            IniHttpClient();
        }
       

        void IniHttpClient()
        {



            handler = new HttpClientHandler
            {
                CookieContainer = cookie,
                ClientCertificateOptions = ClientCertificateOption.Automatic,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                AllowAutoRedirect = true,
                UseDefaultCredentials = false
            };

            httpClient = new HttpClient(handler);




            httpClient.BaseAddress = new Uri(HomePage);
            
        }


        string CrawlDataFromURL(string url)
        {
            string html = "";

            html = WebUtility.HtmlDecode(httpClient.GetStringAsync(url).Result);

            //html = httpClient.PostAsync(url,new StringContent("")).Result.Content.ReadAsStringAsync().Result;

            return html;
        }
        private void btn_Crawl_Click(object sender, EventArgs e)
        {

            string url = txbUrl.Text;
            string html = CrawlDataFromURL(url);

            string r = Regex.Match(html, @"<div class=""product-list__container"".*?<div id=""addToCartTitle", RegexOptions.Singleline).Value;
            var LinkList = Regex.Matches(r, @"<a class=""product-item__thumbnail""(.*?)/>", RegexOptions.Singleline);

            foreach (var item in LinkList)
            {
                //Link:href="/.*?>

                string Link = Regex.Match(item.ToString(), @"href=""/.*?>", RegexOptions.Singleline).Value;
                Link = Link.Replace("href=", "");
                Link = Link.Replace(@"""", "");
                Link = Link.Replace(@">", "");
                Link = "https://www.adayroi.com" + Link;
                //listView1.Items.Add(new ListViewItem(new string[] { Link }));

                // image link:< img class="hover ".*?alt
                string ImageLink = Regex.Match(item.ToString(), @"data-src=""(.*?)src", RegexOptions.Singleline).Value;
                ImageLink = ImageLink.Replace(@"data-src=", "");
                ImageLink = ImageLink.Replace(@"""", "");
                ImageLink = ImageLink.Replace(@"src", "");


                //Name: title=".*?/>
                string Name = Regex.Match(item.ToString(), @"title="".*?/>").Value;
                Name = Name.Replace("title=", "");
                Name = Name.Replace(@"""", "");
                Name = Name.Replace(@"/>", "");

                string htmlProduct = CrawlDataFromURL(Link);

                ////short description: <div class="short-des__content".*?</div>
                string ShortDescription = Regex.Match(htmlProduct.ToString(), @"<div class=""short-des__content"".*?</div>", RegexOptions.Singleline).Value;
                ShortDescription = ShortDescription.Replace(@"<div class=""short-des__content"" data-role=""content"" data-total=""6""", "");
                ShortDescription = ShortDescription.Replace(@"data-item-height=""27"">", "");
                ShortDescription = ShortDescription.Replace(@"</div>", "");

                //description:<div class="col-sm-12 detail__info">.*?<!-- end detail product -->
                string Description = Regex.Match(htmlProduct.ToString(), @"<div class=""col-sm-12 detail__info"">(.*?)<!-- end detail product -->", RegexOptions.Singleline).Value;
                Description = Description.Replace(@"<div class=""col-sm-12 detail__info"">", "");
                Description = Description.Replace(@"<!-- end detail product -->", "");

                //regular price: <span class="price-info__sale">.*?</span>
                string RegularPrice = Regex.Match(htmlProduct.ToString(), @"<span class=""price-info__sale"">(.*?)</span>", RegexOptions.Singleline).Value;
                RegularPrice = RegularPrice.Replace(@"<span class=""price-info__sale"">", "");
                RegularPrice = RegularPrice.Replace(@".", "");
                RegularPrice = RegularPrice.Replace(@"đ", "");
                RegularPrice = RegularPrice.Replace(@"</span>", "");

                //sale price:<span class="price-vinid__value">.*?</span> 
                string SalePrice = Regex.Match(htmlProduct.ToString(), @"<span class=""price-vinid__value"">.*?</span> ", RegexOptions.Singleline).Value;
                SalePrice = SalePrice.Replace(@"<span class=""price-vinid__value"">", "");
                SalePrice = SalePrice.Replace(@".", "");
                SalePrice = SalePrice.Replace(@"đ", "");
                SalePrice = SalePrice.Replace(@"</span>", "");

                // successfully product
                string ID = "";
                string Type = "external";
                string SKU = "";
                string Published = "1";
                string Isfeatured = "0";
                string Visibility_in_catalog = "visible";
                string Date_sale_price_starts = "";
                string Date_sale_price_ends = "";
                string Tax_status = "taxable";
                string Tax_class = "";
                string In_stock = "";
                string Stock = "";
                string Backorders_allowed = "0";
                string Sold_individually = "0";
                string Weight = "";
                string Length = "";
                string Width = "";
                string Height = "";
                string Allow_customer_reviews = "0";
                string Purchase_note = "";
                string Categories = txbCategory.Text;
                string Tags = "";
                string Shipping_class = "";
                string Download_limit = "";
                string Download_expiry_days = "";
                string Parent = "";
                string Grouped_products = "";
                string Upsells = "";
                string Cross_sells = "";
                string Button_text = txbButton_text.Text;
                string Position = "0";
                listView1.Items.Add(new ListViewItem(new string[] { ID, Type, SKU, Name, Published, Isfeatured, Visibility_in_catalog, ShortDescription, Description, Date_sale_price_starts, Date_sale_price_ends, Tax_status, Tax_class, In_stock, Stock, Backorders_allowed, Sold_individually, Weight, Length, Width, Height, Allow_customer_reviews, Purchase_note, SalePrice, RegularPrice, Categories, Tags, Shipping_class, ImageLink, Download_limit, Download_expiry_days, Parent, Grouped_products, Upsells, Cross_sells, Link, Button_text, Position }));

            }

        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            #region 1
            // Tạo đối tượng lưu tệp
            SaveFileDialog fsave = new SaveFileDialog();

            //Chỉ ra đuôi

            fsave.Filter = "(Tất cả các tệp)|*.*|(Các tệp excel)|*.csv";
            fsave.ShowDialog();

            //Xử lý
            if (fsave.FileName != "")
            {
                //tạo app
                Excel.Application app = new Excel.Application();

                //Tạo workbook
                Excel.Workbook wb = app.Workbooks.Open(fsave.FileName);


                // Tạo sheet

                Excel.Worksheet sheet = null;

                try
                {
                    sheet = wb.ActiveSheet;
                    sheet.Name = "Data import";
                    //Sinh dữ liệu
                    for (int i = 1; i <= listView1.Items.Count; i++)
                    {
                        ListViewItem item = listView1.Items[i - 1];
                        sheet.Cells[i + 1, 1] = item.Text;
                        for (int j = 2; j <= listView1.Columns.Count; j++)
                        {
                            sheet.Cells[i + 1, j] = item.SubItems[j - 1].Text;

                        }

                    }
                    wb.SaveAs(fsave.FileName);
                    MessageBox.Show("Ghi thành công", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    app.Quit();
                    wb = null;
                }

            }
            else
            {

                MessageBox.Show("Bạn không chọn tệp nào", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            #endregion
        }

        private void btnCrawl_Tiki_Click(object sender, EventArgs e)
        {

            string url = txbUrl.Text;
            string html = CrawlDataFromURL(url);

            string r = Regex.Match(html, @"<div class=""product-box-list""(.*?)<div class=""list-pager"">", RegexOptions.Singleline).Value;
            var LinkList = Regex.Matches(r, @"<div data-seller-product-id=(.*?)</div>", RegexOptions.Singleline);

            foreach (var item in LinkList)
            {
                // Link : href=""(.*?)title

                string Link = Regex.Match(item.ToString(), @"href=""(.*?)title", RegexOptions.Singleline).Value;
                Link = Link.Replace("href=", "");
                Link = Link.Replace(@"""", "");
                Link = Link.Replace(@"title", "");



                //Name: title=".*?/>
                string Name = Regex.Match(item.ToString(), @"data-title=""(.*?)data-price").Value;
                Name = Name.Replace("data-title=", "");
                Name = Name.Replace(@"""", "");
                Name = Name.Replace(@"data-price", "");

                // image link:itemprop=""image""(.*?)>
                string ImageLink = Regex.Match(item.ToString(), @"<img class=""product-image img-responsive""(.*?)alt="""">", RegexOptions.Singleline).Value;
                ImageLink = ImageLink.Replace(@"<img class=""product-image img-responsive"" src=", "");
                ImageLink = ImageLink.Replace(@"""", "");
                ImageLink = ImageLink.Replace(@"alt=>", "");


                string htmlProduct = CrawlDataFromURL(Link);

                //SKU: <input id="product_sku"(.*?)>
                string SKU = Regex.Match(htmlProduct.ToString(), @"<input id=""product_sku""(.*?)>", RegexOptions.Singleline).Value;
                SKU = SKU.Replace(@"<input id=""product_sku"" name=""sku"" type=""hidden"" value=", "");
                SKU = SKU.Replace(@"""", "");
                SKU = SKU.Replace(@">", "");


                //short description: <table id="chi-tiet" (.*?)</table>

                //string ShortDescription = Regex.Match(htmlProduct.ToString(), @"<table id=""chi-tiet""(.*?)</table>", RegexOptions.Singleline).Value;



                //description:<div class="product-content-box">(.*?)<div class="right">
                string Description = Regex.Match(htmlProduct.ToString(), @"<div class=""product-content-box"">(.*?)<div class=""right"">", RegexOptions.Singleline).Value;
                Description = Description.Replace(@"<div class=""product-content-box"">", "");
                Description = Description.Replace(@"<div class=""right"">", "");

                //regular price: class="old-price-item"(.*?)>
                string RegularPrice = Regex.Match(htmlProduct.ToString(), @"class=""old-price-item""(.*?)>", RegexOptions.Singleline).Value;
                RegularPrice = RegularPrice.Replace(@"class=""old-price-item"" data-value=", "");
                RegularPrice = RegularPrice.Replace(@"id=""p-listpirce"">", "");
                RegularPrice = RegularPrice.Replace(@"""", "");

                //sale price:<p class="special-price-item"(.*?)>
                string SalePrice = Regex.Match(htmlProduct.ToString(), @"<p class=""special-price-item""(.*?)>", RegexOptions.Singleline).Value;
                SalePrice = SalePrice.Replace(@"<p class=""special-price-item"" data-value=", "");
                SalePrice = SalePrice.Replace(@"id=""p-specialprice"">", "");
                SalePrice = SalePrice.Replace(@"""", "");


                // successfully product
                string ID = "";
                string Type = "external";
                string ShortDescription="";
                string Published = "1";
                string Isfeatured = "0";
                string Visibility_in_catalog = "visible";
                string Date_sale_price_starts = "";
                string Date_sale_price_ends = "";
                string Tax_status = "taxable";
                string Tax_class = "";
                string In_stock = "";
                string Stock = "";
                string Backorders_allowed = "0";
                string Sold_individually = "0";
                string Weight = "";
                string Length = "";
                string Width = "";
                string Height = "";
                string Allow_customer_reviews = "1";
                string Purchase_note = "";
                string Categories = txbCategory.Text;
                string Tags = "";
                string Shipping_class = "";
                string Download_limit = "";
                string Download_expiry_days = "";
                string Parent = "";
                string Grouped_products = "";
                string Upsells = "";
                string Cross_sells = "";
                string Button_text = txbButton_text.Text;
                string Position = "0";
                listView1.Items.Add(new ListViewItem(new string[] { ID, Type, SKU, Name, Published, Isfeatured, Visibility_in_catalog, ShortDescription, Description, Date_sale_price_starts, Date_sale_price_ends, Tax_status, Tax_class, In_stock, Stock, Backorders_allowed, Sold_individually, Weight, Length, Width, Height, Allow_customer_reviews, Purchase_note, SalePrice, RegularPrice, Categories, Tags, Shipping_class, ImageLink, Download_limit, Download_expiry_days, Parent, Grouped_products, Upsells, Cross_sells, Link, Button_text, Position }));

            }

        }

        

        private void btnVinabook_Click(object sender, EventArgs e)
        {
            string url = txbUrl.Text;
            string html = CrawlDataFromURL(url);

            string r = Regex.Match(html, @"<div class=""product_recommend-box product-details-box"">(.*?)<!--category_products-->", RegexOptions.Singleline).Value;
            var LinkList = Regex.Matches(r, @"<div class=""product_thumb""(.*?)</div>", RegexOptions.Singleline);

            foreach (var item in LinkList)
            {
                // Link :href=(.*?)>

                string Link = Regex.Match(item.ToString(), @"href=(.*?)>", RegexOptions.Singleline).Value;
                Link = Link.Replace("href=", "");
                Link = Link.Replace(@"""", "");
                Link = Link.Replace(@">", "");



                

                


                string htmlProduct = CrawlDataFromURL(Link);

                //Name: <h1 class="mainbox-title" itemprop="name">(.*?)</h1>
                string Name = Regex.Match(htmlProduct.ToString(), @"<h1 class=""mainbox-title"" itemprop=""name"">(.*?)</h1>").Value;
                Name = Name.Replace(@"<h1 class=""mainbox-title"" itemprop=""name"">", "");
                Name = Name.Replace(@"</h1>", "");
                
                // image :<div class=""bk-front"">(.*?)</div>
                string image = Regex.Match(htmlProduct.ToString(), @"<div class=""bk-front"">(.*?)</div>", RegexOptions.Singleline).Value;
               // ImageLink: src=(.*?)alt
                string ImageLink = Regex.Match(image.ToString(), @"src=(.*?)alt", RegexOptions.Singleline).Value;
                ImageLink = ImageLink.Replace(@"src=", "");
                ImageLink = ImageLink.Replace(@"""", "");
                ImageLink = ImageLink.Replace(@"alt", "");

                //SKU: <div id="product_detail_recommend_by_category" data-product_id=
                string SKU = Regex.Match(htmlProduct.ToString(), @"<div id=""product_detail_recommend_by_category"" data-product_id=(.*?)>", RegexOptions.Singleline).Value;
                SKU = SKU.Replace(@"<div id=""product_detail_recommend_by_category"" data-product_id=", "");
                SKU = SKU.Replace(@"""", "");
                SKU = SKU.Replace(@">", "");
                SKU = SKU.Trim();
                

                //short description: <table id="chi-tiet" (.*?)</table>

                string ShortDescription = Regex.Match(htmlProduct.ToString(), @"itemprop=""description"">(.*?)<a", RegexOptions.Singleline).Value;
                ShortDescription = ShortDescription.Replace(@"itemprop=""description"">","");
                ShortDescription = ShortDescription.Replace(@"""", "");
                ShortDescription = ShortDescription.Replace(@"<a", "");
                //description:<h3 class="mainbox2-title clearfix margin-top-20">(.*?)<div class="mainbox2-bottom">
                string Description1 = Regex.Match(htmlProduct.ToString(), @"<h3 class=""mainbox2-title clearfix margin-top-20"">(.*?)<div class=""mainbox2-bottom"">", RegexOptions.Singleline).Value;
                Description1 = Description1.Replace(@"<div class=""mainbox2-bottom"">", "");

                string Description2 = Regex.Match(htmlProduct.ToString(), @"<div id=""product-details-box""(.*?)<div class=""product_recommend-box product-details-box  other-people-buy-this"">", RegexOptions.Singleline).Value;
                Description2 = Description2.Replace(@"<div class=""product_recommend-box product-details-box  other-people-buy-this"">", "");
                string Description = Description1 + Description2;
                //regular price: misc1:(.*?)//
                string RegularPrice = Regex.Match(htmlProduct.ToString(), @"misc1:(.*?)//", RegexOptions.Singleline).Value;
                RegularPrice = RegularPrice.Replace(@"misc1:", "");
                RegularPrice = RegularPrice.Replace(@"""", "");
                RegularPrice = RegularPrice.Replace(@",", "");
                RegularPrice = RegularPrice.Replace(@"//", "").Trim();

                //sale price:value: (.*?), currency
                string SalePrice = Regex.Match(htmlProduct.ToString(), @"value: (.*?), currency", RegexOptions.Singleline).Value;
                Regex reg = new Regex("[*'\",_&#^@]");
                SalePrice = reg.Replace(SalePrice, string.Empty);

                Regex reg1 = new Regex("[ ]");
                SalePrice = reg.Replace(SalePrice, "");
                SalePrice = SalePrice.Replace(@"value:", ""); 
               // SalePrice = SalePrice.Replace(@",","");
                SalePrice = SalePrice.Replace(@"currency", "").Trim();
               //SalePrice = SalePrice.Replace(@"[^a-zA-Z0-9_.]+", "").Trim();


                // successfully product
                string ID = "";
                string Type = "external";
               
                string Published = "1";
                string Isfeatured = "0";
                string Visibility_in_catalog = "visible";
                string Date_sale_price_starts = "";
                string Date_sale_price_ends = "";
                string Tax_status = "taxable";
                string Tax_class = "";
                string In_stock = "";
                string Stock = "";
                string Backorders_allowed = "0";
                string Sold_individually = "0";
                string Weight = "";
                string Length = "";
                string Width = "";
                string Height = "";
                string Allow_customer_reviews = "1";
                string Purchase_note = "";
                string Categories = txbCategory.Text;
                string Tags = "";
                string Shipping_class = "";
                string Download_limit = "";
                string Download_expiry_days = "";
                string Parent = "";
                string Grouped_products = "";
                string Upsells = "";
                string Cross_sells = "";
                string Button_text = txbButton_text.Text;
                string Position = "0";
                listView1.Items.Add(new ListViewItem(new string[] { ID, Type, SKU, Name, Published, Isfeatured, Visibility_in_catalog, ShortDescription, Description, Date_sale_price_starts, Date_sale_price_ends, Tax_status, Tax_class, In_stock, Stock, Backorders_allowed, Sold_individually, Weight, Length, Width, Height, Allow_customer_reviews, Purchase_note, SalePrice, RegularPrice, Categories, Tags, Shipping_class, ImageLink, Download_limit, Download_expiry_days, Parent, Grouped_products, Upsells, Cross_sells, Link, Button_text, Position }));

            }
        }

        private void btnCrawl_Fahasa_Click(object sender, EventArgs e)
        {
            string url = txbUrl.Text;
            string html = CrawlDataFromURL(url);

            string r = Regex.Match(html, @"<div class=""product_recommend-box product-details-box"">(.*?)<!--category_products-->", RegexOptions.Singleline).Value;
            var LinkList = Regex.Matches(r, @"<div class=""product_thumb""(.*?)</div>", RegexOptions.Singleline);

            foreach (var item in LinkList)
            {
                // Link :href=(.*?)>

                string Link = Regex.Match(item.ToString(), @"href=(.*?)>", RegexOptions.Singleline).Value;
                Link = Link.Replace("href=", "");
                Link = Link.Replace(@"""", "");
                Link = Link.Replace(@">", "");

                string htmlProduct = CrawlDataFromURL(Link);

                //Name: <h1 class="mainbox-title" itemprop="name">(.*?)</h1>
                string Name = Regex.Match(htmlProduct.ToString(), @"<h1 class=""mainbox-title"" itemprop=""name"">(.*?)</h1>").Value;
                Name = Name.Replace(@"<h1 class=""mainbox-title"" itemprop=""name"">", "");
                Name = Name.Replace(@"</h1>", "");

                // image :<div class=""bk-front"">(.*?)</div>
                string image = Regex.Match(htmlProduct.ToString(), @"<div class=""bk-front"">(.*?)</div>", RegexOptions.Singleline).Value;
                // ImageLink: src=(.*?)alt
                string ImageLink = Regex.Match(image.ToString(), @"src=(.*?)alt", RegexOptions.Singleline).Value;
                ImageLink = ImageLink.Replace(@"src=", "");
                ImageLink = ImageLink.Replace(@"""", "");
                ImageLink = ImageLink.Replace(@"alt", "");

                //SKU: <div id="product_detail_recommend_by_category" data-product_id=
                string SKU = Regex.Match(htmlProduct.ToString(), @"<div id=""product_detail_recommend_by_category"" data-product_id=(.*?)>", RegexOptions.Singleline).Value;
                SKU = SKU.Replace(@"<div id=""product_detail_recommend_by_category"" data-product_id=", "");
                SKU = SKU.Replace(@"""", "");
                SKU = SKU.Replace(@">", "");
                SKU = SKU.Trim();


                //short description: <table id="chi-tiet" (.*?)</table>

                string ShortDescription = Regex.Match(htmlProduct.ToString(), @"itemprop=""description"">(.*?)<a", RegexOptions.Singleline).Value;
                ShortDescription = ShortDescription.Replace(@"itemprop=""description"">", "");
                ShortDescription = ShortDescription.Replace(@"""", "");
                ShortDescription = ShortDescription.Replace(@"<a", "");
                //description:<h3 class="mainbox2-title clearfix margin-top-20">(.*?)<div class="mainbox2-bottom">
                string Description1 = Regex.Match(htmlProduct.ToString(), @"<h3 class=""mainbox2-title clearfix margin-top-20"">(.*?)<div class=""mainbox2-bottom"">", RegexOptions.Singleline).Value;
                Description1 = Description1.Replace(@"<div class=""mainbox2-bottom"">", "");

                string Description2 = Regex.Match(htmlProduct.ToString(), @"<div id=""product-details-box""(.*?)<div class=""product_recommend-box product-details-box  other-people-buy-this"">", RegexOptions.Singleline).Value;
                Description2 = Description2.Replace(@"<div class=""product_recommend-box product-details-box  other-people-buy-this"">", "");
                string Description = Description1 + Description2;
                //regular price: misc1:(.*?)//
                string RegularPrice = Regex.Match(htmlProduct.ToString(), @"misc1:(.*?)//", RegexOptions.Singleline).Value;
                RegularPrice = RegularPrice.Replace(@"misc1:", "");
                RegularPrice = RegularPrice.Replace(@"""", "");
                RegularPrice = RegularPrice.Replace(@",", "");
                RegularPrice = RegularPrice.Replace(@"//", "").Trim();

                //sale price:value: (.*?), currency
                string SalePrice = Regex.Match(htmlProduct.ToString(), @"value: (.*?), currency", RegexOptions.Singleline).Value;
                Regex reg = new Regex("[*'\",_&#^@]");
                SalePrice = reg.Replace(SalePrice, string.Empty);

                Regex reg1 = new Regex("[ ]");
                SalePrice = reg.Replace(SalePrice, "");
                SalePrice = SalePrice.Replace(@"value:", "");
                // SalePrice = SalePrice.Replace(@",","");
                SalePrice = SalePrice.Replace(@"currency", "").Trim();
                //SalePrice = SalePrice.Replace(@"[^a-zA-Z0-9_.]+", "").Trim();



                // successfully product
                string ID = "";
                string Type = "external";

                string Published = "1";
                string Isfeatured = "0";
                string Visibility_in_catalog = "visible";
                string Date_sale_price_starts = "";
                string Date_sale_price_ends = "";
                string Tax_status = "taxable";
                string Tax_class = "";
                string In_stock = "";
                string Stock = "";
                string Backorders_allowed = "0";
                string Sold_individually = "0";
                string Weight = "";
                string Length = "";
                string Width = "";
                string Height = "";
                string Allow_customer_reviews = "1";
                string Purchase_note = "";
                string Categories = txbCategory.Text;
                string Tags = "";
                string Shipping_class = "";
                string Download_limit = "";
                string Download_expiry_days = "";
                string Parent = "";
                string Grouped_products = "";
                string Upsells = "";
                string Cross_sells = "";
                string Button_text = txbButton_text.Text;
                string Position = "0";
                listView1.Items.Add(new ListViewItem(new string[] { ID, Type, SKU, Name, Published, Isfeatured, Visibility_in_catalog, ShortDescription, Description, Date_sale_price_starts, Date_sale_price_ends, Tax_status, Tax_class, In_stock, Stock, Backorders_allowed, Sold_individually, Weight, Length, Width, Height, Allow_customer_reviews, Purchase_note, SalePrice, RegularPrice, Categories, Tags, Shipping_class, ImageLink, Download_limit, Download_expiry_days, Parent, Grouped_products, Upsells, Cross_sells, Link, Button_text, Position }));

            }

        }

        

       
    }
        
    
}
