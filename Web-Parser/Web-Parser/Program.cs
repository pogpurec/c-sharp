using System;
using HtmlAgilityPack;
using System.Net.Http;
using System.Xml;
using System.Text.RegularExpressions;

namespace WebParser
{
    class Program
    {
        static void Main(string[] args)
        {
            // Отправляем GET запрос на указанную страницу магазина
            Console.Write("\n Введите категорию для сбора с сайта https://hi-tech.md/ \n");
            string url = Console.ReadLine();
            //string url = "https://hi-tech.md/televizory-i-elektronika/televizory/";
            var httpClient = new HttpClient();
            var html = httpClient.GetStringAsync(url).Result;
            var htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);
            var page = 1;
            bool lastPage = false;
            var startUrl = url;

            // Добавляем метод ClearFromHtml для очищения текста от HTML разметки
            static string ClearFromHtml(string value)
            {
                if (value.Length > 0) {
                    value = Regex.Replace(value, @"<[^>]+>|&nbsp;|&#xE86C;", "").Trim();
                }
                return value;
            }

            while (lastPage == false) {
                // Получаем название категории и следующую страницу, если она есть
                var mainCategoryElement = htmlDocument.DocumentNode.SelectSingleNode("//h1[@class='ty-mainbox-title']");
                var mainCategory = mainCategoryElement.InnerText.Trim();
                Console.WriteLine($"Название категории: {mainCategory}");

                // Получаем массив товаров на странице и их количество
                var productListElement = htmlDocument.DocumentNode.SelectNodes("//div[@class='ty-grid-list__item ty-quick-view-button__wrapper']");
                var productCount = productListElement.Count;
                Console.WriteLine($"Собираем {page} страницу, всего товаров на странице: {productCount}");
                page = page + 1;

                // Проходим по всем товарам на странице и получаем основные поля для каждого товара
                for (int i = 0; i < productCount; i++) {
                    var producName = productListElement[i].SelectSingleNode(".//div[@class='ty-grid-list__item-name']").InnerText;
                    var productId = productListElement[i].SelectSingleNode(".//span[@class='ty-control-group__item']").InnerText;
                    var productPrice = productListElement[i].SelectSingleNode(".//span[@class='ty-price']").InnerText;
                    var productOldPrice = (productListElement[i].SelectSingleNode(".//span[@class='ty-strike']") == null) ? "" : "/" + productListElement[i].SelectSingleNode(".//span[@class='ty-strike']").InnerText;
                    var productStock = "В наличии";
                    if (productListElement[i].SelectSingleNode(".//span[@class='ty-qty-in-stock preorder ty-control-group__item']") != null) {
                        productStock = "Предзаказ";
                    }
                    Console.WriteLine($"Товар: {ClearFromHtml(producName)} {ClearFromHtml(productId)}; Цена: {ClearFromHtml(productPrice)} {ClearFromHtml(productOldPrice)}, {ClearFromHtml(productStock)}");
                }

                // Загрузка следующей страницы, если на текущей нет идентификатора последней страницы
                if (htmlDocument.DocumentNode.SelectSingleNode("//div[@class='more-products-link product_list_page0']") == null) {
                    url = startUrl + "page-" + page + "/";
                    httpClient = new HttpClient();
                    html = httpClient.GetStringAsync(url).Result;
                    htmlDocument = new HtmlDocument();
                    htmlDocument.LoadHtml(html);
                } else lastPage = true;
            } 
        }
    }
}