using System.Text;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using ShapeCrawler;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;

// ФАК http://api.test.lk.sidorinlab.ru/Help

class Program
{
    static async Task Main()
    {
        string apiUrlToken = "http://api.test.lk.sidorinlab.ru/api/token";
        string apiUrlClient = "http://api.test.lk.sidorinlab.ru/api/client";
        string apiUrlProject = "http://api.test.lk.sidorinlab.ru/api/client/project";
        string apiUrlSearchResult = "http://api.test.lk.sidorinlab.ru/api/client/searchresult";
        string username = "apitest@mail.ru";
        string password = "aSEiM?^mTn";
        string clientName = "Аякс Краснодар"; // сюда вводим название клиента
        string projectName = "Аякс "; // сюда вводим название проекта
        string clientId = "";
        string projectId = "";

        // Получаем токен
        string token = await GetToken(apiUrlToken, username, password);
        Console.WriteLine($"Токен: {token}");

        // Получаем список клиентов и ищем нужного клиента по имени:
        string clientsData = await GetClients(apiUrlClient, token);
        JToken[] clientJsonArray = JArray.Parse(clientsData).ToArray();
        foreach (JToken obj in clientJsonArray)
            {   
                if (obj["Name"]?.ToString() == clientName)
                    {
                        clientId = obj["Id"]?.ToString();
                        Console.WriteLine($"Id клиента: {clientId}");
                        break; // Найден нужный Id, прерываем цикл
                    }
            }

        // Получаем список проектов клиента и ищем нужный проект по имени:
        string projectsData = await GetProjects(apiUrlProject, token, clientId);
        JToken[] projectsJsonArray = JArray.Parse(projectsData).ToArray();
        foreach (JToken obj in projectsJsonArray)
            {   
                if (obj["Name"]?.ToString() == projectName)
                    {
                        projectId = obj["Id"]?.ToString();
                        Console.WriteLine($"Id проекта: {projectId}");
                        break; // Найден нужный Id, прерываем цикл
                    }
            }

        // Получаем и выводим данные всех выдач и сохраняем негативныые ссылки
        string searchResultData = await GetSearchResult(apiUrlSearchResult, token, projectId);
        JToken[] searchResultJsonArray = JArray.Parse(searchResultData).ToArray();
        var firstSearchResultNegativeItemsYandex = searchResultJsonArray[0]["YandexSearchItems"].Where(item => item["Tonality"]["Name"]?.ToString() == "Негатив").ToArray();
        var firstSearchResultNegativeItemsGoogle = searchResultJsonArray[0]["GoogleSearchItems"].Where(item => item["Tonality"]["Name"]?.ToString() == "Негатив").ToArray();
        var firstSearchResultNegativeItems = firstSearchResultNegativeItemsYandex.Concat(firstSearchResultNegativeItemsGoogle).ToArray(); // Объединили негативные страницы Я + G из первой выдачи
        var secondSearchResultNegativeItemsYandex = searchResultJsonArray[1]["YandexSearchItems"].Where(item => item["Tonality"]["Name"]?.ToString() == "Негатив").ToArray();
        var secondSearchResultNegativeItemsGoogle = searchResultJsonArray[1]["GoogleSearchItems"].Where(item => item["Tonality"]["Name"]?.ToString() == "Негатив").ToArray();
        var secondSearchResultNegativeItems = secondSearchResultNegativeItemsYandex.Concat(secondSearchResultNegativeItemsGoogle).ToArray(); // Объединили негативные страницы Я + G из второй выдачи

        var uniqueFirstSearchResultNegativeItems = firstSearchResultNegativeItems.GroupBy(item => item["Url"]).Select(group => new
            {
                Url = group.Key,
                SiteType = group.First()["SiteType"],
                SearchResultCount = group.Count(),
            }).ToArray();
        var uniqueSecondSearchResultNegativeItems = secondSearchResultNegativeItems.GroupBy(item => item["Url"]).Select(group => new
            {
                Url = group.Key,
                SiteType = group.First()["SiteType"],
                SearchResultCount = group.Count(),
            }).ToArray();

        // Проходим по обоим массивам с негативными ссылками из 2х разных выдач и формируем один с уникальными общими ссылками и повторами из разных выдач
        // Создаем пустой список
        var finalNegative = new List<object>();
        int negativeCountfirstSearch = 0;
        int negativeCountSecondSearch = 0;
        // Цикл добавления негативных страниц в список из старой выдачи
        foreach (var firstSearchNegative in uniqueFirstSearchResultNegativeItems)
        {
            string Url = firstSearchNegative.Url.ToString();
            string SiteType = firstSearchNegative.SiteType.ToString();
            string FirstSearchResultCount = firstSearchNegative.SearchResultCount.ToString();
            string SecondSearchResultCount = "0";
            string Target = "Вытеснение";
            negativeCountfirstSearch += firstSearchNegative.SearchResultCount;
            if (SiteType == "Отзовик (HR)" || SiteType == "Отзовик (product)") {
                Target = "Смена тона";
            }
            foreach (var secondSearchNegative in uniqueSecondSearchResultNegativeItems)
            {   
                if (Url.Equals(secondSearchNegative.Url)) {
                    SecondSearchResultCount = secondSearchNegative.SearchResultCount.ToString();
                    negativeCountSecondSearch += secondSearchNegative.SearchResultCount;
                }
            }
            finalNegative.Add(new
                {
                    Url = Url,
                    SiteType = SiteType,
                    FirstSearchResultCount = FirstSearchResultCount,
                    SecondSearchResultCount = SecondSearchResultCount,
                    Target = Target
                });
        }
        // Цикл добавления негативных страниц в список из новой выдачи
        foreach (var secondSearchNegative in uniqueSecondSearchResultNegativeItems)
        {
            string Url = secondSearchNegative.Url.ToString();
            string SiteType = secondSearchNegative.SiteType.ToString();
            string FirstSearchResultCount = "0";
            string SecondSearchResultCount = secondSearchNegative.SearchResultCount.ToString();;
            string Target = "Вытеснение";
            if (SiteType == "Отзовик (HR)" || SiteType == "Отзовик (product)") {
                Target = "Смена тона";
            }
            bool duplicate = false;
            foreach (var firstSearchNegative in uniqueFirstSearchResultNegativeItems)
            {   
                if (Url.Equals(firstSearchNegative.Url)) {
                    duplicate = true;
                }
            }
            if (duplicate == false) {
                finalNegative.Add(new
                {
                    Url = Url,
                    SiteType = SiteType,
                    FirstSearchResultCount = FirstSearchResultCount,
                    SecondSearchResultCount = SecondSearchResultCount,
                    Target = Target
                });
                negativeCountSecondSearch += secondSearchNegative.SearchResultCount;
            }
        }

        // Создаем новую презентацию
        var pres = SCPresentation.Create();
        var shapeCollection = pres.Slides[0].Shapes;

        // Добавляем текст
        var addedShape = shapeCollection.AddRectangle(x: 25, y: 25, w: 350, h: 100);
        addedShape.TextFrame!.Text = $"На данный момент в выдаче динамика уникальных негативных ресурсов = {uniqueSecondSearchResultNegativeItems.Length - uniqueFirstSearchResultNegativeItems.Length}";
        var addedShape2 = shapeCollection.AddRectangle(x: 400, y: 25, w: 350, h: 100);
        addedShape2.TextFrame!.Text = $"Динамика общего количества повторов негативных ресурсов = {negativeCountSecondSearch - negativeCountfirstSearch}";
        var addedShape3 = shapeCollection.AddRectangle(x: 775, y: 25, w: 450, h: 100);
        addedShape3.TextFrame!.Text = "Для сокращения доли негатива в выдаче топ-10 мы продолжим работу по смене тональности страниц.";

        var table = shapeCollection.AddTable(25, 150, 6, finalNegative.Count+2);

        // Задаем ширину столбцов и высоту ячеек
        table.Columns[0].Width = 40;
        table.Columns[1].Width = 550;
        table.Columns[2].Width = table.Columns[3].Width = table.Columns[4].Width = table.Columns[5].Width = 150;

        // Заполняем ячейки таблицы
        table[0, 0].TextFrame.Text = "№";
        table[0, 1].TextFrame.Text = "URL";
        table[0, 2].TextFrame.Text = "Тип сайта";
        table[0, 3].TextFrame.Text = "Пов. раньше";
        table[0, 4].TextFrame.Text = "Пов. сейчас";
        table[0, 5].TextFrame.Text = "Цель";
        table[finalNegative.Count, 2].TextFrame.Text = "Всего:";
        table[finalNegative.Count, 3].TextFrame.Text = $"{negativeCountfirstSearch}";
        table[finalNegative.Count, 4].TextFrame.Text = $"{negativeCountSecondSearch}";
        for (int rowIndex = 1; rowIndex < finalNegative.Count; rowIndex++)
        {
            for (int colIndex = 0; colIndex <=5 ; colIndex++)
            {
                var cell = table[rowIndex, colIndex];
                if (colIndex == 0) { cell.TextFrame.Text = $"{rowIndex}"; // 1я колонка - указываем номер строки
                } else if (colIndex == 1) { cell.TextFrame.Text = $"{finalNegative[rowIndex-1].GetType().GetProperty("Url").GetValue(finalNegative[rowIndex-1])}"; //2я колонка - указываем урл
                } else if (colIndex == 2) { cell.TextFrame.Text = $"{finalNegative[rowIndex-1].GetType().GetProperty("SiteType").GetValue(finalNegative[rowIndex-1])}"; //2я колонка - указываем тип сайта
                } else if (colIndex == 3) { cell.TextFrame.Text = $"{finalNegative[rowIndex-1].GetType().GetProperty("FirstSearchResultCount").GetValue(finalNegative[rowIndex-1])}"; //2я колонка - указываем повторы в прошлой выдаче
                } else if (colIndex == 4) { cell.TextFrame.Text = $"{finalNegative[rowIndex-1].GetType().GetProperty("SecondSearchResultCount").GetValue(finalNegative[rowIndex-1])}"; //2я колонка - указываем повторы в текущей выдаче
                } else if (colIndex == 5) { cell.TextFrame.Text = $"{finalNegative[rowIndex-1].GetType().GetProperty("Target").GetValue(finalNegative[rowIndex-1])}"; //2я колонка - указываем цель
                }
            }
        }

        pres.SaveAs("my_pres.pptx");

        // Вывод результата
        //Console.WriteLine("Список с элементами:");
        //foreach (var result in finalNegative)
        //{
        //    Console.WriteLine($"Url: {result.GetType().GetProperty("Url").GetValue(result)}, Тип сайта: {result.GetType().GetProperty("SiteType").GetValue(result)}, Повторы в прошлом мес.: {result.GetType().GetProperty("FirstSearchResultCount").GetValue(result)}, Повторы в текущем мес.: {result.GetType().GetProperty("SecondSearchResultCount").GetValue(result)}, Цель: {result.GetType().GetProperty("Target").GetValue(result)}");
        //}
        //Console.WriteLine($"негатива в 1й выдаче: {firstSearchResultNegativeItems.Length}, из них уникальных {uniqueFirstSearchResultNegativeItems.Length}");
        //Console.WriteLine($"негатива во 2й выдаче: {secondSearchResultNegativeItems.Length} из них уникальных {uniqueSecondSearchResultNegativeItems.Length}"); 
    }

    // Получение токена по апи POST запросом
    static async Task<string> GetToken(string apiUrl, string username, string password)
    {
        using (HttpClient client = new HttpClient())
        {
            var requestData = new
            {
                grant_type = "password",
                Email = username,
                Password = password,
                RememberMe = true
            };

            string jsonContent = Newtonsoft.Json.JsonConvert.SerializeObject(requestData);
            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            HttpResponseMessage response = await client.PostAsync(apiUrl, content);

            if (response.IsSuccessStatusCode)
            {
                string tokenResponse = await response.Content.ReadAsStringAsync();
                JObject jsonToken = JObject.Parse(tokenResponse);
                string token = jsonToken["token"].ToString();
                return token;
            }
            else
            {
                Console.WriteLine($"Ошибка при получении токена. Код ответа: {response.StatusCode}");
                return null;
            }
        }
    }

    // Получение списка клиентов по апи GET запросом
    static async Task<string> GetClients(string apiUrl, string token)
    {
        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            HttpResponseMessage response = await client.GetAsync(apiUrl);

            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync();
            }
            else
            {
                Console.WriteLine($"Ошибка при получении списка клиентов GET-запросом. Код ответа: {response.StatusCode}");
                return null;
            }
        }
    }

    // Получение списка проектов для конкретного клиента по апи GET запросом
    static async Task<string> GetProjects(string apiUrl, string token, string clientId)
    {
        using (HttpClient client = new HttpClient())
        {
            // Добавляем параметр clientId к URL
            string fullUrl = $"{apiUrl}?clientId={clientId}";

            // Устанавливаем токен в заголовок запроса
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            // Выполняем GET-запрос для получения данных
            HttpResponseMessage response = await client.GetAsync(fullUrl);

            if (response.IsSuccessStatusCode)
            {
                // Получаем данные из успешного ответа API
                return await response.Content.ReadAsStringAsync();
            }
            else
            {
                Console.WriteLine($"Ошибка при получении списка проектов GET-запросом. Код ответа: {response.StatusCode}");
                return null;
            }
        }
    }

    // Получение списка проектов для конкретного клиента по апи GET запросом
    static async Task<string> GetSearchResult(string apiUrl, string token, string projectId)
    {
        using (HttpClient client = new HttpClient())
        {
            // Добавляем параметр clientId к URL
            string fullUrl = $"{apiUrl}?projectId={projectId}";

            // Устанавливаем токен в заголовок запроса
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            // Выполняем GET-запрос для получения данных
            HttpResponseMessage response = await client.GetAsync(fullUrl);

            if (response.IsSuccessStatusCode)
            {
                // Получаем данные из успешного ответа API
                return await response.Content.ReadAsStringAsync();
            }
            else
            {
                Console.WriteLine($"Ошибка при получении выдач GET-запросом. Код ответа: {response.StatusCode}");
                return null;
            }
        }
    }
}
