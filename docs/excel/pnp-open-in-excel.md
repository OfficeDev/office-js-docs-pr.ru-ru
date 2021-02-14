---
title: Откройте Excel на веб-странице и встрайте надстройки Office
description: Откройте Excel на своей веб-странице и встрайте надстройки Office.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 437174b2fe9d04e3b25d42159efe7b38f45eb90c
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237926"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Откройте Excel на веб-странице и встрайте надстройки Office

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Изображение кнопки Excel на веб-странице, которая открывает новый документ Excel с внедренными и автоматически открываваемой надстройками.":::

Расширите веб-приложение SaaS, чтобы клиенты могли открывать свои данные с веб-страницы непосредственно в Microsoft Excel. Распространенный сценарий состоит в том, что клиенты будут работать с данными в веб-приложении. Затем им нужно будет скопировать данные в документ Excel. Например, им может потребоваться выполнить дополнительный анализ с помощью Excel. Как правило, клиент должен экспортировать данные в файл, например CSV-файл, а затем импортировать эти данные в Excel. Им также необходимо вручную добавить надстройки Office в документ.

Уменьшите количество действий, чтобы нажать одну кнопку на веб-странице, которая создает и открывает документ Excel. Вы также можете встраить свою надстройка Office в документ и отобразить ее при его открываемом документе. Это гарантирует, что у клиента по-прежнему будет доступ к функциям вашего приложения. Когда документ откроется, данные, выбранные клиентом, и ваша надстройка Office уже доступны для продолжения работы.

В этой статье показано, как реализовать этот сценарий в собственном веб-приложении SaaS.

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>Создание нового документа Excel и встраивайте надстройки Office

Сначала рассмотрим, как создать документ Excel на веб-странице и встраить надстройки в документ. В [примере кода надстройки Office OOXML Embed](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) показано, как встраить надстройку [Script Lab](https://appsource.microsoft.com/product/office/wa104380862) в новый документ Office. Хотя пример работает с любым документом Office, мы просто сосредоточимся на электронных таблицах Excel в этой статье. Чтобы создать и запустить пример, с помощью следующих действий.

1. Извлеките пример кода  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip из папки на компьютере.
2. Чтобы выполнить сборку и запустить пример, выполните действия, которые необходимо выполнить в разделе "Использование **проекта"** в документе readme.
3. При запуске примера отобразится веб-страница, аналогичная показанной на следующем снимке экрана. Используйте веб-страницу, чтобы создать новый документ Excel, содержащий Script Lab при его открытие.
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Снимок экрана веб-страницы, отображаемой в примере лаборатории встраив сценариев для выбора файла Excel и встраив в него надстройку лаборатории сценариев.":::

### <a name="how-the-sample-works"></a>Как работает пример

Пример кода использует пакет SDK OOXML для встраив надстройку Script Lab в выбираемый документ Excel. Следующая информация взята из раздела [ **"О коде"**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) в файле readme.

Файл **Home.aspx.cs:**

- Предоставляет обработчики событий кнопок и основные манипуляции с пользовательским интерфейсом.
- Использует стандартные ASP.NET для отправки и загрузки файла.
- Использует расширение имени загруженного файла (XLSX, DOCX или PPTX) для определения типа файла. Это необходимо сделать с самого начала, так как в open XML SDK обычно есть отдельные API для каждого типа файла.
- Вызывает **OOXMLHelper** для проверки файла и вызывает **AddInEmbedder,** чтобы встраить Script Lab в файл и настроить автоматическое открытие.

Файл **AddInEmbedder.cs:**

- Предоставляет основную бизнес-логику, которая в данном примере представляет собой метод, встраив в script Lab.
- Вызывает помощник OOXML в зависимости от типа файла.

Файл **OOXMLHelper.cs:**

- Предоставляет все подробные манипуляции OOXML.
- Использует стандартный метод проверки файла Office, который просто вызвать метод **Document.Open.** Если файл недействителен, метод выдаст исключение.
- Содержит в основном код, созданный средствами open XML 2.5 SDK Productivity Tools, доступный по ссылке для [Open XML 2.5 SDK.](/office/open-xml/open-xml-sdk)

Метод **GenerateWebExtensionPart1Content** в файле **OOXMLHelper.cs** задает ссылку на ИД Script Lab в Microsoft AppSource:

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- Значение **StoreType** — "OMEX", псевдоним для Microsoft AppSource.
- Значением **в Магазине** является en-US, найденный в разделе о культуре Microsoft AppSource для Script Lab.
- Значение **id** — это ИД ресурсов Microsoft AppSource для Script Lab.

При настройке надстройки из каталога файловой папки для автоматического открытия используются разные значения:

Значение **StoreType** — FileSystem.

- Значение **Store** — это URL-адрес сетевой обоймы; например, \\ \\ "MyComputer \\ MySharedFolder". Это должен быть точный URL-адрес, который отображается в качестве адреса доверенного каталога для share в центре управления доверием Office.
- Значение **id** — это ИД приложения в манифесте надстройки.
> [!NOTE]
> Дополнительные сведения об альтернативных значениях для этих атрибутов см. в подключе "Автоматическое открытие области задач [с документом".](../develop/automatically-open-a-task-pane-with-a-document.md)

## <a name="use-the-fluent-ui"></a>Использование пользовательского интерфейса Fluent

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Значки пользовательского интерфейса Fluent для Word, Excel и PowerPoint.":::

Лучше всего использовать пользовательский интерфейс Fluent, чтобы помочь пользователям перейти между продуктами Майкрософт. Всегда следует использовать значок Office, чтобы указать, какое приложение Office будет запущено с веб-страницы. Изменяем пример кода, чтобы использовать значок Excel, чтобы указать, что приложение Excel запускается.

1. Откройте пример в Visual Studio.
1. Откройте **страницу Home.aspx.**
1. Найдите следующий код, который является кнопкой скачивания в форме:
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. Замените код кнопки на следующий тег изображения.
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. Нажмите **F5** (или **отладка > начать отладку).** Значок появится при загрузке домашней страницы.

Дополнительные сведения [см.](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) на портале разработчика пользовательского интерфейса Fluent.  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Отправка документа Excel в Microsoft OneDrive

Мы рекомендуем отправить новые документы в OneDrive, если ваш клиент использует OneDrive. Это упрощает поиск документов и работу с ними. Давайте создадим новый пример кода и посмотрим, как можно использовать пакет SDK Microsoft Graph для отправки нового документа Excel в OneDrive.

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>Использование краткого запуска для создания нового веб-приложения Microsoft Graph

1. Чтобы создать и открыть пример кода для краткого запуска, взаимодействующих со службами Office, перейдите к этой [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) теме и следуйте этим шагам.
1. На **шаге 1. Выберите язык** или платформу, выберите ASP.NET **MVC.** Несмотря на то, что в процедурах, которые ASP.NET MVC, используются шаблоны, применимые к любому языку или платформе.
1. На **шаге 2. Получите ИД** и секрет приложения, выберите **"Получить ИД и секрет приложения".**
1. Во sign in to your Microsoft 365 account.  
1. На **веб-странице please save your app secret** web page, save the app secret to a file location where you can retrieve and use it later.
1. Choose **Got it, take me back to the quick start**.
1. Шаг **2. Регистрация прошла успешно!** Введите созданный секрет приложения.
1. На **шаге 3. Начните кодирование,** выберите Скачать пример кода **на основе SDK.**
1. Извлеките ZIP-папку скачивания в локализованную папку.  
1. Откройте файл graph-tutorial.sln в Visual Studio 2019.
1. Создайте и запустите решение и подтвердите, что оно работает правильно. Для просмотра календаря Microsoft 365 вы сможете использовать веб-страницу календаря.

### <a name="upload-a-file-to-onedrive"></a>Отправка файла в OneDrive

1. Откройте решение **graph-tutorial.sln** в Visual Studio 2019 г. и откройтеPrivateSettings.config **файла.**
1. Добавьте новую область **Files.ReadWrite в** ключ   **ida:AppScopes,** чтобы он выглядел следующим образом:
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. Откройте **файл Index.cshtml.**
1. Вставьте следующий код ActionLink, чтобы создать кнопку для отправки файла в OneDrive.
    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```
1. Откройте файл **HomeController.cs** файла.
1. Вставьте следующий код для обработки запроса из ссылки на действие.
    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```
1. Откройте **GraphHelper.cs** файла.
1. Вставьте следующий код, чтобы вызвать API Microsoft Graph для создания нового файла в OneDrive.
    ```csharp
    public static async Task UploadFile(string fileName, System.IO.MemoryStream stream)
        {
           var graphClient = GetAuthenticatedClient();
            await graphClient.Me
                .Drive
                .Root
                .ItemWithPath(fileName)
                .Content
                .Request()
                .PutAsync<DriveItem>(stream);
            return;
        }
    ```
1. Нажмите **F5** (или **отладка > начать отладку).** Запустится веб-приложение.
1. Нажмите **кнопку "Щелкните здесь", чтобы войти** и войти.
1. Щелкните **здесь, чтобы создать файл в OneDrive.**
1. Откройте новую вкладку браузера и вопишите в свою учетную запись OneDrive. Файл test.txt в корневой папке.

Теперь, когда вы узнали, как отправить файл в OneDrive, вы можете повторно использовать этот код для отправки любого документа Excel, который вы создаете.

## <a name="additional-considerations-for-your-solution"></a>Дополнительные соображения по вашему решению

Каждое решение отличается с точки зрения технологий и подходов. Следующие вопросы помогут вам спланировать изменение решения для открытия документов и встраив надстройки Office.

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Создание новой таблицы Excel на веб-странице

В примере изменяется существующий документ Excel. Более распространенный сценарий — создание новой таблицы Excel на веб-странице. Дополнительные сведения о создании новой таблицы можно  найти в документе "Создание таблицы" с помощью имени файла. В этой статье показано, как создать файл локально, но вы также можете создать файл в потоке с помощью перегрузки метода SpreadsheetDocument.Create.

### <a name="read-custom-properties-when-your-add-in-starts"></a>Чтение настраиваемого свойства при его начале

В примере кода код сохраняет код фрагмента в новом документе Excel с помощью OOXML SDK. Script Lab считыет код фрагмента из документа Excel, а затем отображает этот код при его открывлении. Может потребоваться отправить настраиваемые свойства в собственную надстройку (например, строку запроса или временный маркер проверки подлинности).) Подробные сведения о том, как читать настраиваемые свойства при ее создании, см. в документе **Persisting add-in** state and settings.

### <a name="initialize-the-excel-document-with-data"></a>Инициализация документа Excel с данными

Как правило, когда клиент открывает документ Excel с веб-сайта, он ожидает, что он будет содержать некоторые данные с веб-сайта. Существует несколько способов записи данных в документ.

- **Используйте OOXML SDK для записи данных.** Пакет SDK можно использовать для прямой записи любых данных в документ. Этот подход удобен, если вы хотите, чтобы данные были доступны при его открытом документе.
- **Передайте настраиваемые свойства запроса в надстройки Office.** При генерации документа в надстройку Office встраилось настраиваемые свойства, которые содержат строку запроса, которая извлекает все необходимые данные. Когда надстройка откроется, она извлекает запрос, запускает запрос и использует API JS Для вставки результата запроса в документ.

### <a name="working-with-the-ooxml-sdk"></a>Работа с OOXML SDK

OOXML SDK основан на .NET. Если веб-приложение не является .NET, необходимо найти альтернативный способ работы с OOXML.

Существует версия JavaScript OOXML SDK, доступная в [Open XML SDK для JavaScript.](https://archive.codeplex.com/?p=openxmlsdkjs)

Код OOXML можно разместить в функции Azure, чтобы отделить код .NET от остальной части веб-приложения. Затем вызовите функцию Azure (для создания документа Excel) из веб-приложения. Дополнительные сведения о функциях Azure см. [в введение в функции Azure.](/azure/azure-functions/functions-overview)

### <a name="use-single-sign-on"></a>Использование единого вход

Чтобы упростить проверку подлинности, рекомендуем надстройка реализовать единый вход. Дополнительные сведения см. в документе ["Включить единый вход для надстройки Office"](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>См. также

- [Добро пожаловать на страницу пакета SDK 2.5 Open XML для Office](/office/open-xml/open-xml-sdk)
- [Автоматическое открытие области задач с документом](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md)
- [Создайте документ электронной таблицы, указав имя файла](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)