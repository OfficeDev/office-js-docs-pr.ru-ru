---
title: Откройте Excel веб-страницы и встроите Office надстройки
description: Откройте Excel веб-страницы и встроите Office надстройки.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: b3343fb5e34e010d8bcbf0245be6fac6b1187f3b
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744055"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Откройте Excel веб-страницы и встроите Office надстройки

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Изображение кнопки Excel на веб-странице, открываемой Excel с помощью встроенной надстройки и автоматического открытия.":::

Расширите веб-приложение SaaS, чтобы клиенты могли открывать данные с веб-страницы непосредственно Microsoft Excel. Распространенный сценарий состоит в том, что клиенты будут работать с данными в вашем веб-приложении. Затем они захотят скопировать данные в Excel документа. Например, им может потребоваться выполнить дополнительный анализ с помощью Excel. Как правило, клиент должен экспортировать данные в файл, например .csv файл, а затем импортировать эти данные в Excel. Они также должны вручную добавлять Office надстройки в документ.

Уменьшите количество действий, чтобы одна кнопка нажала на веб-страницу, которая создает и открывает Excel документ. Вы также можете встраить Office в документ и отобразить его при открываемом документе. Это гарантирует, что клиент по-прежнему имеет доступ к функциям приложения. Когда документ откроется, данные, выбранные клиентом, и Office надстройка уже доступны для их продолжения работы.

В этой статье показаны код и методы реализации этого сценария в собственном веб-приложении SaaS.

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>Создание нового документа Excel и встраить Office надстройки

Сначала рассмотрим, как создать Excel на веб-странице и встраить надстройки в документ. В Office пример кода надстройки [OOXML](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) показано, как встраить Script Lab [в](https://appsource.microsoft.com/product/office/wa104380862) новый Office документ. Хотя пример работает с любым Office документом, мы просто сосредоточимся на Excel таблицах в этой статье. Для создания и запуска примера используйте следующие действия.

1. Извлечение примера кода  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip из папки на компьютере.
2. Чтобы создать и запустить пример, выполните действия в разделе **Использование раздела проекта** readme.
3. При запуске примера будет отображаться веб-страница, аналогичная следующему скриншоту. Используйте веб-страницу для создания нового документа Excel, который содержит Script Lab при его открытиях.
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Снимок экрана веб-страницы, отображаемой в примере лаборатории сценариев для выбора файла Excel и встраив в него надстройку лаборатории скриптов.":::

### <a name="how-the-sample-works"></a>Как работает пример

Пример кода использует SDK OOXML для встройки надстройки Script Lab в Excel документ, который вы выбираете. Следующие сведения взяты из раздела [**О коде**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) в файле readme.

Файл **Home.aspx.cs**:

- Предоставляет обработчики событий кнопки и основные манипуляции с пользовательским интерфейсом.
- Использует стандартные ASP.NET для загрузки и загрузки файла.
- Для определения типа файла используется расширение имени файла загруженного файла (xlsx, docx или pptx). Это необходимо сделать с самого начала, так как SDK Open XML обычно имеет отдельные API для каждого типа файла.
- Вызовы в **OOXMLHelper** для проверки файла и вызовы в **AddInEmbedder** для встраить Script Lab файл и установить для автоматического открытия.

Файл **AddInEmbedder.cs**:

- Предоставляет основную бизнес-логику, которая в этом примере представляет собой метод, который встраит Script Lab.
- Делает вызовы в помощник OOXML в зависимости от типа файла.

Файл **OOXMLHelper.cs**:

- Предоставляет все подробные манипуляции OOXML.
- Используется стандартный метод проверки файла Office, который является просто вызовом метода **Document.Open** на нем. Если файл недействителен, метод бросает исключение.
- Содержит в основном код, созданный средствами производительности Open XML 2.5 SDK, которые доступны по ссылке [для SDK Open XML 2.5](/office/open-xml/open-xml-sdk).

Метод **GenerateWebExtensionPart1Content** в файле **OOXMLHelper.cs** задает ссылку на ID Script Lab в Microsoft AppSource:

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- Значение **StoreType** — "OMEX", псевдоним Microsoft AppSource.
- Значение **Store** — "en-US", найденное в разделе Культура Microsoft AppSource для Script Lab.
- Значение **Id — это ID** актива Microsoft AppSource для Script Lab.

Если вы настраивает надстройка из каталога файлового обмена для автоматического открытия, вы будете использовать различные значения:

Значение **StoreType** — "FileSystem".

- Значение **Store** — ЭТО URL-адрес сетевой доли; например, "\\\\MyComputerMySharedFolder\\". Это должен быть точный URL-адрес, который отображается как доверенный адрес каталога в Office Центре доверия.
- Значение **Id** — это ID приложения в манифесте надстройки.
> [!NOTE]
> Дополнительные сведения об альтернативных значениях для этих атрибутов см. в тексте Автоматическое открытие области [задач с помощью документа](../develop/automatically-open-a-task-pane-with-a-document.md).

## <a name="use-the-fluent-ui"></a>Использование пользовательского Fluent интерфейса

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Fluent пользовательского интерфейса для Word, Excel и PowerPoint.":::

Лучше всего использовать пользовательский интерфейс Fluent, чтобы помочь пользователям перейти между продуктами Майкрософт. Всегда следует использовать значок Office, чтобы указать, Office приложение будет запущено с вашей веб-страницы. Давайте изменяем пример кода, чтобы использовать значок Excel, чтобы указать, что оно запускает Excel приложение.

1. Откройте пример в Visual Studio.
1. Откройте **страницу Home.aspx** .
1. Найдите следующий код, который является кнопкой загрузки в форме.

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. Замените код кнопки на следующий тег изображения.

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. Нажмите **кнопку F5** (или **отладка > начать отладку**). Значок появится при загрузке домашней страницы.

Дополнительные сведения см. [Office Значки](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) бренда на портале Fluent пользовательского интерфейса.  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Upload Excel для Microsoft OneDrive

Мы рекомендуем загружать новые документы в OneDrive, если клиент использует OneDrive. Это упрощает поиск документов и работу с ними. Давайте создадим новый пример кода и посмотрим, как можно использовать SDK microsoft Graph для отправки нового документа Excel в OneDrive.

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>Использование быстрого запуска для создания нового веб-приложения Graph Microsoft

1. Выполните [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) действия по созданию и запуску примера кода быстрого запуска, который взаимодействует с Office службами.
1. В **шаге 1. Выберите язык** или платформу, выберите ASP.NET **MVC**. Хотя в этой процедуре используется ASP.NET MVC, действия следуют шаблону, который применяется к любому языку или платформе.
1. На **шаге 2. Получите ID** приложения и секрет, выберите **Get an app ID and secret**.
1. Вопишите в свою Microsoft 365 учетную запись.  
1. На странице **Пожалуйста, сохраните секретную** веб-страницу приложения, сохраните секрет приложения в расположении файла, где вы можете получить и использовать его позже.
1. Выберите **Got it, отбери меня к быстрому началу**.
1. В **шаге 2: Регистрация успешна!** Введите созданный секрет приложения.
1. В **шаге 3. Начните кодирование**, выберите **Скачайте образец кода на основе SDK**.
1. Извлечение папки для скачивания почтовых индексов в локализованную папку.  
1. Откройте файл graph-tutorial.sln в Visual Studio 2019 г.
1. Создайте и запустите решение и подтвердите, что оно работает правильно. Вы должны иметь возможность использовать веб-страницу календаря для просмотра Microsoft 365 календаря.

### <a name="upload-a-file-to-onedrive"></a>Upload файл для OneDrive

1. Откройте решение **graph-tutorial.sln** в Visual Studio 2019 г. и **откройтеPrivateSettings.config** файл.

1. Добавьте новую область **Files.ReadWriteto**  клавишу **ida:AppScopes** , чтобы она выглядела как следующий код.

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. Откройте **файл Index.cshtml** .
1. Вставьте следующий код ActionLink, чтобы создать кнопку для отправки файла в OneDrive.

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. Откройте **файл HomeController.cs** .
1. Вставьте следующий код для обработки запроса из ссылки действия.

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. Откройте **файл GraphHelper.cs** .
1. Вставьте следующий код, чтобы вызвать API microsoft Graph, чтобы создать новый файл в OneDrive.

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

1. Нажмите **кнопку F5** (или **отладка > начать отладку**). Начнет работу веб-приложение.
1. Выберите **нажмите здесь, чтобы войти** и войти.
1. Выберите **нажмите здесь, чтобы создать новый файл на OneDrive**.
1. Откройте новую вкладку браузера и вопишитесь в свою OneDrive учетную запись. Вы увидите файл test.txt в корневой папке.

Теперь, когда вы узнали, как загрузить файл в OneDrive, вы можете повторно использовать этот код для загрузки любого Excel, который вы создаете.

## <a name="additional-considerations-for-your-solution"></a>Дополнительные соображения для решения

Каждое решение отличается с точки зрения технологий и подходов. Следующие соображения помогут вам спланировать изменение решения, чтобы открыть документы и Office надстройки.

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Создание новой Excel таблицы с веб-страницы

В примере изменяется существующий Excel документ. Более распространенным сценарием является создание новой Excel таблицы с веб-страницы. Дополнительные сведения о создании новой таблицы можно найти в документе **Create a spreadsheet,** предоставив имя файла. В этой статье показано, как создать файл локально, но вы также можете создать файл в потоке с помощью перегрузки в методе SpreadsheetDocument.Create.

### <a name="read-custom-properties-when-your-add-in-starts"></a>Чтение пользовательских свойств при старте надстройки

В примере кода в новом документе Excel SDK OOXML. Script Lab код фрагмента из документа Excel, а затем отображает этот фрагмент кода при его открывлении. Возможно, вам потребуется отправить настраиваемые свойства в собственную надстройку (например, строку запроса или временный маркер проверки подлинности).) Дополнительные сведения о том, как читать настраиваемые свойства при старте надстройки, см. в публикации **Persisting add-in state and** settings.

### <a name="initialize-the-excel-document-with-data"></a>Инициализация Excel с данными

Как правило, когда клиент открывает Excel документа с веб-сайта, он ожидает, что документ содержит некоторые данные с веб-сайта. Существует несколько способов записи данных в документ.

- **Для записи данных используйте SDK OOXML**. Вы можете использовать SDK для непосредственного записи любых данных в документ. Этот подход полезен, если вы хотите, чтобы данные были доступны сразу после открытия документа.
- **Передай пользовательское свойство запроса Office надстройки**. При генерации документа встраив настраиваемую свойство для надстройки Office, содержаную строку запроса, которая извлекает все необходимые данные. Когда надстройка открывается, она извлекает запрос, запускает запрос и использует API Office JS, чтобы вставить результат запроса в документ.

### <a name="working-with-the-ooxml-sdk"></a>Работа с SDK OOXML

SDK OOXML основан на .NET. Если в вашем веб-приложении нет .NET, необходимо искать альтернативный способ работы с OOXML.

Код OOXML можно разместить в функции Azure, чтобы отделить код .NET от остальной части веб-приложения. Затем вызывайте функцию Azure (для создания Excel документа) из веб-приложения. Дополнительные сведения о функциях Azure см. [в предисловии к Azure Functions](/azure/azure-functions/functions-overview).

### <a name="use-single-sign-on"></a>Использование единого входного

Чтобы упростить проверку подлинности, рекомендуется, чтобы надстройка реализовала один вход. Дополнительные сведения см. в [документе Enable single sign-on for Office надстройки](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>См. также

- [Добро пожаловать на страницу пакета SDK 2.5 Open XML для Office](/office/open-xml/open-xml-sdk)
- [Автоматическое открытие области задач с документом](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md)
- [Создайте документ электронной таблицы, указав имя файла](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)