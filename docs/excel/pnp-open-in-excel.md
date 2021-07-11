---
title: Откройте Excel веб-страницы и встроите Office надстройки
description: Откройте Excel веб-страницы и встроите Office надстройки.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 18f40b0030f4132a413a879e8b3419af49984b45
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349380"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a><span data-ttu-id="86dcb-103">Откройте Excel веб-страницы и встроите Office надстройки</span><span class="sxs-lookup"><span data-stu-id="86dcb-103">Open Excel from your web page and embed your Office Add-in</span></span>

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Изображение кнопки Excel на веб-странице, открываемой Excel документа с помощью встроенной надстройки и автоматического открытия.":::

<span data-ttu-id="86dcb-105">Расширите веб-приложение SaaS, чтобы клиенты могли открывать данные с веб-страницы непосредственно Microsoft Excel.</span><span class="sxs-lookup"><span data-stu-id="86dcb-105">Extend your SaaS web application so that your customers can open their data from a web page directly to Microsoft Excel.</span></span> <span data-ttu-id="86dcb-106">Распространенный сценарий состоит в том, что клиенты будут работать с данными в вашем веб-приложении.</span><span class="sxs-lookup"><span data-stu-id="86dcb-106">A common scenario is that customers will be working with data in your web application.</span></span> <span data-ttu-id="86dcb-107">Затем они захотят скопировать данные в Excel документ.</span><span class="sxs-lookup"><span data-stu-id="86dcb-107">Then they’ll want to copy the data into an Excel document.</span></span> <span data-ttu-id="86dcb-108">Например, им может потребоваться выполнить дополнительный анализ с помощью Excel.</span><span class="sxs-lookup"><span data-stu-id="86dcb-108">For example, they may want to perform additional analysis using Excel.</span></span> <span data-ttu-id="86dcb-109">Как правило, клиент должен экспортировать данные в файл, например .csv файл, а затем импортировать эти данные в Excel.</span><span class="sxs-lookup"><span data-stu-id="86dcb-109">Typically, the customer is required to export the data to a file, such as a .csv file, and then import that data into Excel.</span></span> <span data-ttu-id="86dcb-110">Они также должны вручную добавлять Office надстройки в документ.</span><span class="sxs-lookup"><span data-stu-id="86dcb-110">They also have to manually add your Office Add-in to the document.</span></span>

<span data-ttu-id="86dcb-111">Уменьшите количество действий до одной кнопки на веб-странице, которая создает и открывает Excel документа.</span><span class="sxs-lookup"><span data-stu-id="86dcb-111">Reduce the number of steps to a single button click on your web page that generates and opens the Excel document.</span></span> <span data-ttu-id="86dcb-112">Вы также можете встраить Office надстройки в документ и отобразить его при открываемом документе.</span><span class="sxs-lookup"><span data-stu-id="86dcb-112">You can also embed your Office Add-in inside the document and display it when the document opens.</span></span> <span data-ttu-id="86dcb-113">Это гарантирует, что клиент по-прежнему имеет доступ к функциям приложения.</span><span class="sxs-lookup"><span data-stu-id="86dcb-113">This ensures the customer still has access to your application features.</span></span> <span data-ttu-id="86dcb-114">Когда документ откроется, данные, выбранные клиентом, и Office надстройка уже доступны для их продолжения работы.</span><span class="sxs-lookup"><span data-stu-id="86dcb-114">When the document opens, the data the customer selected, and your Office Add-in is already available for them to continue working.</span></span>

<span data-ttu-id="86dcb-115">В этой статье показаны код и методы реализации этого сценария в собственном веб-приложении SaaS.</span><span class="sxs-lookup"><span data-stu-id="86dcb-115">This article shows you code and techniques for implementing this scenario in your own SaaS web application.</span></span>

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a><span data-ttu-id="86dcb-116">Создание нового документа Excel и встраив Office надстройки</span><span class="sxs-lookup"><span data-stu-id="86dcb-116">Create a new Excel document and embed an Office Add-in</span></span>

<span data-ttu-id="86dcb-117">Сначала рассмотрим, как создать документ Excel веб-страницы и встраить надстройки в документ.</span><span class="sxs-lookup"><span data-stu-id="86dcb-117">First, let’s learn how to create an Excel document from a web page, and embed an add-in into the document.</span></span> <span data-ttu-id="86dcb-118">В Office пример кода надстройки [OOXML](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) показано, как встраить Script Lab [в](https://appsource.microsoft.com/product/office/wa104380862) новый Office документ.</span><span class="sxs-lookup"><span data-stu-id="86dcb-118">The [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document.</span></span> <span data-ttu-id="86dcb-119">Хотя пример работает с любым Office документом, мы сосредоточимся на Excel таблицах в этой статье.</span><span class="sxs-lookup"><span data-stu-id="86dcb-119">Although the sample works with any Office document, we’ll just focus on Excel spreadsheets in this article.</span></span> <span data-ttu-id="86dcb-120">Для создания и запуска примера используйте следующие действия.</span><span class="sxs-lookup"><span data-stu-id="86dcb-120">Use the following steps to build and run the sample.</span></span>

1. <span data-ttu-id="86dcb-121">Извлечение примера кода  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip из папки на компьютере.</span><span class="sxs-lookup"><span data-stu-id="86dcb-121">Extract the sample code from  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip into a folder on your computer.</span></span>
2. <span data-ttu-id="86dcb-122">Чтобы создать и запустить пример, выполните действия в разделе **Использование раздела проекта** readme.</span><span class="sxs-lookup"><span data-stu-id="86dcb-122">To build and run the sample, follow the steps in the **To use the project** section of the readme.</span></span>
3. <span data-ttu-id="86dcb-123">При запуске примера будет отображаться веб-страница, аналогичная следующему скриншоту.</span><span class="sxs-lookup"><span data-stu-id="86dcb-123">When you run the sample it will display a web page similar to the following screenshot.</span></span> <span data-ttu-id="86dcb-124">Используйте веб-страницу для создания нового документа Excel, который содержит Script Lab при ее открываемом ок.</span><span class="sxs-lookup"><span data-stu-id="86dcb-124">Use the web page to create a new Excel document that contains Script Lab when it opens.</span></span>
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Снимок экрана веб-страницы, отображаемой в примере лаборатории сценариев для выбора Excel файла и встраив в него надстройку лаборатории скриптов.":::

### <a name="how-the-sample-works"></a><span data-ttu-id="86dcb-126">Как работает пример</span><span class="sxs-lookup"><span data-stu-id="86dcb-126">How the sample works</span></span>

<span data-ttu-id="86dcb-127">Пример кода использует SDK OOXML для встройки надстройки Script Lab в Excel документ, который вы выбираете.</span><span class="sxs-lookup"><span data-stu-id="86dcb-127">The sample code uses the OOXML SDK to embed the Script Lab add-in to the Excel document that you choose.</span></span> <span data-ttu-id="86dcb-128">Следующие сведения взяты из раздела [ **О коде**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) в файле readme.</span><span class="sxs-lookup"><span data-stu-id="86dcb-128">The following information is taken from the [**About the code** section](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) in the readme file.</span></span>

<span data-ttu-id="86dcb-129">Файл **Home.aspx.cs:**</span><span class="sxs-lookup"><span data-stu-id="86dcb-129">The file **Home.aspx.cs**:</span></span>

- <span data-ttu-id="86dcb-130">Предоставляет обработчики событий кнопки и основные манипуляции с пользовательским интерфейсом.</span><span class="sxs-lookup"><span data-stu-id="86dcb-130">Provides the button event handlers and basic UI manipulation.</span></span>
- <span data-ttu-id="86dcb-131">Использует стандартные ASP.NET для загрузки и загрузки файла.</span><span class="sxs-lookup"><span data-stu-id="86dcb-131">Uses standard ASP.NET techniques to upload and download the file.</span></span>
- <span data-ttu-id="86dcb-132">Для определения типа файла используется расширение имени файла загруженного файла (xlsx, docx или pptx).</span><span class="sxs-lookup"><span data-stu-id="86dcb-132">Uses the file name extension of the uploaded file (xlsx, docx, or pptx) to determine the type of file.</span></span> <span data-ttu-id="86dcb-133">Это необходимо сделать с самого начала, так как SDK Open XML обычно имеет отдельные API для каждого типа файла.</span><span class="sxs-lookup"><span data-stu-id="86dcb-133">This needs to be done at the outset because the Open XML SDK generally has distinct APIs for each type of file.</span></span>
- <span data-ttu-id="86dcb-134">Звонки в **OOXMLHelper** для проверки файла и вызовы в **AddInEmbedder** для встраить Script Lab в файл и установить для автоматического открытия.</span><span class="sxs-lookup"><span data-stu-id="86dcb-134">Calls into the **OOXMLHelper** to validate the file and calls into the **AddInEmbedder** to embed Script Lab in the file and set to automatically open.</span></span>

<span data-ttu-id="86dcb-135">Файл **AddInEmbedder.cs:**</span><span class="sxs-lookup"><span data-stu-id="86dcb-135">The file **AddInEmbedder.cs**:</span></span>

- <span data-ttu-id="86dcb-136">Предоставляет основную бизнес-логику, которая в этом примере представляет собой метод, который встраит Script Lab.</span><span class="sxs-lookup"><span data-stu-id="86dcb-136">Provides the main business logic, which in this sample is a method that embeds Script Lab.</span></span>
- <span data-ttu-id="86dcb-137">Делает вызовы в помощник OOXML в зависимости от типа файла.</span><span class="sxs-lookup"><span data-stu-id="86dcb-137">Makes calls into the OOXML helper based on the type of the file.</span></span>

<span data-ttu-id="86dcb-138">Файл **OOXMLHelper.cs:**</span><span class="sxs-lookup"><span data-stu-id="86dcb-138">The file **OOXMLHelper.cs**:</span></span>

- <span data-ttu-id="86dcb-139">Предоставляет все подробные манипуляции OOXML.</span><span class="sxs-lookup"><span data-stu-id="86dcb-139">Provides all the detailed OOXML manipulation.</span></span>
- <span data-ttu-id="86dcb-140">Используется стандартный метод проверки Office файла, который является просто вызовом метода **Document.Open** на нем.</span><span class="sxs-lookup"><span data-stu-id="86dcb-140">Uses a standard technique for validating the Office file, which is simply to call the **Document.Open** method on it.</span></span> <span data-ttu-id="86dcb-141">Если файл недействителен, метод бросает исключение.</span><span class="sxs-lookup"><span data-stu-id="86dcb-141">If the file is invalid, the method throws an exception.</span></span>
- <span data-ttu-id="86dcb-142">Содержит в основном код, созданный средствами производительности Open XML 2.5 SDK, доступными по ссылке для [SDK Open XML 2.5](/office/open-xml/open-xml-sdk).</span><span class="sxs-lookup"><span data-stu-id="86dcb-142">Contains mainly code that was generated by the Open XML 2.5 SDK Productivity Tools which are available at the link for the [Open XML 2.5 SDK](/office/open-xml/open-xml-sdk).</span></span>

<span data-ttu-id="86dcb-143">Метод **GenerateWebExtensionPart1Content** в файле **OOXMLHelper.cs** задает ссылку на ID Script Lab в Microsoft AppSource:</span><span class="sxs-lookup"><span data-stu-id="86dcb-143">The **GenerateWebExtensionPart1Content** method in the **OOXMLHelper.cs** file sets the reference to the ID of Script Lab in Microsoft AppSource:</span></span>

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- <span data-ttu-id="86dcb-144">Значение **StoreType** — "OMEX", псевдоним Microsoft AppSource.</span><span class="sxs-lookup"><span data-stu-id="86dcb-144">The **StoreType** value is "OMEX", an alias for Microsoft AppSource.</span></span>
- <span data-ttu-id="86dcb-145">Значение **Store** — "en-US", найденное в разделе Культура Microsoft AppSource для Script Lab.</span><span class="sxs-lookup"><span data-stu-id="86dcb-145">The **Store** value is "en-US" found in the Microsoft AppSource culture section for Script Lab.</span></span>
- <span data-ttu-id="86dcb-146">Значение Id — это **ID** актива Microsoft AppSource для Script Lab.</span><span class="sxs-lookup"><span data-stu-id="86dcb-146">The **Id** value is the Microsoft AppSource asset ID for Script Lab.</span></span>

<span data-ttu-id="86dcb-147">Если вы настраивает надстройка из каталога файлового обмена для автоматического открытия, вы будете использовать различные значения:</span><span class="sxs-lookup"><span data-stu-id="86dcb-147">If you are setting up an add-in from a file share catalog for auto-open, you will use different values:</span></span>

<span data-ttu-id="86dcb-148">Значение **StoreType** — "FileSystem".</span><span class="sxs-lookup"><span data-stu-id="86dcb-148">The **StoreType** value is "FileSystem".</span></span>

- <span data-ttu-id="86dcb-149">Значение **Store** — ЭТО URL-адрес сетевой доли; например, \\ \\ "MyComputer \\ MySharedFolder".</span><span class="sxs-lookup"><span data-stu-id="86dcb-149">The **Store** value is the URL of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span> <span data-ttu-id="86dcb-150">Это должен быть точный URL-адрес, который отображается как доверенный адрес каталога в Office Центре доверия.</span><span class="sxs-lookup"><span data-stu-id="86dcb-150">This should be the exact URL that appears as the share's Trusted Catalog Address in the Office Trust Center.</span></span>
- <span data-ttu-id="86dcb-151">Значение **Id** — это ID приложения в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="86dcb-151">The **Id** value is the app ID in the add-ins manifest.</span></span>
> [!NOTE]
> <span data-ttu-id="86dcb-152">Дополнительные сведения об альтернативных значениях для этих атрибутов см. в тексте Автоматическое открытие области задач [с помощью документа.](../develop/automatically-open-a-task-pane-with-a-document.md)</span><span class="sxs-lookup"><span data-stu-id="86dcb-152">For more information about alternative values for these attributes, see [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md).</span></span>

## <a name="use-the-fluent-ui"></a><span data-ttu-id="86dcb-153">Использование Fluent пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="86dcb-153">Use the Fluent UI</span></span>

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Fluent Значки пользовательского интерфейса для Word, Excel и PowerPoint.":::

<span data-ttu-id="86dcb-155">Лучше всего использовать пользовательский Fluent, чтобы помочь пользователям перейти между продуктами Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="86dcb-155">A best practice is to use the Fluent UI to help your users transition between Microsoft products.</span></span> <span data-ttu-id="86dcb-156">Всегда следует использовать значок Office, чтобы указать, Office приложение будет запущено с вашей веб-страницы.</span><span class="sxs-lookup"><span data-stu-id="86dcb-156">You should always use an Office icon to indicate which Office application will be launched from your web page.</span></span> <span data-ttu-id="86dcb-157">Давайте изменяем пример кода, чтобы использовать значок Excel, чтобы указать, что оно запускает Excel приложение.</span><span class="sxs-lookup"><span data-stu-id="86dcb-157">Let’s modify the sample code to use the Excel icon to indicate that it launches the Excel application.</span></span>

1. <span data-ttu-id="86dcb-158">Откройте пример в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="86dcb-158">Open the sample in Visual Studio.</span></span>
1. <span data-ttu-id="86dcb-159">Откройте **страницу Home.aspx.**</span><span class="sxs-lookup"><span data-stu-id="86dcb-159">Open the **Home.aspx** page.</span></span>
1. <span data-ttu-id="86dcb-160">Найдите следующий код, который является кнопкой загрузки в форме.</span><span class="sxs-lookup"><span data-stu-id="86dcb-160">Find following code that is the download button on the form.</span></span>

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. <span data-ttu-id="86dcb-161">Замените код кнопки на следующий тег изображения.</span><span class="sxs-lookup"><span data-stu-id="86dcb-161">Replace the button code with the following image tag.</span></span>

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. <span data-ttu-id="86dcb-162">Нажмите **кнопку F5** (или **отладка > начать отладку).**</span><span class="sxs-lookup"><span data-stu-id="86dcb-162">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="86dcb-163">Значок появится при загрузке домашней страницы.</span><span class="sxs-lookup"><span data-stu-id="86dcb-163">You'll see the icon appear when the home page loads.</span></span>

<span data-ttu-id="86dcb-164">Дополнительные сведения см. [Office Значки](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) бренда на портале Fluent пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="86dcb-164">For more information, see [Office Brand Icons](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) on the Fluent UI developer portal.</span></span>  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a><span data-ttu-id="86dcb-165">Upload Excel документ для Microsoft OneDrive</span><span class="sxs-lookup"><span data-stu-id="86dcb-165">Upload the Excel document to Microsoft OneDrive</span></span>

<span data-ttu-id="86dcb-166">Мы рекомендуем загружать новые документы в OneDrive, если клиент использует OneDrive.</span><span class="sxs-lookup"><span data-stu-id="86dcb-166">We recommend uploading new documents to OneDrive if your customer uses OneDrive.</span></span> <span data-ttu-id="86dcb-167">Это упрощает поиск документов и работу с ними.</span><span class="sxs-lookup"><span data-stu-id="86dcb-167">This makes it easier for them to find and work with the documents.</span></span> <span data-ttu-id="86dcb-168">Давайте создадим новый пример кода и посмотрим, как можно использовать SDK microsoft Graph для отправки нового документа Excel в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="86dcb-168">Let’s create a new code sample and see how you can use the Microsoft Graph SDK to upload a new Excel document to OneDrive.</span></span>

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a><span data-ttu-id="86dcb-169">Используйте быстрое начало для создания нового веб-приложения Graph Microsoft</span><span class="sxs-lookup"><span data-stu-id="86dcb-169">Use a quick-start to build a new Microsoft Graph web application</span></span>

1. <span data-ttu-id="86dcb-170">Выполните действия по созданию и запуску примера кода быстрого запуска, который взаимодействует с Office [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) службами.</span><span class="sxs-lookup"><span data-stu-id="86dcb-170">Go to [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) and follow the steps to create and open a quick start code sample that interacts with Office services.</span></span>
1. <span data-ttu-id="86dcb-171">В **шаге 1. Выберите язык** или платформу, выберите ASP.NET **MVC**.</span><span class="sxs-lookup"><span data-stu-id="86dcb-171">In **step 1: Pick you language or platform**, choose **ASP.NET MVC**.</span></span> <span data-ttu-id="86dcb-172">Хотя в этой процедуре используется ASP.NET MVC, действия следуют шаблону, который применяется к любому языку или платформе.</span><span class="sxs-lookup"><span data-stu-id="86dcb-172">Although the steps in this procedure use the ASP.NET MVC option, the steps follow a pattern that apply to any language or platform.</span></span>
1. <span data-ttu-id="86dcb-173">В **шаге 2. Получите ID приложения** и секрет , выберите **Получить ID** приложения и секрет .</span><span class="sxs-lookup"><span data-stu-id="86dcb-173">In **step 2: Get an app ID and secret**, choose **Get an app ID and secret**.</span></span>
1. <span data-ttu-id="86dcb-174">Вопишите в свою Microsoft 365 учетную запись.</span><span class="sxs-lookup"><span data-stu-id="86dcb-174">Sign in to your Microsoft 365 account.</span></span>  
1. <span data-ttu-id="86dcb-175">На странице **Пожалуйста, сохраните секретную** веб-страницу приложения, сохраните секрет приложения в расположении файла, где вы можете получить и использовать его позже.</span><span class="sxs-lookup"><span data-stu-id="86dcb-175">On the **Please save your app secret** web page, save the app secret to a file location where you can retrieve and use it later.</span></span>
1. <span data-ttu-id="86dcb-176">Выберите **Got it, отбери меня к быстрому началу**.</span><span class="sxs-lookup"><span data-stu-id="86dcb-176">Choose **Got it, take me back to the quick start**.</span></span>
1. <span data-ttu-id="86dcb-177">В **шаге 2: Регистрация успешна!**</span><span class="sxs-lookup"><span data-stu-id="86dcb-177">In **step 2: Registration Successful!**</span></span> <span data-ttu-id="86dcb-178">Введите созданный секрет приложения.</span><span class="sxs-lookup"><span data-stu-id="86dcb-178">Enter the generated app secret.</span></span>
1. <span data-ttu-id="86dcb-179">В **шаге 3. Начните кодирование,** выберите Скачайте образец кода на основе **SDK.**</span><span class="sxs-lookup"><span data-stu-id="86dcb-179">In **step 3: Start coding**, choose **Download the SDK-based code sample**.</span></span>
1. <span data-ttu-id="86dcb-180">Извлечение папки для скачивания почтовых индексов в локализованную папку.</span><span class="sxs-lookup"><span data-stu-id="86dcb-180">Extract the download zip folder into a local folder.</span></span>  
1. <span data-ttu-id="86dcb-181">Откройте файл graph-tutorial.sln в Visual Studio 2019 г.</span><span class="sxs-lookup"><span data-stu-id="86dcb-181">Open the graph-tutorial.sln file in Visual Studio 2019.</span></span>
1. <span data-ttu-id="86dcb-182">Создайте и запустите решение и подтвердите, что оно работает правильно.</span><span class="sxs-lookup"><span data-stu-id="86dcb-182">Build and run the solution and confirm it is working correctly.</span></span> <span data-ttu-id="86dcb-183">Вы должны иметь возможность использовать веб-страницу календаря для просмотра Microsoft 365 календаря.</span><span class="sxs-lookup"><span data-stu-id="86dcb-183">You should be able to use the calendar web page to view your Microsoft 365 calendar.</span></span>

### <a name="upload-a-file-to-onedrive"></a><span data-ttu-id="86dcb-184">Upload файл для OneDrive</span><span class="sxs-lookup"><span data-stu-id="86dcb-184">Upload a file to OneDrive</span></span>

1. <span data-ttu-id="86dcb-185">Откройте решение **graph-tutorial.sln** в Visual Studio 2019 г. и откройте **PrivateSettings.config** файл.</span><span class="sxs-lookup"><span data-stu-id="86dcb-185">Open the **graph-tutorial.sln** solution in Visual Studio 2019, and open the **PrivateSettings.config** file.</span></span>
1. <span data-ttu-id="86dcb-186">Добавьте новую область **Files.ReadWrite** в ключ   **ida:AppScopes,** чтобы он выглядел как следующий код.</span><span class="sxs-lookup"><span data-stu-id="86dcb-186">Add a new scope **Files.ReadWrite** to the **ida:AppScopes** key so that it looks like the following code.</span></span>

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. <span data-ttu-id="86dcb-187">Откройте **файл Index.cshtml.**</span><span class="sxs-lookup"><span data-stu-id="86dcb-187">Open the **Index.cshtml** file.</span></span>
1. <span data-ttu-id="86dcb-188">Вставьте следующий код ActionLink, чтобы создать кнопку для отправки файла в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="86dcb-188">Insert the following ActionLink code to create a button to upload a file to OneDrive.</span></span>

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. <span data-ttu-id="86dcb-189">Откройте **файл HomeController.cs.**</span><span class="sxs-lookup"><span data-stu-id="86dcb-189">Open the **HomeController.cs** file.</span></span>
1. <span data-ttu-id="86dcb-190">Вставьте следующий код для обработки запроса из ссылки действия.</span><span class="sxs-lookup"><span data-stu-id="86dcb-190">Insert the following code to handle the request from the action link.</span></span>

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. <span data-ttu-id="86dcb-191">Откройте **файл GraphHelper.cs.**</span><span class="sxs-lookup"><span data-stu-id="86dcb-191">Open the **GraphHelper.cs** file.</span></span>
1. <span data-ttu-id="86dcb-192">Вставьте следующий код, чтобы вызвать API microsoft Graph, чтобы создать новый файл в OneDrive.</span><span class="sxs-lookup"><span data-stu-id="86dcb-192">Insert the following code to call the Microsoft Graph API to create a new file on OneDrive.</span></span>

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

1. <span data-ttu-id="86dcb-193">Нажмите **кнопку F5** (или **отладка > начать отладку).**</span><span class="sxs-lookup"><span data-stu-id="86dcb-193">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="86dcb-194">Начнет работу веб-приложение.</span><span class="sxs-lookup"><span data-stu-id="86dcb-194">The web application will start.</span></span>
1. <span data-ttu-id="86dcb-195">Выберите **нажмите здесь, чтобы войти** и войти.</span><span class="sxs-lookup"><span data-stu-id="86dcb-195">Choose **Click here to sign in**, and sign in.</span></span>
1. <span data-ttu-id="86dcb-196">Выберите **нажмите здесь, чтобы создать новый файл на OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="86dcb-196">Choose **Click here to create a new file on OneDrive**.</span></span>
1. <span data-ttu-id="86dcb-197">Откройте новую вкладку браузера и вопишитесь в свою OneDrive учетную запись.</span><span class="sxs-lookup"><span data-stu-id="86dcb-197">Open a new browser tab and sign in to your OneDrive account.</span></span> <span data-ttu-id="86dcb-198">В корневой папке test.txt файл.</span><span class="sxs-lookup"><span data-stu-id="86dcb-198">You'll see the test.txt file in the root folder.</span></span>

<span data-ttu-id="86dcb-199">Теперь, когда вы узнали, как загрузить файл в OneDrive, вы можете повторно использовать этот код, чтобы загрузить любой Excel, который вы создаете.</span><span class="sxs-lookup"><span data-stu-id="86dcb-199">Now that you've learned how to upload a file to OneDrive, you can reuse this code to upload any Excel document that you create.</span></span>

## <a name="additional-considerations-for-your-solution"></a><span data-ttu-id="86dcb-200">Дополнительные соображения для решения</span><span class="sxs-lookup"><span data-stu-id="86dcb-200">Additional considerations for your solution</span></span>

<span data-ttu-id="86dcb-201">Каждое решение отличается с точки зрения технологий и подходов.</span><span class="sxs-lookup"><span data-stu-id="86dcb-201">Everyone’s solution is different in terms of technologies and approaches.</span></span> <span data-ttu-id="86dcb-202">Следующие соображения помогут вам спланировать изменение решения, чтобы открыть документы и Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="86dcb-202">The following considerations will help you plan how to modify your solution to open documents and embed your Office Add-in.</span></span>

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a><span data-ttu-id="86dcb-203">Создание новой Excel таблицы с веб-страницы</span><span class="sxs-lookup"><span data-stu-id="86dcb-203">Create a new Excel spreadsheet from the web page</span></span>

<span data-ttu-id="86dcb-204">В примере изменяется существующий Excel документ.</span><span class="sxs-lookup"><span data-stu-id="86dcb-204">The sample modifies an existing Excel document.</span></span> <span data-ttu-id="86dcb-205">Более распространенным сценарием является создание новой Excel таблицы с веб-страницы.</span><span class="sxs-lookup"><span data-stu-id="86dcb-205">A more common scenario is that you’ll create a new Excel spreadsheet from your web page.</span></span> <span data-ttu-id="86dcb-206">Дополнительные сведения о создании новой таблицы можно найти в документе **Create a spreadsheet,** предоставив имя файла.</span><span class="sxs-lookup"><span data-stu-id="86dcb-206">You can find additional details on how to create a new spreadsheet in **Create a spreadsheet document** by providing a file name.</span></span> <span data-ttu-id="86dcb-207">В этой статье показано, как создать файл локально, но вы также можете создать файл в потоке с помощью перегрузки в методе SpreadsheetDocument.Create.</span><span class="sxs-lookup"><span data-stu-id="86dcb-207">This article shows how to create the file locally, but you can also create the file in a stream by using an overload on the SpreadsheetDocument.Create method.</span></span>

### <a name="read-custom-properties-when-your-add-in-starts"></a><span data-ttu-id="86dcb-208">Чтение пользовательских свойств при старте надстройки</span><span class="sxs-lookup"><span data-stu-id="86dcb-208">Read custom properties when your add-in starts</span></span>

<span data-ttu-id="86dcb-209">В примере кода хранится код фрагмента в новом документе Excel с помощью SDK OOXML.</span><span class="sxs-lookup"><span data-stu-id="86dcb-209">The code sample stores a snippet ID in the new Excel document using the OOXML SDK.</span></span> <span data-ttu-id="86dcb-210">Script Lab код фрагмента из документа Excel, а затем отображает этот фрагмент кода при его открывлении.</span><span class="sxs-lookup"><span data-stu-id="86dcb-210">Script Lab reads the snippet ID from the Excel document and then displays that snippet code when it opens.</span></span> <span data-ttu-id="86dcb-211">Возможно, вам потребуется отправить настраиваемые свойства в собственную надстройку (например, строку запроса или временный маркер проверки подлинности).) Дополнительные сведения о том, как читать настраиваемые свойства при старте надстройки, см. в публикации **Persisting add-in** state and settings.</span><span class="sxs-lookup"><span data-stu-id="86dcb-211">You may need to send custom properties to your own add-in (such as a query string, or temporary authentication token.) See **Persisting add-in state and settings** for complete details on how to read custom properties when your add-in starts.</span></span>

### <a name="initialize-the-excel-document-with-data"></a><span data-ttu-id="86dcb-212">Инициализация Excel с данными</span><span class="sxs-lookup"><span data-stu-id="86dcb-212">Initialize the Excel document with data</span></span>

<span data-ttu-id="86dcb-213">Обычно, когда клиент открывает Excel документа с веб-сайта, он ожидает, что документ будет содержать некоторые данные с веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="86dcb-213">Typically, when the customer opens up an Excel document from your web site, they expect the document to contain some data from the web site.</span></span> <span data-ttu-id="86dcb-214">Существует несколько способов записи данных в документ.</span><span class="sxs-lookup"><span data-stu-id="86dcb-214">There are a couple of ways to write data into the document.</span></span>

- <span data-ttu-id="86dcb-215">**Для записи данных используйте SDK OOXML.**</span><span class="sxs-lookup"><span data-stu-id="86dcb-215">**Use the OOXML SDK to write the data**.</span></span> <span data-ttu-id="86dcb-216">Вы можете использовать SDK для непосредственного записи любых данных в документ.</span><span class="sxs-lookup"><span data-stu-id="86dcb-216">You can use the SDK to directly write any data into the document.</span></span> <span data-ttu-id="86dcb-217">Этот подход полезен, если вы хотите, чтобы данные были доступны сразу после открытия документа.</span><span class="sxs-lookup"><span data-stu-id="86dcb-217">This approach is useful if you want the data to be available the instant the document is opened.</span></span>
- <span data-ttu-id="86dcb-218">**Передай свойство настраиваемого запроса Office надстройки.**</span><span class="sxs-lookup"><span data-stu-id="86dcb-218">**Pass a custom query property to your Office Add-in**.</span></span> <span data-ttu-id="86dcb-219">При генерации документа встраив настраиваемую свойство для надстройки Office, содержаную строку запроса, которая извлекает все необходимые данные.</span><span class="sxs-lookup"><span data-stu-id="86dcb-219">When you generate the document, you embed a custom property for the Office Add-in that contains a query string that retrieves all the required data.</span></span> <span data-ttu-id="86dcb-220">Когда надстройка открывается, она извлекает запрос, запускает запрос и использует API Office JS, чтобы вставить результат запроса в документ.</span><span class="sxs-lookup"><span data-stu-id="86dcb-220">When your add-in opens, it retrieves the query, runs the query, and uses the Office JS API to insert the result of the query into the document.</span></span>

### <a name="working-with-the-ooxml-sdk"></a><span data-ttu-id="86dcb-221">Работа с SDK OOXML</span><span class="sxs-lookup"><span data-stu-id="86dcb-221">Working with the OOXML SDK</span></span>

<span data-ttu-id="86dcb-222">SDK OOXML основан на .NET.</span><span class="sxs-lookup"><span data-stu-id="86dcb-222">The OOXML SDK is based on .NET.</span></span> <span data-ttu-id="86dcb-223">Если в вашем веб-приложении нет .NET, необходимо искать альтернативный способ работы с OOXML.</span><span class="sxs-lookup"><span data-stu-id="86dcb-223">If your web application does not .NET, you’ll need to look for an alternative way to work with OOXML.</span></span>

<span data-ttu-id="86dcb-224">Существует версия JavaScript SDK OOXML, доступная в [Open XML SDK для JavaScript.](https://archive.codeplex.com/?p=openxmlsdkjs)</span><span class="sxs-lookup"><span data-stu-id="86dcb-224">There is a JavaScript version of the OOXML SDK available at [Open XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).</span></span>

<span data-ttu-id="86dcb-225">Код OOXML можно разместить в функции Azure, чтобы отделить код .NET от остальной части веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="86dcb-225">You can place the OOXML code in an Azure function to separate the .NET code from the rest of your web application.</span></span> <span data-ttu-id="86dcb-226">Затем вызывайте функцию Azure (для создания Excel документа) из веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="86dcb-226">Then call the Azure function (to generate the Excel document) from your Web application.</span></span> <span data-ttu-id="86dcb-227">Дополнительные сведения о функциях Azure см. [в предисловии к Azure Functions.](/azure/azure-functions/functions-overview)</span><span class="sxs-lookup"><span data-stu-id="86dcb-227">For more information on Azure functions, see [An introduction to Azure Functions](/azure/azure-functions/functions-overview).</span></span>

### <a name="use-single-sign-on"></a><span data-ttu-id="86dcb-228">Использование единого входного</span><span class="sxs-lookup"><span data-stu-id="86dcb-228">Use single sign-on</span></span>

<span data-ttu-id="86dcb-229">Чтобы упростить проверку подлинности, рекомендуется, чтобы надстройка реализовала один вход.</span><span class="sxs-lookup"><span data-stu-id="86dcb-229">To simplify authentication, we recommend your add-in implements single sign-on.</span></span> <span data-ttu-id="86dcb-230">Дополнительные сведения см. в [документе Enable single sign-on for Office надстройки](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="86dcb-230">For more information, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)</span></span>

## <a name="see-also"></a><span data-ttu-id="86dcb-231">См. также</span><span class="sxs-lookup"><span data-stu-id="86dcb-231">See also</span></span>

- [<span data-ttu-id="86dcb-232">Добро пожаловать на страницу пакета SDK 2.5 Open XML для Office</span><span class="sxs-lookup"><span data-stu-id="86dcb-232">Welcome to the Open XML SDK 2.5 for Office</span></span>](/office/open-xml/open-xml-sdk)
- [<span data-ttu-id="86dcb-233">Автоматическое открытие области задач с документом</span><span class="sxs-lookup"><span data-stu-id="86dcb-233">Automatically open a task pane with a document</span></span>](../develop/automatically-open-a-task-pane-with-a-document.md)
- [<span data-ttu-id="86dcb-234">Persisting add-in state and settings</span><span class="sxs-lookup"><span data-stu-id="86dcb-234">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="86dcb-235">Создайте документ электронной таблицы, указав имя файла</span><span class="sxs-lookup"><span data-stu-id="86dcb-235">Create a spreadsheet document by providing a file name</span></span>](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)