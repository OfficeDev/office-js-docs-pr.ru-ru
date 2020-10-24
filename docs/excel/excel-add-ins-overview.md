---
title: Общие сведения о надстройках Excel
description: Надстройка Excel позволяет расширять функциональность приложения Excel на нескольких платформах, включая Windows, Mac, iPad и в браузере.
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6ecb702649a3cef0f3c3acb8432d3f8d3db65115
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740856"
---
# <a name="excel-add-ins-overview"></a><span data-ttu-id="cb635-103">Общие сведения о надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="cb635-103">Excel add-ins overview</span></span>

<span data-ttu-id="cb635-p101">С помощью надстройки Excel можно расширить возможности приложения Excel на различных платформах, в том числе Windows, Mac, iPad и в браузере. Используйте надстройки в книге Excel, чтобы:</span><span class="sxs-lookup"><span data-stu-id="cb635-p101">An Excel add-in allows you to extend Excel application functionality across multiple platforms including Windows, Mac, iPad, and in a browser. Use Excel add-ins within a workbook to:</span></span>

- <span data-ttu-id="cb635-106">взаимодействовать с объектами Excel, считывать и записывать данные Excel;</span><span class="sxs-lookup"><span data-stu-id="cb635-106">Interact with Excel objects, read and write Excel data.</span></span>
- <span data-ttu-id="cb635-107">расширять возможности с помощью области задач или области содержимого;</span><span class="sxs-lookup"><span data-stu-id="cb635-107">Extend functionality using web based task pane or content pane</span></span>
- <span data-ttu-id="cb635-108">добавлять настраиваемые кнопки ленты или элементы контекстного меню;</span><span class="sxs-lookup"><span data-stu-id="cb635-108">Add custom ribbon buttons or contextual menu items</span></span>
- <span data-ttu-id="cb635-109">добавлять пользовательские функции;</span><span class="sxs-lookup"><span data-stu-id="cb635-109">Add custom functions</span></span>
- <span data-ttu-id="cb635-110">расширять возможности взаимодействия с помощью диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="cb635-110">Provide richer interaction using dialog window</span></span>

<span data-ttu-id="cb635-111">В качестве основы используется платформа надстроек Office, предоставляющая API JavaScript для Office.js, с помощью которых можно создавать и запускать надстройки Excel. Используя платформу надстроек Office для создания надстройки Excel, вы получаете следующие преимущества:</span><span class="sxs-lookup"><span data-stu-id="cb635-111">The Office Add-ins platform provides the framework and Office.js JavaScript APIs that enable you to create and run Excel add-ins. By using the Office Add-ins platform to create your Excel add-in, you'll get the following benefits:</span></span>

* <span data-ttu-id="cb635-112">**Кроссплатформенная поддержка**. Надстройки Excel работают в Office в Интернете, Office для Windows, Office для Mac и Office для iPad.</span><span class="sxs-lookup"><span data-stu-id="cb635-112">**Cross-platform support**: Excel add-ins run in Office on the web, Windows, Mac, and iPad.</span></span>
* <span data-ttu-id="cb635-113">**Централизованное развертывание.** Администраторы могут легко и быстро развертывать надстройки Excel для пользователей в организации.</span><span class="sxs-lookup"><span data-stu-id="cb635-113">**Centralized deployment**: Admins can quickly and easily deploy Excel add-ins to users throughout an organization.</span></span>
* <span data-ttu-id="cb635-114">**Использование стандартных веб-технологий**. Создавайте надстройки Excel, используя знакомые веб-технологии — HTML, CSS и JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cb635-114">**Use of standard web technology**: Create your Excel add-in using familiar web technologies such as HTML, CSS, and JavaScript.</span></span>
* <span data-ttu-id="cb635-115">**Распространение через AppSource**. Представьте свою надстройку Excel широкой аудитории, опубликовав ее в [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d).</span><span class="sxs-lookup"><span data-stu-id="cb635-115">**Distribution via AppSource**: Share your Excel add-in with a broad audience by publishing it to [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d).</span></span>

> [!NOTE]
> <span data-ttu-id="cb635-p102">Надстройки Excel отличаются от надстроек COM и VSTO — устаревших решений для интеграции с Office, работающих только в Office для Windows. В отличие от надстроек COM, надстройки Excel не требуют установки какого-либо кода на устройстве пользователя или в Excel.</span><span class="sxs-lookup"><span data-stu-id="cb635-p102">Excel add-ins are different from COM and VSTO add-ins, which are earlier Office integration solutions that run only in Office on Windows. Unlike COM add-ins, Excel add-ins do not require you to install any code on a user's device, or within Excel.</span></span>

## <a name="components-of-an-excel-add-in"></a><span data-ttu-id="cb635-118">Компоненты надстройки Excel</span><span class="sxs-lookup"><span data-stu-id="cb635-118">Components of an Excel add-in</span></span>

<span data-ttu-id="cb635-119">Надстройка Excel включает два основных компонента: веб-приложение и файл конфигурации, называемый файлом манифеста.</span><span class="sxs-lookup"><span data-stu-id="cb635-119">An Excel add-in includes two basic components: a web application and a configuration file, called a manifest file.</span></span> 

<span data-ttu-id="cb635-p103">Веб-приложение использует [API JavaScript для Office](../reference/javascript-api-for-office.md) для взаимодействия с объектами в Excel, а также может упрощать работу с ресурсами в Интернете. Например, надстройка может выполнять следующие действия:</span><span class="sxs-lookup"><span data-stu-id="cb635-p103">The web application uses the [Office JavaScript API](../reference/javascript-api-for-office.md) to interact with objects in Excel, and can also facilitate interaction with online resources. For example, an add-in can perform any of the following tasks:</span></span>

* <span data-ttu-id="cb635-122">создавать, читать, обновлять и удалять данные в книге (листы, диапазоны, таблицы, диаграммы, именованные элементы и т. д.);</span><span class="sxs-lookup"><span data-stu-id="cb635-122">Create, read, update, and delete data in the workbook (worksheets, ranges, tables, charts, named items, and more).</span></span>
* <span data-ttu-id="cb635-123">выполнять авторизацию пользователя в веб-службе с помощью стандартного потока OAuth 2.0;</span><span class="sxs-lookup"><span data-stu-id="cb635-123">Perform user authorization with an online service by using the standard OAuth 2.0 flow.</span></span>
* <span data-ttu-id="cb635-124">отправлять запросы к API Microsoft Graph или другому API.</span><span class="sxs-lookup"><span data-stu-id="cb635-124">Issue API requests to Microsoft Graph or any other API.</span></span>

<span data-ttu-id="cb635-125">Веб-приложение может размещаться на любом веб-сервере, а для его создания можно использовать как клиентские платформы (например, Angular, React, jQuery), так и серверные технологии (например, ASP.NET, Node.js, PHP).</span><span class="sxs-lookup"><span data-stu-id="cb635-125">The web application can be hosted on any web server, and can be built using client-side frameworks (such as Angular, React, jQuery) or server-side technologies (such as ASP.NET, Node.js, PHP).</span></span>

<span data-ttu-id="cb635-126">[Манифест](../develop/add-in-manifests.md) — это XML-файл конфигурации, который определяет, как надстройка интегрируется с клиентами Office, указывая параметры и возможности, такие как:</span><span class="sxs-lookup"><span data-stu-id="cb635-126">The [manifest](../develop/add-in-manifests.md) is an XML configuration file that defines how the add-in integrates with Office clients by specifying settings and capabilities such as:</span></span>

* <span data-ttu-id="cb635-127">URL-адрес веб-приложения надстройки;</span><span class="sxs-lookup"><span data-stu-id="cb635-127">The URL of the add-in's web application.</span></span>
* <span data-ttu-id="cb635-128">отображаемое имя, описание, идентификатор, версию и языковой стандарт по умолчанию для надстройки;</span><span class="sxs-lookup"><span data-stu-id="cb635-128">The add-in's display name, description, ID, version, and default locale.</span></span>
* <span data-ttu-id="cb635-129">способ интеграции надстройки с Excel, включая настраиваемый пользовательский интерфейс, создаваемый надстройкой (кнопки ленты, контекстные меню и т. д.);</span><span class="sxs-lookup"><span data-stu-id="cb635-129">How the add-in integrates with Excel, including any custom UI that the add-in creates (ribbon buttons, context menus, and so on).</span></span>
* <span data-ttu-id="cb635-130">разрешения, необходимые надстройке, например чтение и запись документа.</span><span class="sxs-lookup"><span data-stu-id="cb635-130">Permissions that the add-in requires, such as reading and writing to the document.</span></span>

<span data-ttu-id="cb635-131">Чтобы пользователи могли устанавливать и использовать надстройку Excel, необходимо опубликовать ее манифест в AppSource или каталоге надстроек.</span><span class="sxs-lookup"><span data-stu-id="cb635-131">To enable end users to install and use an Excel add-in, you must publish its manifest either to AppSource or to an add-ins catalog.</span></span> <span data-ttu-id="cb635-132">Дополнительные сведения о публикации в AppSource см. в статье [Публикация решений в AppSource и в Office](/office/dev/store/submit-to-appsource-via-partner-center).</span><span class="sxs-lookup"><span data-stu-id="cb635-132">For details about publishing to AppSource, see [Make your solutions available in AppSource and within Office](/office/dev/store/submit-to-appsource-via-partner-center).</span></span>

## <a name="capabilities-of-an-excel-add-in"></a><span data-ttu-id="cb635-133">Возможности надстройки Excel</span><span class="sxs-lookup"><span data-stu-id="cb635-133">Capabilities of an Excel add-in</span></span>

<span data-ttu-id="cb635-134">Надстройки Excel могут не только взаимодействовать с содержимым книги, но и добавлять настраиваемые кнопки ленты и команды меню, вставлять области задач, добавлять пользовательские функции, открывать диалоговые окна и даже внедрять в лист многофункциональные веб-объекты, например диаграммы или интерактивные визуализации.</span><span class="sxs-lookup"><span data-stu-id="cb635-134">In addition to interacting with the content in the workbook, Excel add-ins can add custom ribbon buttons or menu commands, insert task panes, add custom functions, open dialog boxes, and even embed rich, web-based objects such as charts or interactive visualizations within a worksheet.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="cb635-135">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="cb635-135">Add-in commands</span></span>

<span data-ttu-id="cb635-p105">Команды надстройки — это элементы пользовательского интерфейса, расширяющие возможности пользовательского интерфейса Excel по умолчанию и запускающие действия в надстройке. С помощью команд надстроек можно добавить кнопку на ленту или пункт в контекстное меню в Excel. Когда пользователи выбирают команду надстройки, выполняется действие, например запуск кода JavaScript или отображение страницы надстройки на панели задач.</span><span class="sxs-lookup"><span data-stu-id="cb635-p105">Add-in commands are UI elements that extend the Excel UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu in Excel. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span></span> 

<span data-ttu-id="cb635-139">**Команды надстроек**</span><span class="sxs-lookup"><span data-stu-id="cb635-139">**Add-in commands**</span></span>

![Команды надстроек в Excel](../images/excel-add-in-commands-script-lab.png)

<span data-ttu-id="cb635-141">Дополнительные сведения о возможностях команд и поддерживаемых платформах, а также рекомендации по разработке команд надстроек см. в статье [Команды надстроек для Excel, Word и PowerPoint](../design/add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="cb635-141">For more information about command capabilities, supported platforms, and best practices for developing add-in commands, see [Add-in commands for Excel, Word, and PowerPoint](../design/add-in-commands.md).</span></span>

### <a name="task-panes"></a><span data-ttu-id="cb635-142">Области задач</span><span class="sxs-lookup"><span data-stu-id="cb635-142">Task panes</span></span>

<span data-ttu-id="cb635-p106">Области задач — это области в интерфейсе, которые обычно отображаются в правой части окна приложения Excel. В областях задач расположены элементы управления, с помощью которых запускается код для изменения документа Excel или отображения данных из источника данных.</span><span class="sxs-lookup"><span data-stu-id="cb635-p106">Task panes are interface surfaces that typically appear on the right side of the window within Excel. Task panes give users access to interface controls that run code to modify the Excel document or display data from a data source.</span></span> 

<span data-ttu-id="cb635-145">**Область задач**</span><span class="sxs-lookup"><span data-stu-id="cb635-145">**Task pane**</span></span>

![Надстройка области задач в Excel](../images/excel-add-in-task-pane-insights.png)

<span data-ttu-id="cb635-147">Дополнительные сведения об областях задач см. в статье [Области задач в надстройках Office](../design/task-pane-add-ins.md). Пример реализации области задач в Excel: [Тенденции расходов банка WoodGrove](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) на JS.</span><span class="sxs-lookup"><span data-stu-id="cb635-147">For more information about task panes, see [Task panes in Office Add-ins](../design/task-pane-add-ins.md). For a sample that implements a task pane in Excel, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends).</span></span>

### <a name="custom-functions"></a><span data-ttu-id="cb635-148">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="cb635-148">Custom functions</span></span>

<span data-ttu-id="cb635-149">Пользовательские функции позволяют разработчикам добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="cb635-149">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="cb635-150">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="cb635-150">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> 

<span data-ttu-id="cb635-151">**Пользовательская функция**</span><span class="sxs-lookup"><span data-stu-id="cb635-151">**Custom function**</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="cb635-152">Дополнительные сведения о пользовательских функциях см. в статье [Создание пользовательских функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="cb635-152">For more information about custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

### <a name="dialog-boxes"></a><span data-ttu-id="cb635-153">Диалоговые окна</span><span class="sxs-lookup"><span data-stu-id="cb635-153">Dialog boxes</span></span>

<span data-ttu-id="cb635-p108">Диалоговые окна — это поверхности, которые накладываются на активное окно приложения Excel. Например, с помощью диалоговых окон можно отображать страницы входа, которые невозможно открыть непосредственно в области задач, запрашивать подтверждение действий пользователем и размещать видео, которые могут не помещаться в области задач. Чтобы открывать диалоговые окна в надстройке Excel, используйте [API диалоговых окон](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="cb635-p108">Dialog boxes are surfaces that float above the active Excel application window. You can use dialog boxes for tasks such as displaying sign-in pages that can't be opened directly in a task pane, requesting that the user confirm an action, or hosting videos that might be too small if confined to a task pane. To open dialog boxes in your Excel add-in, use the [Dialog API](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="cb635-157">**Диалоговое окно**</span><span class="sxs-lookup"><span data-stu-id="cb635-157">**Dialog box**</span></span>

![Диалоговое окно надстройки в Excel](../images/excel-add-in-dialog-choose-number.png)

<span data-ttu-id="cb635-159">Дополнительные сведения о диалоговых окнах и соответствующем API см. в статьях [Диалоговые окна в надстройках Office](../design/dialog-boxes.md) и [Использование API диалоговых окон в надстройках Office](../develop/dialog-api-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="cb635-159">For more information about dialog boxes and the Dialog API, see [Dialog boxes in Office Add-ins](../design/dialog-boxes.md) and [Use the Dialog API in your Office Add-ins](../develop/dialog-api-in-office-add-ins.md).</span></span>

### <a name="content-add-ins"></a><span data-ttu-id="cb635-160">Контентные надстройки</span><span class="sxs-lookup"><span data-stu-id="cb635-160">Content add-ins</span></span>

<span data-ttu-id="cb635-p109">Контентные надстройки — это поверхности, которые можно внедрять непосредственно в документы Excel. С помощью контентных надстроек можно внедрять в лист многофункциональные веб-объекты, например диаграммы, визуализации данных и файлы мультимедиа, или предоставлять пользователям доступ к элементам управления интерфейса, выполняющим код для изменения документа Excel или отображения данных из источника. Используйте контентные надстройки, когда требуется внедрить функции непосредственно в документ.</span><span class="sxs-lookup"><span data-stu-id="cb635-p109">Content add-ins are surfaces that you can embed directly into Excel documents. You can use content add-ins to embed rich, web-based objects such as charts, data visualizations, or media into a worksheet or to give users access to interface controls that run code to modify the Excel document or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.</span></span>

<span data-ttu-id="cb635-164">**Контентная надстройка**</span><span class="sxs-lookup"><span data-stu-id="cb635-164">**Content add-in**</span></span>

![Контентная надстройка в Excel](../images/excel-add-in-content-map.png)

<span data-ttu-id="cb635-166">Дополнительные сведения о контентных надстройках см. в статье [Контентные надстройки Office](../design/content-add-ins.md). Пример контентной надстройки Excel: [Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="cb635-166">For more information about content add-ins, see [Content Office Add-ins](../design/content-add-ins.md). For a sample that implements a content add-in in Excel, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="javascript-apis-to-interact-with-workbook-content"></a><span data-ttu-id="cb635-167">API JavaScript для взаимодействия с содержимым книги</span><span class="sxs-lookup"><span data-stu-id="cb635-167">JavaScript APIs to interact with workbook content</span></span>

<span data-ttu-id="cb635-168">Надстройка Excel взаимодействует с объектами в Excel с помощью [API JavaScript для Office](../reference/javascript-api-for-office.md), включающего две объектных модели JavaScript:</span><span class="sxs-lookup"><span data-stu-id="cb635-168">An Excel add-in interacts with objects in Excel by using the [Office JavaScript API](../reference/javascript-api-for-office.md), which includes two JavaScript object models:</span></span>

* <span data-ttu-id="cb635-169">**API JavaScript для Excel**. Появившийся в Office 2016 [API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к листам, диапазонам, таблицам, диаграммам и другим объектам.</span><span class="sxs-lookup"><span data-stu-id="cb635-169">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed Excel objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="cb635-170">**Общие API**. Появившиеся в Office 2013 общие API позволяют получать доступ к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="cb635-170">**Common API**: Introduced with Office 2013, the Common API enables you to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span> <span data-ttu-id="cb635-171">Общий API предоставляет ограниченные возможности по взаимодействию с Excel, поэтому его можно использовать, если надстройка должна работать в Excel 2013.</span><span class="sxs-lookup"><span data-stu-id="cb635-171">Because the Common API does provide limited functionality for Excel interaction, you can use it if your add-in needs to run on Excel 2013.</span></span>

## <a name="next-steps"></a><span data-ttu-id="cb635-172">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="cb635-172">Next steps</span></span>

<span data-ttu-id="cb635-173">Приступите к [созданию своей первой надстройки Excel](../quickstarts/excel-quickstart-jquery.md).</span><span class="sxs-lookup"><span data-stu-id="cb635-173">Get started by [creating your first Excel add-in](../quickstarts/excel-quickstart-jquery.md).</span></span> <span data-ttu-id="cb635-174">Затем ознакомьтесь с [основными понятиями](excel-add-ins-core-concepts.md), связанными с созданием надстроек Excel.</span><span class="sxs-lookup"><span data-stu-id="cb635-174">Then, learn about the [core concepts](excel-add-ins-core-concepts.md) of building Excel add-ins.</span></span>

## <a name="see-also"></a><span data-ttu-id="cb635-175">См. также</span><span class="sxs-lookup"><span data-stu-id="cb635-175">See also</span></span>

- [<span data-ttu-id="cb635-176">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="cb635-176">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="cb635-177">Сведения о программе для разработчиков Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="cb635-177">Learn about Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
- [<span data-ttu-id="cb635-178">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="cb635-178">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="cb635-179">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="cb635-179">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="cb635-180">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="cb635-180">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
