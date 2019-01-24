---
title: Надстройки PowerPoint
description: ''
ms.date: 10/16/2018
localization_priority: Priority
ms.openlocfilehash: 022bed349dde061b61a8db0711a94a0a4d77f2e1
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388635"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="c8449-102">Надстройки PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c8449-102">PowerPoint add-ins</span></span>

<span data-ttu-id="c8449-103">С помощью надстроек PowerPoint можно создавать удобные решения, подходящие для использования в презентациях на различных платформах, таких как Windows, iOS, Office Online и Mac.</span><span class="sxs-lookup"><span data-stu-id="c8449-103">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac.</span></span> <span data-ttu-id="c8449-104">Можно создать два типа надстроек PowerPoint:</span><span class="sxs-lookup"><span data-stu-id="c8449-104">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="c8449-p102">**Контентные надстройки** позволяют добавлять динамический контент HTML5 в презентации. Например, ознакомьтесь с надстройкой [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false), с помощью которой можно добавить интерактивные схемы LucidChart в набор слайдов.</span><span class="sxs-lookup"><span data-stu-id="c8449-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="c8449-107">**Надстройки области задач** позволяют добавлять справочные сведения или данные в презентацию с помощью службы.</span><span class="sxs-lookup"><span data-stu-id="c8449-107">Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service.</span></span> <span data-ttu-id="c8449-108">Например, используя надстройку [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), вы можете вставить профессиональные фотографии в свою презентацию.</span><span class="sxs-lookup"><span data-stu-id="c8449-108">For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="c8449-109">Сценарии надстроек PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c8449-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="c8449-110">В приведенных в этой статье примерах кода показаны основные задачи по разработке надстроек для PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="c8449-110">The code examples in this article demonstrate some basic tasks for developing add-ins for PowerPoint.</span></span> <span data-ttu-id="c8449-111">Обратите внимание на следующее:</span><span class="sxs-lookup"><span data-stu-id="c8449-111">Please note the following:</span></span>

- <span data-ttu-id="c8449-112">При отображении сведений эти примеры используют функцию `app.showNotification`, включенную в шаблоны проектов надстроек Office в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="c8449-112">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="c8449-113">Если для разработки надстройки вы не используете Visual Studio, замените функцию `showNotification` собственным кодом.</span><span class="sxs-lookup"><span data-stu-id="c8449-113">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span> 

- <span data-ttu-id="c8449-114">Некоторые из этих примеров также используют объект `Globals`, объявленный за пределами указанных функций как `var Globals = {activeViewHandler:0, firstSlideId:0};`.</span><span class="sxs-lookup"><span data-stu-id="c8449-114">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="c8449-115">Для использования этих примеров необходимо, чтобы проект надстройки [ссылался на библиотеку Office.js 1.1 или более поздней версии](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="c8449-115">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="c8449-116">Определение активного представления презентации и обработка события ActiveViewChanged</span><span class="sxs-lookup"><span data-stu-id="c8449-116">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="c8449-117">При создании контентной надстройки вам понадобится получить активное представление презентации, а также обработать событие `ActiveViewChanged` в рамках обработчика событий `Office.Initialize`.</span><span class="sxs-lookup"><span data-stu-id="c8449-117">If you are building a content add-in, you will need to get the presentation's active view and handle the `ActiveViewChanged` event, as part of your `Office.Initialize` handler.</span></span> 

> [!NOTE]
> <span data-ttu-id="c8449-118">В PowerPoint Online не удастся запустить событие [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document), поскольку режим показа слайдов обрабатывается как новый сеанс.</span><span class="sxs-lookup"><span data-stu-id="c8449-118">In PowerPoint Online, the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document) event will never fire as Slide Show mode is treated as a new session.</span></span> <span data-ttu-id="c8449-119">В этом случае надстройке необходимо получить активное представление по загрузке, как показано в примере кода ниже.</span><span class="sxs-lookup"><span data-stu-id="c8449-119">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="c8449-120">В представленном ниже примере кода:</span><span class="sxs-lookup"><span data-stu-id="c8449-120">In the following code sample:</span></span>

- <span data-ttu-id="c8449-121">Функция `getActiveFileView` вызывает метод [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document#getactiveviewasync-options--callback-), который возвращает текущее представление презентации: "edit" (представления, в которых можно редактировать слайды, например  **Обычный режим** или **Режим структуры**) или "read" (**Показ слайдов** или **Режим чтения**).</span><span class="sxs-lookup"><span data-stu-id="c8449-121">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" (**Slide Show** or **Reading View**).</span></span>

- <span data-ttu-id="c8449-122">Функция `registerActiveViewChanged` вызывает метод [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) для регистрации обработчика для события [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document).</span><span class="sxs-lookup"><span data-stu-id="c8449-122">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document) event.</span></span> 


```js
//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}

function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                app.showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                app.showNotification(asyncResult.status);
            }
        });
}
```

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="c8449-123">Переход к определенному слайду презентации</span><span class="sxs-lookup"><span data-stu-id="c8449-123">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="c8449-124">В приведенном ниже примере кода функция `getSelectedRange` вызывает метод [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-), чтобы получить возвращаемый свойством `asyncResult.value` объект JSON, содержащий массив с именем **slides**.</span><span class="sxs-lookup"><span data-stu-id="c8449-124">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named **slides**.</span></span> <span data-ttu-id="c8449-125">Массив **slides** содержит идентификаторы, заголовки и индексы выбранного диапазона слайдов (или текущего слайда, если не выбрано несколько слайдов).</span><span class="sxs-lookup"><span data-stu-id="c8449-125">The **slides** array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="c8449-126">Кроме того, он сохраняет идентификатор первого слайда в выбранном диапазоне в глобальной переменной.</span><span class="sxs-lookup"><span data-stu-id="c8449-126">It also saves the id of the first slide in the selected range to a global variable.</span></span>

```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

<span data-ttu-id="c8449-127">В приведенном ниже примере кода функция `goToFirstSlide` вызывает метод [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) для перехода к первому слайду, который был определен показанной ранее функцией `getSelectedRange`.</span><span class="sxs-lookup"><span data-stu-id="c8449-127">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="c8449-128">Переход между слайдами презентации</span><span class="sxs-lookup"><span data-stu-id="c8449-128">Navigate between slides in the presentation</span></span>

<span data-ttu-id="c8449-129">В приведенном ниже примере кода функция `goToSlideByIndex` вызывает метод **Document.goToByIdAsync** для перехода к следующему слайду в презентации.</span><span class="sxs-lookup"><span data-stu-id="c8449-129">In the following code sample, the `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>

```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="c8449-130">Получение URL-адреса презентации</span><span class="sxs-lookup"><span data-stu-id="c8449-130">Get the URL of the presentation</span></span>

<span data-ttu-id="c8449-131">В приведенном ниже примере кода функция `getFileUrl` вызывает метод [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document#getfilepropertiesasync-options--callback-), чтобы получить URL-адрес файла презентации.</span><span class="sxs-lookup"><span data-stu-id="c8449-131">In the following code sample, the  `getFileUrl` function calls the [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```



## <a name="see-also"></a><span data-ttu-id="c8449-132">См. также</span><span class="sxs-lookup"><span data-stu-id="c8449-132">See also</span></span>
- <span data-ttu-id="c8449-133">
  [Примеры кода PowerPoint](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="c8449-133">[PowerPoint Code Samples](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)</span></span>
- [<span data-ttu-id="c8449-134">Сохранение состояния надстройки и параметров документа для надстроек области задач и контентных надстроек</span><span class="sxs-lookup"><span data-stu-id="c8449-134">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="c8449-135">Чтение и запись данных при активном выделении фрагмента в документе или электронной таблице</span><span class="sxs-lookup"><span data-stu-id="c8449-135">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="c8449-136">Получение всего документа из надстройки для PowerPoint или Word</span><span class="sxs-lookup"><span data-stu-id="c8449-136">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="c8449-137">Использование тем документов в надстройках PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c8449-137">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
