---
title: Надстройки PowerPoint
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5c605410601d711e28ca04ff6e26387019cbb41
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925320"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="8a932-102">Надстройки PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8a932-102">PowerPoint add-ins</span></span>

<span data-ttu-id="8a932-p101">С помощью надстроек PowerPoint можно создавать удобные решения, подходящие для использования в презентациях на различных платформах, таких как Windows, iOS, Office Online и Mac. Можно создать два типа надстроек:</span><span class="sxs-lookup"><span data-stu-id="8a932-p101">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create one of two types of add-ins:</span></span>

- <span data-ttu-id="8a932-p102">**Контентные надстройки** позволяют добавлять динамический контент HTML5 в презентации. Например, ознакомьтесь с надстройкой [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false), с помощью которой можно добавить интерактивные схемы LucidChart в набор слайдов.</span><span class="sxs-lookup"><span data-stu-id="8a932-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>
- <span data-ttu-id="8a932-p103">**Надстройки области задач** позволяют добавлять справочные сведения или данные в слайд с помощью службы. Например, используя надстройку [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), вы можете вставить профессиональные фотографии в свою презентацию.</span><span class="sxs-lookup"><span data-stu-id="8a932-p103">Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="8a932-109">Сценарии надстроек PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8a932-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="8a932-110">В приведенных в статье примерах кода показаны основные задачи по разработке контентных надстроек для PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="8a932-110">The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint.</span></span> 

<span data-ttu-id="8a932-p104">При отображении сведений эти примеры зависят от функции `app.showNotification`, включенной в шаблоны проектов надстроек Office в Visual Studio. Если для разработки надстройки вы не используете Visual Studio, замените функцию `showNotification` собственным кодом. Некоторые из этих примеров также зависят от объекта `globals`, объявленного за пределами указанных функций: `var globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="8a932-p104">To display information, these examples depend on the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code. Several of these examples also depend on this `globals` object that is declared outside of the scope of these functions: `var globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

<span data-ttu-id="8a932-114">Эти примеры кода требуют, чтобы проект [ссылался на библиотеку Office.js 1.1 или более поздней версии](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="8a932-114">These code examples require your project to [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="8a932-115">Определение активного представления презентации и обработка события ActiveViewChanged</span><span class="sxs-lookup"><span data-stu-id="8a932-115">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="8a932-116">При создании контентной надстройки вам понадобится получить активное представление презентации, а также обработать событие ActiveViewChanged в рамках обработчика событий Office.Initialize.</span><span class="sxs-lookup"><span data-stu-id="8a932-116">If you are building a content add-in, you will need to get the presentation's active view and handle the ActiveViewChanged event, as part of your Office.Initialize handler.</span></span>


- <span data-ttu-id="8a932-117">Функция `getActiveFileView` вызывает метод [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync), который возвращает текущее представление презентации: "edit" (представления, в которых можно редактировать слайды, например  **Обычный режим** или **Режим структуры**) или "read" (**Показ слайдов** или **Режим чтения**).</span><span class="sxs-lookup"><span data-stu-id="8a932-117">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" ( **Slide Show** or **Reading View**) view.</span></span>


- <span data-ttu-id="8a932-118">Функция `registerActiveViewChanged` вызывает метод [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) для регистрации обработчика для события [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged).</span><span class="sxs-lookup"><span data-stu-id="8a932-118">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) method to register a handler for the [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event.</span></span> 

> [!NOTE]
> <span data-ttu-id="8a932-p105">В PowerPoint Online не удастся запустить событие [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged), поскольку режим показа слайдов обрабатывается как новый сеанс. В этом случае надстройке необходимо получить активное представление по загрузке, как указано ниже.</span><span class="sxs-lookup"><span data-stu-id="8a932-p105">In PowerPoint Online, the [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as noted below.</span></span>

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
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="8a932-121">Переход к определенному слайду презентации</span><span class="sxs-lookup"><span data-stu-id="8a932-121">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="8a932-p106">Функция `getSelectedRange` вызывает метод [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), чтобы получить объект JSON, возвращаемый свойством  `asyncResult.value` и который включает в себя массив с именем slides, содержащий идентификаторы, заголовки и индексы выбранного диапазона слайдов (или текущего слайда). Кроме того, он сохраняет идентификатор первого слайда в выбранном диапазоне в глобальной переменной.</span><span class="sxs-lookup"><span data-stu-id="8a932-p106">The  `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) method to get a JSON object returned by `asyncResult.value`, which contains an array named "slides" that contains the ids, titles, and indexes of selected range of slides (or just the current slide). It also saves the id of the first slide in the selected range to a global variable.</span></span>


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

<span data-ttu-id="8a932-124">Функция `goToFirstSlide` вызывает метод [Document.goToByIdAsync](https://dev.office.com/reference/add-ins/shared/document.gotobyidasync) для перехода к идентификатору первого слайда, сохраненному описанной выше функцией `getSelectedRange`.</span><span class="sxs-lookup"><span data-stu-id="8a932-124">The  `goToFirstSlide` function calls the [Document.goToByIdAsync](https://dev.office.com/reference/add-ins/shared/document.gotobyidasync) method to go to the id of the first slide stored by the `getSelectedRange` function above.</span></span>




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


## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="8a932-125">Переход между слайдами презентации</span><span class="sxs-lookup"><span data-stu-id="8a932-125">Navigate between slides in the presentation</span></span>

<span data-ttu-id="8a932-126">Функция `goToSlideByIndex` вызывает метод **Document.goToByIdAsync** для перехода к следующему слайду в презентации.</span><span class="sxs-lookup"><span data-stu-id="8a932-126">The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>


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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="8a932-127">Получение URL-адреса презентации</span><span class="sxs-lookup"><span data-stu-id="8a932-127">Get the URL of the presentation</span></span>

<span data-ttu-id="8a932-128">Функция `getFileUrl` вызывает метод [Document.getFileProperties](https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync), чтобы получить URL-адрес файла презентации.</span><span class="sxs-lookup"><span data-stu-id="8a932-128">The  `getFileUrl` function calls the [Document.getFileProperties](https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) method to get the URL of the presentation file.</span></span>


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



## <a name="see-also"></a><span data-ttu-id="8a932-129">См. также</span><span class="sxs-lookup"><span data-stu-id="8a932-129">See also</span></span>
- [<span data-ttu-id="8a932-130">Примеры кода PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8a932-130">PowerPoint Code Samples</span></span>](https://dev.office.com/code-samples#?filters=powerpoint)
- [<span data-ttu-id="8a932-131">Сохранение состояния надстройки и параметров документа для надстроек области задач и контентных надстроек</span><span class="sxs-lookup"><span data-stu-id="8a932-131">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="8a932-132">Чтение и запись данных при активном выделении фрагмента в документе или электронной таблице</span><span class="sxs-lookup"><span data-stu-id="8a932-132">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="8a932-133">Получение всего документа из надстройки для PowerPoint или Word</span><span class="sxs-lookup"><span data-stu-id="8a932-133">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="8a932-134">Использование тем документов в надстройках PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8a932-134">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
