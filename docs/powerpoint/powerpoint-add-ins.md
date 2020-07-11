---
title: Надстройки PowerPoint
description: Узнайте, как использовать надстройки PowerPoint для создания удобных решений для презентаций на разных платформах, включая Windows, iPad, Mac и в браузере.
ms.date: 06/29/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 314b441f3d4b6d2188ed630fe2b254aec42a86bc
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006453"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="bcdd1-103">Надстройки PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bcdd1-103">PowerPoint add-ins</span></span>

<span data-ttu-id="bcdd1-104">С помощью надстроек PowerPoint можно создавать удобные решения, подходящие для использования в презентациях на различных платформах, таких как Windows, iPad, Mac и браузеры.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-104">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iPad, Mac, and in a browser.</span></span> <span data-ttu-id="bcdd1-105">Можно создать два типа надстроек PowerPoint:</span><span class="sxs-lookup"><span data-stu-id="bcdd1-105">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="bcdd1-106">Use **content add-ins** to add dynamic HTML5 content to your presentations.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-106">Use **content add-ins** to add dynamic HTML5 content to your presentations.</span></span> <span data-ttu-id="bcdd1-107">For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-107">For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="bcdd1-108">**Надстройки области задач** позволяют добавлять справочные сведения или данные в презентацию с помощью службы.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-108">Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service.</span></span> <span data-ttu-id="bcdd1-109">Например, используя надстройку [Pexels - Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997), вы можете вставить профессиональные фотографии в свою презентацию.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-109">For example, see the [Pexels - Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997) add-in, which you can use to add professional photos to your presentation.</span></span>

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="bcdd1-110">Сценарии надстроек PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bcdd1-110">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="bcdd1-111">В приведенных в этой статье примерах кода показаны основные задачи по разработке надстроек для PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-111">The code examples in this article demonstrate some basic tasks for developing add-ins for PowerPoint.</span></span> <span data-ttu-id="bcdd1-112">Обратите внимание на следующее:</span><span class="sxs-lookup"><span data-stu-id="bcdd1-112">Please note the following:</span></span>

- <span data-ttu-id="bcdd1-113">При отображении сведений эти примеры используют функцию `app.showNotification`, включенную в шаблоны проектов надстроек Office в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-113">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="bcdd1-114">Если для разработки надстройки вы не используете Visual Studio, замените функцию `showNotification` собственным кодом.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-114">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span>

- <span data-ttu-id="bcdd1-115">Некоторые из этих примеров также используют объект `Globals`, объявленный за пределами указанных функций как `var Globals = {activeViewHandler:0, firstSlideId:0};`.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-115">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="bcdd1-116">Для использования этих примеров необходимо, чтобы проект надстройки [ссылался на библиотеку Office.js 1.1 или более поздней версии](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="bcdd1-116">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="bcdd1-117">Определение активного представления презентации и обработка события ActiveViewChanged</span><span class="sxs-lookup"><span data-stu-id="bcdd1-117">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="bcdd1-118">При создании контентной надстройки вам понадобится получить активное представление презентации, а также обработать событие `ActiveViewChanged` в рамках обработчика событий `Office.Initialize`.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-118">If you are building a content add-in, you will need to get the presentation's active view and handle the `ActiveViewChanged` event, as part of your `Office.Initialize` handler.</span></span>

> [!NOTE]
> <span data-ttu-id="bcdd1-119">В PowerPoint в Интернете не удастся запустить событие [Document.ActiveViewChanged](/javascript/api/office/office.document), поскольку режим показа слайдов обрабатывается как новый сеанс.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-119">In PowerPoint on the web, the [Document.ActiveViewChanged](/javascript/api/office/office.document) event will never fire as Slide Show mode is treated as a new session.</span></span> <span data-ttu-id="bcdd1-120">В этом случае надстройке необходимо получить активное представление по загрузке, как показано в примере кода ниже.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-120">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="bcdd1-121">В представленном ниже примере кода:</span><span class="sxs-lookup"><span data-stu-id="bcdd1-121">In the following code sample:</span></span>

- <span data-ttu-id="bcdd1-122">Функция `getActiveFileView` вызывает метод [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-), который возвращает текущее представление презентации: "edit" (представления, в которых можно редактировать слайды, например  **Обычный режим** или **Режим структуры**) или "read" (**Показ слайдов** или **Режим чтения**).</span><span class="sxs-lookup"><span data-stu-id="bcdd1-122">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" (**Slide Show** or **Reading View**).</span></span>

- <span data-ttu-id="bcdd1-123">Функция `registerActiveViewChanged` вызывает метод [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) для регистрации обработчика для события [Document.ActiveViewChanged](/javascript/api/office/office.document).</span><span class="sxs-lookup"><span data-stu-id="bcdd1-123">The  `registerActiveViewChanged` function calls the [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](/javascript/api/office/office.document) event.</span></span>


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="bcdd1-124">Переход к определенному слайду презентации</span><span class="sxs-lookup"><span data-stu-id="bcdd1-124">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="bcdd1-125">В следующем примере кода функция `getSelectedRange` вызывает метод [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) для получения объекта JSON, возвращаемого свойством `asyncResult.value`. Этот объект содержит массив с именем `slides`.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-125">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named `slides`.</span></span> <span data-ttu-id="bcdd1-126">Массив `slides` содержит идентификаторы, заголовки и индексы выбранного диапазона слайдов (или текущего слайда, если не выбрано несколько слайдов).</span><span class="sxs-lookup"><span data-stu-id="bcdd1-126">The `slides` array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="bcdd1-127">Кроме того, он сохраняет идентификатор первого слайда в выбранном диапазоне в глобальной переменной.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-127">It also saves the id of the first slide in the selected range to a global variable.</span></span>

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

<span data-ttu-id="bcdd1-128">В приведенном ниже примере кода функция `goToFirstSlide` вызывает метод [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) для перехода к первому слайду, который был определен показанной ранее функцией `getSelectedRange`.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-128">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

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

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="bcdd1-129">Переход между слайдами презентации</span><span class="sxs-lookup"><span data-stu-id="bcdd1-129">Navigate between slides in the presentation</span></span>

<span data-ttu-id="bcdd1-130">В следующем примере кода функция `goToSlideByIndex` вызывает метод `Document.goToByIdAsync` для перехода к следующему слайду презентации.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-130">In the following code sample, the `goToSlideByIndex` function calls the `Document.goToByIdAsync` method to navigate to the next slide in the presentation.</span></span>

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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="bcdd1-131">Получение URL-адреса презентации</span><span class="sxs-lookup"><span data-stu-id="bcdd1-131">Get the URL of the presentation</span></span>

<span data-ttu-id="bcdd1-132">В приведенном ниже примере кода функция `getFileUrl` вызывает метод [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-), чтобы получить URL-адрес файла презентации.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-132">In the following code sample, the  `getFileUrl` function calls the [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

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

## <a name="create-a-presentation"></a><span data-ttu-id="bcdd1-133">Создание презентации</span><span class="sxs-lookup"><span data-stu-id="bcdd1-133">Create a presentation</span></span>

<span data-ttu-id="bcdd1-134">Ваша надстройка может создать новую презентацию, отдельную от экземпляра PowerPoint, в котором в настоящее время работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-134">Your add-in can create a new presentation, separate from the PowerPoint instance in which the add-in is currently running.</span></span> <span data-ttu-id="bcdd1-135">Для этой цели в пространстве имен PowerPoint есть метод `createPresentation`.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-135">The PowerPoint namespace has the `createPresentation` method for this purpose.</span></span> <span data-ttu-id="bcdd1-136">При вызове этого метода сразу открывается и отображается новая презентация в новом экземпляре программы PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-136">When this method is called, the new presentation is immediately opened and displayed in a new instance of PowerPoint.</span></span> <span data-ttu-id="bcdd1-137">Ваша надстройка остается открытой и запущенной в предыдущей презентации.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-137">Your add-in remains open and running with the previous presentation.</span></span>

```js
PowerPoint.createPresentation();
```

<span data-ttu-id="bcdd1-138">С помощью метода `createPresentation` также можно создать копию существующей презентации.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-138">The `createPresentation` method can also create a copy of an existing presentation.</span></span> <span data-ttu-id="bcdd1-139">Метод принимает в качестве необязательного параметра строковое представление PPTX-файла в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-139">The method accepts a base64-encoded string representation of an .pptx file as an optional parameter.</span></span> <span data-ttu-id="bcdd1-140">Полученная презентация будет копией этого файла, предполагая, что строковый аргумент является допустимым PPTX-файлом.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-140">The resulting presentation will be a copy of that file, assuming the string argument is a valid .pptx file.</span></span> <span data-ttu-id="bcdd1-141">Преобразование файла в нужную строку в кодировке base64 можно выполнить с помощью класса [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader), как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="bcdd1-141">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = reader.result.toString().indexOf("base64,");
    var copyBase64 = reader.result.toString().substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a><span data-ttu-id="bcdd1-142">См. также</span><span class="sxs-lookup"><span data-stu-id="bcdd1-142">See also</span></span>

- [<span data-ttu-id="bcdd1-143">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bcdd1-143">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="bcdd1-144">Примеры кода PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bcdd1-144">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="bcdd1-145">Сохранение состояния надстройки и параметров документа для надстроек области задач и контентных надстроек</span><span class="sxs-lookup"><span data-stu-id="bcdd1-145">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="bcdd1-146">Чтение и запись данных при активном выделении фрагмента в документе или электронной таблице</span><span class="sxs-lookup"><span data-stu-id="bcdd1-146">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="bcdd1-147">Получение всего документа из надстройки для PowerPoint или Word</span><span class="sxs-lookup"><span data-stu-id="bcdd1-147">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="bcdd1-148">Использование тем документов в надстройках PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bcdd1-148">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
