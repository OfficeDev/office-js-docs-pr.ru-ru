---
title: Общие сведения об интерфейсе API JavaScript для Office
description: ''
ms.date: 06/21/2019
localization_priority: Priority
ms.openlocfilehash: 1954457b477472b8940841bb1ffe5954e49e01ec
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/16/2019
ms.locfileid: "37524236"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="5ae92-102">Общие сведения об интерфейсе API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="5ae92-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="5ae92-p101">В этой статье можно узнать об интерфейсе API JavaScript для Office и о том, как его использовать. Справочные сведения см. в разделе [API JavaScript для Office](/office/dev/add-ins/reference/javascript-api-for-office). О том, как обновить файлы проекта Visual Studio до последней версии API JavaScript для Office, см. в статье [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="5ae92-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="5ae92-p102">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="5ae92-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="5ae92-108">Ссылки на библиотеку API JavaScript для Office в надстройке</span><span class="sxs-lookup"><span data-stu-id="5ae92-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="5ae92-p103">Библиотека [API JavaScript для Office](/office/dev/add-ins/reference/javascript-api-for-office) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например Excel-15.js и Outlook-15.js. Простейший способ сослаться на API — использовать нашу сеть доставки содержимого (CDN), добавив следующий код `<script>` в тег `<head>` страницы:</span><span class="sxs-lookup"><span data-stu-id="5ae92-p103">The [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

<span data-ttu-id="5ae92-111">Это приведет к скачиванию и кэшированию файлов JavaScript API для Office при первой загрузке надстройки, чтобы убедиться, что она использует актуальную реализацию Office.js и сопутствующих файлов для указанной версии.</span><span class="sxs-lookup"><span data-stu-id="5ae92-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="5ae92-112">Подробные сведения о CDN Office.js, включая способы управления версиями и обратной совместимостью, см. в статье [Указание ссылок на библиотеку API JavaScript для Office из сети доставки содержимого (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="5ae92-112">For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="5ae92-113">Инициализация надстройки</span><span class="sxs-lookup"><span data-stu-id="5ae92-113">Initializing your add-in</span></span>

<span data-ttu-id="5ae92-114">**Область применения:** все типы надстроек</span><span class="sxs-lookup"><span data-stu-id="5ae92-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="5ae92-115">Надстройки Office часто поддерживают логику запуска для выполнения следующих действий:</span><span class="sxs-lookup"><span data-stu-id="5ae92-115">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="5ae92-116">Проверьте, что пользовательская версия Office будет поддерживать все API Office, которые вызывает ваш код.</span><span class="sxs-lookup"><span data-stu-id="5ae92-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="5ae92-117">Убедитесь в наличии определенных артефактов, таких как лист с заданным именем.</span><span class="sxs-lookup"><span data-stu-id="5ae92-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="5ae92-118">Попросите пользователя выделить некоторые ячейки в Excel, а затем вставить диаграмму с инициализацией этих выделенных значений.</span><span class="sxs-lookup"><span data-stu-id="5ae92-118">Prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="5ae92-119">Установите привязки.</span><span class="sxs-lookup"><span data-stu-id="5ae92-119">Establish bindings.</span></span>

- <span data-ttu-id="5ae92-120">Используйте API для диалогового окна Office, чтобы запрашивать у пользователя значения по умолчанию для параметров надстроек.</span><span class="sxs-lookup"><span data-stu-id="5ae92-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="5ae92-121">Но ваш код запуска не должен вызывать любой API Office.js, пока библиотека не будет загружена.</span><span class="sxs-lookup"><span data-stu-id="5ae92-121">But your start-up code must not call any Office.js APIs until the library is loaded.</span></span> <span data-ttu-id="5ae92-122">Существует два способа, с помощью которых ваш код может проверять, загружена ли библиотека.</span><span class="sxs-lookup"><span data-stu-id="5ae92-122">There are two ways that your code can ensure that the library is loaded.</span></span> <span data-ttu-id="5ae92-123">Они описаны в следующих разделах:</span><span class="sxs-lookup"><span data-stu-id="5ae92-123">They are described in the following sections:</span></span> 

- [<span data-ttu-id="5ae92-124">Инициализация с использованием Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="5ae92-124">Initialize with Office.onReady()</span></span>](#initialize-with-officeonready)
- [<span data-ttu-id="5ae92-125">Инициализация с использованием Office.initialize</span><span class="sxs-lookup"><span data-stu-id="5ae92-125">Initialize with Office.initialize</span></span>](#initialize-with-officeinitialize)

> [!TIP]
> <span data-ttu-id="5ae92-126">Рекомендуется использовать `Office.onReady()` вместо `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="5ae92-126">We recommend that you use `Office.onReady()` instead of `Office.initialize`.</span></span> <span data-ttu-id="5ae92-127">Хотя `Office.initialize` по-прежнему поддерживается, использование `Office.onReady()` обеспечивает дополнительную гибкость.</span><span class="sxs-lookup"><span data-stu-id="5ae92-127">Although `Office.initialize` is still supported, using `Office.onReady()` provides more flexibility.</span></span> <span data-ttu-id="5ae92-128">Вы можете назначить только один обработчик для `Office.initialize`, который будет вызываться только один раз инфраструктурой Office, но вы можете вызывать `Office.onReady()` в разных местах вашего кода и использовать разные обратные вызовы.</span><span class="sxs-lookup"><span data-stu-id="5ae92-128">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure, but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span>
> 
> <span data-ttu-id="5ae92-129">Сведения о различиях описанных ниже приемов см. в статье [Основные различия между Office.initialize и Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span><span class="sxs-lookup"><span data-stu-id="5ae92-129">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span>

<span data-ttu-id="5ae92-130">Дополнительные сведения о последовательности событий при инициализации надстройки см. в статье [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).</span><span class="sxs-lookup"><span data-stu-id="5ae92-130">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="5ae92-131">Инициализация с использованием Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="5ae92-131">Initialize with Office.onReady()</span></span>

<span data-ttu-id="5ae92-132">`Office.onReady()` — это асинхронный метод, который возвращает объект Promise во время проверки загрузки библиотеки Office.js.</span><span class="sxs-lookup"><span data-stu-id="5ae92-132">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is loaded.</span></span> <span data-ttu-id="5ae92-133">Когда библиотека будет загружена, объект Promise сопоставляется в качестве объекта, определяющего ведущее приложение Office, с числовым значением `Office.HostType` (`Excel`, `Word` и т. д.), а платформа — с числовым значением `Office.PlatformType` (`PC`, `Mac`, `OfficeOnline` и т. д.).</span><span class="sxs-lookup"><span data-stu-id="5ae92-133">When the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="5ae92-134">Объект Promise сопоставляется незамедлительно, если библиотека уже загружена, когда вызывается `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="5ae92-134">The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.</span></span>

<span data-ttu-id="5ae92-135">Один из способов вызова `Office.onReady()` состоит в передаче ему метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ae92-135">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="5ae92-136">Пример:</span><span class="sxs-lookup"><span data-stu-id="5ae92-136">Here's an example:</span></span>

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

<span data-ttu-id="5ae92-137">Кроме того, вы можете привязать метод `then()` к вызову `Office.onReady()`, вместо того чтобы использовать обратный вызов.</span><span class="sxs-lookup"><span data-stu-id="5ae92-137">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="5ae92-138">Например приведенный ниже код проверяет, поддерживает ли версия Excel пользователя использование API, которые может вызывать надстройка.</span><span class="sxs-lookup"><span data-stu-id="5ae92-138">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="5ae92-139">Вот аналогичный же пример, использующий ключевые слова `async` и `await` в TypeScript:</span><span class="sxs-lookup"><span data-stu-id="5ae92-139">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="5ae92-140">При использовании дополнительных платформ JavaScript, включающих собственный обработчик событий инициализации или тесты, они, *как правило*, должны размещаться внутри ответа для `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="5ae92-140">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="5ae92-141">Например, ссылка на [JQuery](https://jquery.com) функция `$(document).ready()` должна выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="5ae92-141">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="5ae92-142">Однако существуют исключения для таких случаев.</span><span class="sxs-lookup"><span data-stu-id="5ae92-142">However, there are exceptions to this practice.</span></span> <span data-ttu-id="5ae92-143">Предположим, например, что вы хотите открыть в браузере вашу надстройку (вместо того чтобы загружать ее в хост Office) для отладки вашего пользовательского интерфейса с помощью инструментов браузера.</span><span class="sxs-lookup"><span data-stu-id="5ae92-143">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="5ae92-144">Так как Office.js не загружается в браузер, `onReady` не будет работать, а `$(document).ready` не будет работать при вызове внутри Office `onReady`.</span><span class="sxs-lookup"><span data-stu-id="5ae92-144">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> <span data-ttu-id="5ae92-145">Другое исключение: вам потребуется индикатор выполнения, который должен отображаться в области задач при загрузке надстройки.</span><span class="sxs-lookup"><span data-stu-id="5ae92-145">Another exception: you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="5ae92-146">В данном сценарии ваш код должен вызывать jQuery `ready` и использовать ее обратный вызов для отображения индикатора выполнения.</span><span class="sxs-lookup"><span data-stu-id="5ae92-146">In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator.</span></span> <span data-ttu-id="5ae92-147">Затем обратный вызов `onReady` Office может заменять индикатор выполнения на окончательный пользовательский интерфейс </span><span class="sxs-lookup"><span data-stu-id="5ae92-147">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="5ae92-148">Инициализация с использованием Office.initialize</span><span class="sxs-lookup"><span data-stu-id="5ae92-148">Initialize with Office.initialize</span></span>

<span data-ttu-id="5ae92-149">Событие инициализации запускается, когда библиотека Office.js будет загружена и готова к взаимодействию с пользователем.</span><span class="sxs-lookup"><span data-stu-id="5ae92-149">An initialize event fires when the Office.js library is loaded and ready for user interaction.</span></span> <span data-ttu-id="5ae92-150">Вы можете назначить обработчик `Office.initialize` для реализации вашей логики инициализации.</span><span class="sxs-lookup"><span data-stu-id="5ae92-150">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="5ae92-151">Например, приведенный ниже код проверяет, поддерживает ли версия Excel пользователя использование API, которые может вызывать надстройка.</span><span class="sxs-lookup"><span data-stu-id="5ae92-151">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="5ae92-152">При использовании дополнительных платформ JavaScript, включающих собственные обработчики событий инициализации или тесты, они, *как правило*, должны размещаться внутри события `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="5ae92-152">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event.</span></span> <span data-ttu-id="5ae92-153">(Исключения, описанные в разделе**Инициализация с помощью Office.onReady()** ранее действуют и в этом случае). Например, на [ JQuery](https://jquery.com) функцию `$(document).ready()` нужно сослаться следующим образом:</span><span class="sxs-lookup"><span data-stu-id="5ae92-153">(But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="5ae92-154">Для надстроек области задач и контентных надстроек `Office.initialize` предоставляет дополнительный параметр _reason_.</span><span class="sxs-lookup"><span data-stu-id="5ae92-154">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="5ae92-155">Этот параметр определяет порядок добавления надстройки в текущий документ.</span><span class="sxs-lookup"><span data-stu-id="5ae92-155">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="5ae92-156">Это поможет обеспечить разную логику в тех случаях, когда надстройка вставляется впервые или когда она уже существует в документе.</span><span class="sxs-lookup"><span data-stu-id="5ae92-156">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

<span data-ttu-id="5ae92-157">Дополнительные сведения см. в статьях [Событие Office.initialize](/javascript/api/office) и [Перечисление InitializationReason](/javascript/api/office/office.initializationreason).</span><span class="sxs-lookup"><span data-stu-id="5ae92-157">For more information, see [Office.initialize Event](/javascript/api/office) and [InitializationReason Enumeration](/javascript/api/office/office.initializationreason).</span></span>

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="5ae92-158">Основные различия между Office.initialize и Office.onReady</span><span class="sxs-lookup"><span data-stu-id="5ae92-158">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="5ae92-159">Вы можете назначить только один обработчик для `Office.initialize`, который будет вызываться только один раз инфраструктурой Office, но вы можете вызывать `Office.onReady()` в разных местах вашего кода и использовать разные обратные вызовы.</span><span class="sxs-lookup"><span data-stu-id="5ae92-159">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="5ae92-160">Например, ваш код может вызвать `Office.onReady()` сразу после загрузки настраиваемого скрипта с обратным вызовом, запускающим логику инициализации. В коде также может применяться кнопка в области задач, чей скрипт вызывает `Office.onReady()` с другим обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5ae92-160">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="5ae92-161">В этом случае второй обратный вызов запускается при нажатии кнопки.</span><span class="sxs-lookup"><span data-stu-id="5ae92-161">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="5ae92-162">Событие `Office.initialize` запускается в конце выполнения внутренних процессов, когда Office.js инициализирует собственное выполнение.</span><span class="sxs-lookup"><span data-stu-id="5ae92-162">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="5ae92-163">И оно срабатывает *сразу же* после окончания внутренних процессов.</span><span class="sxs-lookup"><span data-stu-id="5ae92-163">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="5ae92-164">Если код, в котором вы назначаете обработчик события, выполняется слишком долго после запуска события, тогда обработчик не запускается.</span><span class="sxs-lookup"><span data-stu-id="5ae92-164">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="5ae92-165">Например если вы используете диспетчер задач WebPack, он может настроить домашнюю страницу надстройки для загрузки файлов полизаполнения сразу после загрузки Office.js, но перед загрузкой вашего настраиваемого скрипта JavaScript.</span><span class="sxs-lookup"><span data-stu-id="5ae92-165">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="5ae92-166">К тому моменту, когда ваш скрипт загружается и назначает обработчика, инициализации события уже выполнена.</span><span class="sxs-lookup"><span data-stu-id="5ae92-166">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="5ae92-167">Но никогда не «поздно» выполнить вызов `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="5ae92-167">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="5ae92-168">Если инициализация события уже произошла, обратный вызов выполняется немедленно.</span><span class="sxs-lookup"><span data-stu-id="5ae92-168">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="5ae92-169">Даже если отсутствует логика запуска, следует вызвать `Office.onReady()` или назначить пустую функцию для `Office.initialize`, когда ваша надстройка загружает JavaScript.</span><span class="sxs-lookup"><span data-stu-id="5ae92-169">Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads.</span></span> <span data-ttu-id="5ae92-170">Некоторые ведущие приложения Office и сочетания платформ не загружают область задач, пока не произойдет одно из этих событий.</span><span class="sxs-lookup"><span data-stu-id="5ae92-170">Some Office host and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="5ae92-171">Эти два способа показаны в приведенных ниже примерах.</span><span class="sxs-lookup"><span data-stu-id="5ae92-171">The following examples show these two approaches.</span></span>
>
>```js  
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="5ae92-172">Объектная модель API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="5ae92-172">Office JavaScript API object model</span></span>

<span data-ttu-id="5ae92-173">После инициализации надстройка может взаимодействовать с хостом (например, Excel, Outlook).</span><span class="sxs-lookup"><span data-stu-id="5ae92-173">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook).</span></span> <span data-ttu-id="5ae92-174">Страница [объектной модели Office JavaScript API](office-javascript-api-object-model.md) содержит дополнительную информацию об определенных способах использования.</span><span class="sxs-lookup"><span data-stu-id="5ae92-174">The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns.</span></span> <span data-ttu-id="5ae92-175">Также существует подробная справочная документация для [общих API](/office/dev/add-ins/reference/javascript-api-for-office) и API для определенных ведущих приложений.</span><span class="sxs-lookup"><span data-stu-id="5ae92-175">There is also detailed reference documentation for both [Common APIs](/office/dev/add-ins/reference/javascript-api-for-office) and host-specific APIs.</span></span>
