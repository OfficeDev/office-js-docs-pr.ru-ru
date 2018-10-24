---
title: Общие сведения об API JavaScript для Office
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 58829c623c06225bcc7d15925fb02a082df039c6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640094"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="393c4-102">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="393c4-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="393c4-p101">В этой статье можно узнать об API JavaScript для Office и о том, как его использовать. Справочные сведения см. в статье [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) . О том, как обновить файлы проекта Visual Studio до последней версии API JavaScript для Office, см. в статье [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md) .</span><span class="sxs-lookup"><span data-stu-id="393c4-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="393c4-p102">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и о ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="393c4-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="393c4-108">Ссылки на библиотеку API JavaScript для Office в вашей надстройке</span><span class="sxs-lookup"><span data-stu-id="393c4-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="393c4-p103">Библиотека [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например, Excel-15.js и Outlook-15.js. Простейший способ сослаться на API — использовать нашу сеть CDN, добавив следующий код `<script>` в тег страницы `<head>`:</span><span class="sxs-lookup"><span data-stu-id="393c4-p103">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="393c4-111">Это приведет к скачиванию и кэшированию файлов API JavaScript для Office при первой загрузке надстройки, чтобы убедиться, что она использует самую актуальную реализацию Office.js и сопутствующих файлов для указанной версии.</span><span class="sxs-lookup"><span data-stu-id="393c4-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="393c4-112">Подробные сведения о CDN-версии файла Office.js, включая способы управления версиями и обратной совместимостью, приведены в разделе [Указание ссылок на библиотеку API JavaScript для Office из сети доставки содержимого (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="393c4-112">For more details around the Office.js CDN, including how versioning and backward compatability is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="393c4-113">Инициализация надстройки</span><span class="sxs-lookup"><span data-stu-id="393c4-113">Initializing your add-in</span></span>

<span data-ttu-id="393c4-114">**Область применения:** все типы надстроек</span><span class="sxs-lookup"><span data-stu-id="393c4-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="393c4-115">Надстройки Office часто имеют логику, выполняющую при запуске таких действий, как:</span><span class="sxs-lookup"><span data-stu-id="393c4-115">Office add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="393c4-116">Проверка того, будет ли поддерживать пользовательская версия Office все функции API для Office, вызываемые вашим кодом.</span><span class="sxs-lookup"><span data-stu-id="393c4-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="393c4-117">Проверка наличия некоторых артефактов, таких как лист с конкретным именем.</span><span class="sxs-lookup"><span data-stu-id="393c4-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="393c4-118">Пользователю предлагается выбрать несколько ячеек в Excel, а затем вставить диаграмму, созданную с использованием этих выбранных значений.</span><span class="sxs-lookup"><span data-stu-id="393c4-118">You can use the initialize event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="393c4-119">Установление привязок.</span><span class="sxs-lookup"><span data-stu-id="393c4-119">Establish bindings.</span></span>

- <span data-ttu-id="393c4-120">Используйте API диалога для Office, предлагающий пользователю установить для параметров надстройки значения по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="393c4-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="393c4-p104">Но ваш стартовый код не должен вызывать API-интерфейсы Office.js до тех пор, пока библиотека не будет полностью загружена. Имеется два способа для проверки вашим кодом загрузки библиотеки. Они описаны в следующих разделах:</span><span class="sxs-lookup"><span data-stu-id="393c4-p104">But your start-up code must not call any Office.js APIs until the library is fully loaded. There are two ways that your code can ensure that the library is loaded. They are described in the sections below. We recommend that you use the newer, more flexible, technique, calling . The older technique, assigning a handler to , is still supported. See also Major differences between Office.initialize and Office.onReady().</span></span> 

- [<span data-ttu-id="393c4-124">Инициализация с помощью Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="393c4-124">Initialize with Office.onReady()</span></span>](#initialize-with-officeonready)
- [<span data-ttu-id="393c4-125">Инициализация с использованием функции Office.initialize</span><span class="sxs-lookup"><span data-stu-id="393c4-125">Initialize with Office.initialize</span></span>](#initialize-with-officeinitialize)

<span data-ttu-id="393c4-126">Для получения сведений о различиях в этих методах см. [Основные различия между Office.initialize и Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span><span class="sxs-lookup"><span data-stu-id="393c4-126">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span> <span data-ttu-id="393c4-127">Дополнительные сведения о последовательности событий при инициализации надстройки приведены в разделе [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).</span><span class="sxs-lookup"><span data-stu-id="393c4-127">For more detail about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="393c4-128">Инициализация с помощью Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="393c4-128">Initialize with Office.onReady()</span></span>

<span data-ttu-id="393c4-p106">`Office.onReady()` представляет собой асинхронный метод, который возвращает объект Promise, проверяя при этом, полностью ли загрузилась библиотека Office.js. Когда  библиотека загрузилась (и только тогда), метод разрешает Promise как объект, указывающий ведущее приложение Office со значением перечисления `Office.HostType` (`Excel`, `Word`и т.д.) и платформу со значением перечисления `Office.PlatformType` (`PC`, `Mac`, `OfficeOnline`, и т.д.). Если библиотека уже загружена при вызове `Office.onReady()`, объект Promise разрешается немедленно.</span><span class="sxs-lookup"><span data-stu-id="393c4-p106">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is fully loaded. When, and only when, the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.). If the library is already loaded when `Office.onReady()` is called, then the Promise resolves immediately.</span></span>

<span data-ttu-id="393c4-p107">Один из способов вызвать `Office.onReady()` — это передать его метод обратного вызова. Ниже приведен пример:</span><span class="sxs-lookup"><span data-stu-id="393c4-p107">One way to call `Office.onReady()` is to pass it a callback method. Here's an example:</span></span>

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

<span data-ttu-id="393c4-p108">Кроме того, можно объединять метод `then()` для вызова `Office.onReady()` вместо передачи обратного вызова. Например, следующий код проверяет, поддерживает ли версия Excel пользователя все API-интерфейсы, которые может вызвать надстройка.</span><span class="sxs-lookup"><span data-stu-id="393c4-p108">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback. For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="393c4-136">Ниже приведен тот же пример с использованием ключевых слов `async` и `await` в TypeScript:</span><span class="sxs-lookup"><span data-stu-id="393c4-136">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="393c4-p109">При использовании дополнительных платформ JavaScript, включающих собственный обработчик событий инициализации или тесты, *как правило*, их следует размещать внутри ответа на `Office.onReady()`. Например, ссылка на функцию [JQuery](https://jquery.com) `$(document).ready()` будет выполнена следующим образом:</span><span class="sxs-lookup"><span data-stu-id="393c4-p109">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be *usually* be placed within the response to `Office.onReady()`. For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="393c4-p110">Тем не менее, существуют исключения для этого метода. Предположим, например, что вы хотите открыть надстройку в браузере (а не загрузить ее неопубликованной в ведущем приложении Office) для отладки вашего пользовательского интерфейса с помощью инструментов веб-обозревателя. Поскольку Office.js не будет загружаться в веб-обозревателе, `onReady` и `$(document).ready` не будут выполняться при вызове внутри Office `onReady`. Еще одно исключение: вы хотите, чтобы индикатор хода выполнения отображался на панели задач в процессе загрузки надстройки. В этом сценарии код должен вызывать jQuery `ready` и использовать его обратный вызов, чтобы отобразить индикатор выполнения. Затем обратный вызов Office `onReady` сможет заменить индикатор хода выполнения окончательным вариантом пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="393c4-p110">However, there are exceptions to this practice. For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools. Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`. Another exception: you want a progress indicator to appear in the task pane while the add-in is loading. In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator. Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="393c4-145">Инициализация с использованием функции Office.initialize</span><span class="sxs-lookup"><span data-stu-id="393c4-145">Initialize with Office.initialize</span></span>

<span data-ttu-id="393c4-p111">Событие инициализации вызывается, когда библиотека Office.js полностью загружена и готова к взаимодействию с пользователем. Можно назначить обработчик для `Office.initialize` , который реализует логику инициализации. Ниже приведен пример, в котором проверяется, что пользователь версии Excel поддерживает все интерфейсы API, которые может вызвать надстройка.</span><span class="sxs-lookup"><span data-stu-id="393c4-p111">An initialize event fires when the Office.js library is fully loaded and ready for user interaction. You can assign a handler to `Office.initialize` that implements your initialization logic. The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="393c4-p112">При использовании дополнительных платформ JavaScript, у которых есть собственный обработчик инициализации или тесты, они должны, *как правило*, размещаться в событии `Office.initialize`. (Однако исключения, описанные ранее в разделе **Инициализация с Office.onReady()**, также применяются в этом случае.) Например, ссылка на функцию [JQuery](https://jquery.com) `$(document).ready()` будет выполнена следующим образом:</span><span class="sxs-lookup"><span data-stu-id="393c4-p112">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event. (But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="393c4-p113">Для надстроек области задач и надстроек содержимого `Office.initialize` обеспечивает дополнительный параметр _reason_. Этот параметр указывает, как надстройка была добавлена в текущий документ. Это поможет обеспечить разную логику в тех случаях, когда надстройка вставляется впервые, или когда она уже существует в документе.</span><span class="sxs-lookup"><span data-stu-id="393c4-p113">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter. This parameter specifies how an add-in was added to the current document. You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="393c4-154">Дополнительные сведения см. в статьях [Событие Office.initialize Event](https://docs.microsoft.com/javascript/api/office?view=office-js) и [Перечисление InitializationReason](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="393c4-154">For more information, see [Office.initialize Event](https://docs.microsoft.com/javascript/api/office?view=office-js) and [InitializationReason Enumeration](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js).</span></span>

> [!NOTE]
> <span data-ttu-id="393c4-155">В настоящее время, необходимо установить `Office.Initialize`, независимо от того, вызывается ли еще и `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="393c4-155">Currently, you must set `Office.Initialize`, regardless of whether `Office.onReady()` is also called.</span></span> <span data-ttu-id="393c4-156">Если вы не используете `Office.Initialize`, вы можете задать в нем пустую функцию, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="393c4-156">If you have no use for `Office.Initialize`, you can set it to an empty function as shown in the following example.</span></span>
> 
>```js
>Office.initialize = function () {};
>```

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="393c4-157">Основные различия между Office.initialize и Office.onReady</span><span class="sxs-lookup"><span data-stu-id="393c4-157">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="393c4-p115">Можно назначить только один обработчик для `Office.initialize`, и он вызывается только один раз в инфраструктуре Office. Однако можно вызвать `Office.onReady()` в различных местах вашего кода и использовать различные обратные вызовы. Например, код может вызывать `Office.onReady()` сразу же после того, как ваш пользовательский сценарий загрузится с помощью обратного вызова, на котором выполняется логика инициализации. У вашего кода также может быть кнопка на области задач, сценарий которой вызывает `Office.onReady()` другим обратным вызовом. В этом случае обратный вызов выполняется при нажатии кнопки.</span><span class="sxs-lookup"><span data-stu-id="393c4-p115">You can assign only one handler to `Office.initialize` and it is called, only once, by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks. For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback. If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="393c4-p116">Событие  `Office.initialize` запускается в конце внутреннего процесса, в котором инициализируется Office.js. Оно запускается *сразу же* после завершения внутреннего процесса. Если код, в котором вы присвоили обработчика событию, выполняется слишком долго после запуска события, тогда ваш обработчик не запускается. Например, при использовании диспетчера задач WebPack он может настроить домашнюю страницу надстройки для загрузки файлов polyfill после загрузки файла Office.js, но перед загрузкой настраиваемого JavaScript. К моменту загрузки вашего сценария и назначения им обработчика событие инициализации уже произойдет. Но никогда не «слишком поздно» вызвать `Office.onReady()`. Если событие инициализации уже произошло, обратный вызов выполнится немедленно.</span><span class="sxs-lookup"><span data-stu-id="393c4-p116">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself. And it fires *immediately* after the internal process ends. If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run. For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript. By the time your script loads and assigns the handler, the initialize event has already happened. But it is never "too late" to call `Office.onReady()`. If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="393c4-168">Даже если у вас нет логики запуска, вы должны назначить пустую функцию `Office.initialize` при загрузке надстройки JavaScript, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="393c4-168">Even if you have no start-up logic, you should assign an empty function to `Office.initialize` when your add-in JavaScript loads, as shown in the following example.</span></span> <span data-ttu-id="393c4-169">Некоторые комбинации ведущего приложения и платформы Office не будут загружать панель задач до тех пор, пока не произойдет событие инициализации и не будет запущена указанная функция обработчика событий.</span><span class="sxs-lookup"><span data-stu-id="393c4-169">Some Office host and platform combinations won't load the task pane until the initialize event fires and the specified event handler function runs.</span></span>
> 
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="393c4-170">Объектная модель API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="393c4-170">Office JavaScript API object model</span></span>

<span data-ttu-id="393c4-p118">После инициализации надстройки могут взаимодействовать с узлом (например, Excel, Outlook). Страница [Office JavaScript API object model](office-javascript-api-object-model.md) содержит более подробные сведения по  определенному использованию шаблонов. Также имеются подробные справочные материалы по [общим API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) и конкретным узлам.</span><span class="sxs-lookup"><span data-stu-id="393c4-p118">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook). The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns. There is also detailed reference documentation for both [shared APIs](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) and specific hosts.</span></span>

## <a name="api-support-matrix"></a><span data-ttu-id="393c4-174">Матрица поддержки API</span><span class="sxs-lookup"><span data-stu-id="393c4-174">API support matrix</span></span>

<span data-ttu-id="393c4-175">В этой таблице представлены API и функции, поддерживаемые всеми типами надстроек (надстройками содержимого, области задач и Outlook), а также приложения Office, в которых они могут работать, когда вы указываете ведущие приложения Office, поддерживаемые вашей надстройкой, с помощью [схемы манифестов надстроек версии 1.1 и функций, поддерживаемых API JavaScript для Office версии 1.1](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="393c4-175">This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||<span data-ttu-id="393c4-176">**Имя узла**</span><span class="sxs-lookup"><span data-stu-id="393c4-176">**Host name**</span></span>|<span data-ttu-id="393c4-177">База данных</span><span class="sxs-lookup"><span data-stu-id="393c4-177">Database</span></span>|<span data-ttu-id="393c4-178">Книга</span><span class="sxs-lookup"><span data-stu-id="393c4-178">Workbook</span></span>|<span data-ttu-id="393c4-179">Почтовый ящик</span><span class="sxs-lookup"><span data-stu-id="393c4-179">Mailbox</span></span>|<span data-ttu-id="393c4-180">Презентация</span><span class="sxs-lookup"><span data-stu-id="393c4-180">Presentation</span></span>|<span data-ttu-id="393c4-181">Документ</span><span class="sxs-lookup"><span data-stu-id="393c4-181">Document</span></span>|<span data-ttu-id="393c4-182">Проект</span><span class="sxs-lookup"><span data-stu-id="393c4-182">Project</span></span>|
||<span data-ttu-id="393c4-183">**Поддерживаемые** **ведущие приложения**</span><span class="sxs-lookup"><span data-stu-id="393c4-183">**Supported** **Host applications**</span></span>|<span data-ttu-id="393c4-184">Веб-приложения Access</span><span class="sxs-lookup"><span data-stu-id="393c4-184">Access web apps</span></span>|<span data-ttu-id="393c4-185">Excel,</span><span class="sxs-lookup"><span data-stu-id="393c4-185">Excel,</span></span><br/><span data-ttu-id="393c4-186">Excel Online</span><span class="sxs-lookup"><span data-stu-id="393c4-186">Excel Online</span></span>|<span data-ttu-id="393c4-187">Outlook,</span><span class="sxs-lookup"><span data-stu-id="393c4-187">Outlook,</span></span><br/><span data-ttu-id="393c4-188">веб-приложение Outlook,</span><span class="sxs-lookup"><span data-stu-id="393c4-188">Outlook Web App,</span></span><br/><span data-ttu-id="393c4-189">OWA (веб-приложения Outlook) для устройств</span><span class="sxs-lookup"><span data-stu-id="393c4-189">OWA for Devices</span></span>|<span data-ttu-id="393c4-190">PowerPoint,</span><span class="sxs-lookup"><span data-stu-id="393c4-190">PowerPoint</span></span><br/><span data-ttu-id="393c4-191">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="393c4-191">PowerPoint Online</span></span>|<span data-ttu-id="393c4-192">Word</span><span class="sxs-lookup"><span data-stu-id="393c4-192">Word</span></span>|<span data-ttu-id="393c4-193">Project</span><span class="sxs-lookup"><span data-stu-id="393c4-193">Project</span></span>|
|<span data-ttu-id="393c4-194">**Поддерживаемые типы надстроек**</span><span class="sxs-lookup"><span data-stu-id="393c4-194">**Supported add-in types**</span></span>|<span data-ttu-id="393c4-195">Содержимое</span><span class="sxs-lookup"><span data-stu-id="393c4-195">Content</span></span>|<span data-ttu-id="393c4-196">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-196">Y</span></span>|<span data-ttu-id="393c4-197">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-197">Y</span></span>||<span data-ttu-id="393c4-198">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-198">Y</span></span>|||
||<span data-ttu-id="393c4-199">Область задач</span><span class="sxs-lookup"><span data-stu-id="393c4-199">Task pane</span></span>||<span data-ttu-id="393c4-200">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-200">Y</span></span>||<span data-ttu-id="393c4-201">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-201">Y</span></span>|<span data-ttu-id="393c4-202">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-202">Y</span></span>|<span data-ttu-id="393c4-203">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-203">Y</span></span>|
||<span data-ttu-id="393c4-204">Outlook</span><span class="sxs-lookup"><span data-stu-id="393c4-204">Outlook</span></span>|||<span data-ttu-id="393c4-205">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-205">Y</span></span>||||
|<span data-ttu-id="393c4-206">**Поддерживаемые функции API**</span><span class="sxs-lookup"><span data-stu-id="393c4-206">**Supported API features**</span></span>|<span data-ttu-id="393c4-207">Чтение/запись текста</span><span class="sxs-lookup"><span data-stu-id="393c4-207">Read/Write Text</span></span>||<span data-ttu-id="393c4-208">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-208">Y</span></span>||<span data-ttu-id="393c4-209">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-209">Y</span></span>|<span data-ttu-id="393c4-210">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-210">Y</span></span>|<span data-ttu-id="393c4-211">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-211">Y</span></span><br/><span data-ttu-id="393c4-212">(только для чтения)</span><span class="sxs-lookup"><span data-stu-id="393c4-212">(Read only)</span></span>|
||<span data-ttu-id="393c4-213">Чтение/запись матрицы</span><span class="sxs-lookup"><span data-stu-id="393c4-213">Read/Write Matrix</span></span>||<span data-ttu-id="393c4-214">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-214">Y</span></span>|||<span data-ttu-id="393c4-215">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-215">Y</span></span>||
||<span data-ttu-id="393c4-216">Чтение/запись таблицы</span><span class="sxs-lookup"><span data-stu-id="393c4-216">Read/Write Table</span></span>||<span data-ttu-id="393c4-217">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-217">Y</span></span>|||<span data-ttu-id="393c4-218">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-218">Y</span></span>||
||<span data-ttu-id="393c4-219">Чтение/запись HTML</span><span class="sxs-lookup"><span data-stu-id="393c4-219">Read/Write HTML</span></span>|||||<span data-ttu-id="393c4-220">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-220">Y</span></span>||
||<span data-ttu-id="393c4-221">Чтение/запись</span><span class="sxs-lookup"><span data-stu-id="393c4-221">Read/Write</span></span><br/><span data-ttu-id="393c4-222">Office Open XML</span><span class="sxs-lookup"><span data-stu-id="393c4-222">Office Open XML</span></span>|||||<span data-ttu-id="393c4-223">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-223">Y</span></span>||
||<span data-ttu-id="393c4-224">Чтение свойств task, resource, view и field</span><span class="sxs-lookup"><span data-stu-id="393c4-224">Read task, resource, view, and field properties</span></span>||||||<span data-ttu-id="393c4-225">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-225">Y</span></span>|
||<span data-ttu-id="393c4-226">События изменения выделения</span><span class="sxs-lookup"><span data-stu-id="393c4-226">Selection changed events</span></span>||<span data-ttu-id="393c4-227">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-227">Y</span></span>|||<span data-ttu-id="393c4-228">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-228">Y</span></span>||
||<span data-ttu-id="393c4-229">Загрузка всего документа</span><span class="sxs-lookup"><span data-stu-id="393c4-229">Get whole document</span></span>||||<span data-ttu-id="393c4-230">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-230">Y</span></span>|<span data-ttu-id="393c4-231">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-231">Y</span></span>||
||<span data-ttu-id="393c4-232">Привязки и их события</span><span class="sxs-lookup"><span data-stu-id="393c4-232">Bindings and binding events</span></span>|<span data-ttu-id="393c4-233">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-233">Y</span></span><br/><span data-ttu-id="393c4-234">(только полные и частичные привязки таблиц)</span><span class="sxs-lookup"><span data-stu-id="393c4-234">(Only full and partial table bindings)</span></span>|<span data-ttu-id="393c4-235">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-235">Y</span></span>|||<span data-ttu-id="393c4-236">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-236">Y</span></span>||
||<span data-ttu-id="393c4-237">Чтение/запись настраиваемых XML-частей</span><span class="sxs-lookup"><span data-stu-id="393c4-237">Read/Write Custom XML Parts</span></span>|||||<span data-ttu-id="393c4-238">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-238">Y</span></span>||
||<span data-ttu-id="393c4-239">Сохранение данных состояния надстройки (параметры)</span><span class="sxs-lookup"><span data-stu-id="393c4-239">Persist add-in state data (settings)</span></span>|<span data-ttu-id="393c4-240">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-240">Y</span></span><br/><span data-ttu-id="393c4-241">(на ведущую надстройку)</span><span class="sxs-lookup"><span data-stu-id="393c4-241">(Per host add-in)</span></span>|<span data-ttu-id="393c4-242">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-242">Y</span></span><br/><span data-ttu-id="393c4-243">(на документ)</span><span class="sxs-lookup"><span data-stu-id="393c4-243">(Per document)</span></span>|<span data-ttu-id="393c4-244">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-244">Y</span></span><br/><span data-ttu-id="393c4-245">(на почтовый ящик)</span><span class="sxs-lookup"><span data-stu-id="393c4-245">(Per mailbox)</span></span>|<span data-ttu-id="393c4-246">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-246">Y</span></span><br/><span data-ttu-id="393c4-247">(на документ)</span><span class="sxs-lookup"><span data-stu-id="393c4-247">(Per document)</span></span>|<span data-ttu-id="393c4-248">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-248">Y</span></span><br/><span data-ttu-id="393c4-249">(на документ)</span><span class="sxs-lookup"><span data-stu-id="393c4-249">(Per document)</span></span>||
||<span data-ttu-id="393c4-250">События изменения параметров</span><span class="sxs-lookup"><span data-stu-id="393c4-250">Settings changed events</span></span>|<span data-ttu-id="393c4-251">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-251">Y</span></span>|<span data-ttu-id="393c4-252">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-252">Y</span></span>||<span data-ttu-id="393c4-253">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-253">Y</span></span>|<span data-ttu-id="393c4-254">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-254">Y</span></span>||
||<span data-ttu-id="393c4-255">Получение активного режима просмотра</span><span class="sxs-lookup"><span data-stu-id="393c4-255">Get active view mode</span></span><br/><span data-ttu-id="393c4-256">и просмотр измененных событий</span><span class="sxs-lookup"><span data-stu-id="393c4-256">and view changed events</span></span>||||<span data-ttu-id="393c4-257">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-257">Y</span></span>|||
||<span data-ttu-id="393c4-258">Переход к расположениям</span><span class="sxs-lookup"><span data-stu-id="393c4-258">Navigate to locations</span></span><br/><span data-ttu-id="393c4-259">в документе</span><span class="sxs-lookup"><span data-stu-id="393c4-259">in the document</span></span>||<span data-ttu-id="393c4-260">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-260">Y</span></span>||<span data-ttu-id="393c4-261">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-261">Y</span></span>|<span data-ttu-id="393c4-262">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-262">Y</span></span>||
||<span data-ttu-id="393c4-263">Активация в зависимости от контекста</span><span class="sxs-lookup"><span data-stu-id="393c4-263">Activate contextually</span></span><br/><span data-ttu-id="393c4-264">с помощью правил и регулярных выражений</span><span class="sxs-lookup"><span data-stu-id="393c4-264">using rules and RegEx</span></span>|||<span data-ttu-id="393c4-265">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-265">Y</span></span>||||
||<span data-ttu-id="393c4-266">Чтение свойств элемента</span><span class="sxs-lookup"><span data-stu-id="393c4-266">Read Item properties</span></span>|||<span data-ttu-id="393c4-267">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-267">Y</span></span>||||
||<span data-ttu-id="393c4-268">Чтение профиля пользователя</span><span class="sxs-lookup"><span data-stu-id="393c4-268">Read User profile</span></span>|||<span data-ttu-id="393c4-269">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-269">Y</span></span>||||
||<span data-ttu-id="393c4-270">Получение вложений</span><span class="sxs-lookup"><span data-stu-id="393c4-270">Get attachments</span></span>|||<span data-ttu-id="393c4-271">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-271">Y</span></span>||||
||<span data-ttu-id="393c4-272">Получение маркера удостоверения пользователя</span><span class="sxs-lookup"><span data-stu-id="393c4-272">Get User identity token</span></span>|||<span data-ttu-id="393c4-273">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-273">Y</span></span>||||
||<span data-ttu-id="393c4-274">Вызов веб-служб Exchange</span><span class="sxs-lookup"><span data-stu-id="393c4-274">Call Exchange Web Services</span></span>|||<span data-ttu-id="393c4-275">Да</span><span class="sxs-lookup"><span data-stu-id="393c4-275">Y</span></span>||||
