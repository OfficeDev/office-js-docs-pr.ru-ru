---
title: Общие сведения об API JavaScript для Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e9d9efdda5e237ab076d22d50b1f7ded5e075845
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505953"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="3cd24-102">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="3cd24-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="3cd24-p101">В этой статье можно узнать об API JavaScript для Office и о том, как его использовать. Справочные сведения см. в статье [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) . О том, как обновить файлы проекта Visual Studio до последней версии API JavaScript для Office, см. в статье [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md) .</span><span class="sxs-lookup"><span data-stu-id="3cd24-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="3cd24-p102">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и о ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="3cd24-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="3cd24-108">Ссылки на библиотеку API JavaScript для Office в вашей надстройке</span><span class="sxs-lookup"><span data-stu-id="3cd24-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="3cd24-p103">Библиотека [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например, Excel-15.js и Outlook-15.js. Простейший способ сослаться на API — использовать нашу сеть CDN, добавив следующий код `<script>` в тег страницы `<head>`:</span><span class="sxs-lookup"><span data-stu-id="3cd24-p103">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="3cd24-111">Это приведет к скачиванию и кэшированию файлов API JavaScript для Office при первой загрузке надстройки, чтобы убедиться, что она использует самую актуальную реализацию Office.js и сопутствующих файлов для указанной версии.</span><span class="sxs-lookup"><span data-stu-id="3cd24-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="3cd24-112">Подробные сведения о CDN-версии файла Office.js, включая способы управления версиями и обратной совместимостью, приведены в разделе [Указание ссылок на библиотеку API JavaScript для Office из сети доставки содержимого (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="3cd24-112">For more details around the Office.js CDN, including how versioning and backward compatability is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="3cd24-113">Инициализация надстройки</span><span class="sxs-lookup"><span data-stu-id="3cd24-113">Initializing your add-in</span></span>

<span data-ttu-id="3cd24-114">**Область применения:** все типы надстроек</span><span class="sxs-lookup"><span data-stu-id="3cd24-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="3cd24-115">Надстройки Office часто имеют логику, выполняющую при запуске таких действий, как:</span><span class="sxs-lookup"><span data-stu-id="3cd24-115">Office add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="3cd24-116">Проверка того, будет ли поддерживать пользовательская версия Office все функции API для Office, вызываемые вашим кодом.</span><span class="sxs-lookup"><span data-stu-id="3cd24-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="3cd24-117">Проверка наличия некоторых артефактов, таких как лист с конкретным именем.</span><span class="sxs-lookup"><span data-stu-id="3cd24-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="3cd24-118">Пользователю предлагается выбрать несколько ячеек в Excel, а затем вставить диаграмму, созданную с использованием этих выбранных значений.</span><span class="sxs-lookup"><span data-stu-id="3cd24-118">You can use the initialize event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="3cd24-119">Установление привязок.</span><span class="sxs-lookup"><span data-stu-id="3cd24-119">Establish bindings.</span></span>

- <span data-ttu-id="3cd24-120">Используйте API диалога для Office, предлагающий пользователю установить для параметров надстройки значения по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="3cd24-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="3cd24-p104">Однако запуск кода не должен вызывать API-интерфейсы  Office.js, пока библиотека не загрузится полностью. Существует два способа обеспечить загрузку библиотеки кодом. Они описаны в следующих разделах. Мы рекомендуем использовать более новую и более гибкую технику, вызвав `Office.onReady()`. Предыдущая техника, которая назначает обработчик для `Office.initialize`, по-прежнему поддерживается. См. также статью [Основные различия между Office.initialize и Office.onReady()](#major-differences-between-office-initialize-and-office-onready).</span><span class="sxs-lookup"><span data-stu-id="3cd24-p104">But your start-up code must not call any Office.js APIs until the library is fully loaded. There are two ways that your code can ensure that the library is loaded. They are described in the sections below. We recommend that you use the newer, more flexible, technique, calling `Office.onReady()`. The older technique, assigning a handler to `Office.initialize`, is still supported. See also [Major differences between Office.initialize and Office.onReady()](#major-differences-between-office-initialize-and-office-onready).</span></span>

<span data-ttu-id="3cd24-127">Дополнительные сведения о последовательности событий при инициализации надстройки приведены в разделе [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).</span><span class="sxs-lookup"><span data-stu-id="3cd24-127">For more detail about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="3cd24-128">Инициализация с помощью Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="3cd24-128">Initialize with Office.onReady()</span></span>

<span data-ttu-id="3cd24-p105">`Office.onReady()` представляет собой асинхронный метод, который возвращает объект Promise, проверяя при этом, полностью ли загрузилась библиотека Office.js. Когда  библиотека загрузилась (и только тогда), метод разрешает Promise как объект, указывающий ведущее приложение Office со значением перечисления `Office.HostType` (`Excel`, `Word`и т.д.) и платформу со значением перечисления `Office.PlatformType` (`PC`, `Mac`, `OfficeOnline`, и т.д.). Если библиотека уже загружена при вызове `Office.onReady()`, объект Promise разрешается немедленно.</span><span class="sxs-lookup"><span data-stu-id="3cd24-p105">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is fully loaded. When, and only when, the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.). If the library is already loaded when `Office.onReady()` is called, then the Promise resolves immediately.</span></span>

<span data-ttu-id="3cd24-p106">Один из способов вызвать `Office.onReady()` — это передать его метод обратного вызова. Ниже приведен пример:</span><span class="sxs-lookup"><span data-stu-id="3cd24-p106">One way to call `Office.onReady()` is to pass it a callback method. Here's an example:</span></span>

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

<span data-ttu-id="3cd24-p107">Кроме того, можно объединять метод `then()` для вызова `Office.onReady()` вместо передачи обратного вызова. Например, следующий код проверяет, поддерживает ли версия Excel пользователя все API-интерфейсы, которые может вызвать надстройка.</span><span class="sxs-lookup"><span data-stu-id="3cd24-p107">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback. For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="3cd24-136">Ниже приведен тот же пример с использованием ключевых слов `async` и `await` в TypeScript:</span><span class="sxs-lookup"><span data-stu-id="3cd24-136">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="3cd24-p108">При использовании дополнительных платформ JavaScript, включающих собственный обработчик событий инициализации или тесты, *как правило*, их следует размещать внутри ответа на `Office.onReady()`. Например, ссылка на функцию [JQuery](https://jquery.com) `$(document).ready()` будет выполнена следующим образом:</span><span class="sxs-lookup"><span data-stu-id="3cd24-p108">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be *usually* be placed within the response to `Office.onReady()`. For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="3cd24-p109">Тем не менее, существуют исключения для этого метода. Предположим, например, что вы хотите открыть надстройку в браузере (а не загрузить ее неопубликованной в ведущем приложении Office) для отладки вашего пользовательского интерфейса с помощью инструментов веб-обозревателя. Поскольку Office.js не будет загружаться в веб-обозревателе, `onReady` и `$(document).ready` не будут выполняться при вызове внутри Office `onReady`. Еще одно исключение: вы хотите, чтобы индикатор хода выполнения отображался на панели задач в процессе загрузки надстройки. В этом сценарии код должен вызывать jQuery `ready` и использовать его обратный вызов, чтобы отобразить индикатор выполнения. Затем обратный вызов Office `onReady` сможет заменить индикатор хода выполнения окончательным вариантом пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="3cd24-p109">However, there are exceptions to this practice. For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools. Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`. Another exception: you want a progress indicator to appear in the task pane while the add-in is loading. In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator. Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="3cd24-145">Инициализация с использованием функции Office.initialize</span><span class="sxs-lookup"><span data-stu-id="3cd24-145">Initialize with Office.initialize</span></span>

<span data-ttu-id="3cd24-p110">Событие инициализации вызывается, когда библиотека Office.js полностью загружена и готова к взаимодействию с пользователем. Можно назначить обработчик для `Office.initialize` , который реализует логику инициализации. Ниже приведен пример, в котором проверяется, что пользователь версии Excel поддерживает все интерфейсы API, которые может вызвать надстройка.</span><span class="sxs-lookup"><span data-stu-id="3cd24-p110">An initialize event fires when the Office.js library is fully loaded and ready for user interaction. You can assign a handler to `Office.initialize` that implements your initialization logic. The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="3cd24-p111">При использовании дополнительных платформ JavaScript, у которых есть собственный обработчик инициализации или тесты, они должны, *как правило*, размещаться в событии `Office.initialize`. (Однако исключения, описанные ранее в разделе **Инициализация с Office.onReady()**, также применяются в этом случае.) Например, ссылка на функцию [JQuery](https://jquery.com) `$(document).ready()` будет выполнена следующим образом:</span><span class="sxs-lookup"><span data-stu-id="3cd24-p111">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event. (But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="3cd24-p112">Для надстроек области задач и надстроек содержимого `Office.initialize` обеспечивает дополнительный параметр _reason_. Этот параметр указывает, как надстройка была добавлена в текущий документ. Это поможет обеспечить разную логику в тех случаях, когда надстройка вставляется впервые, или когда она уже существует в документе.</span><span class="sxs-lookup"><span data-stu-id="3cd24-p112">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter. This parameter specifies how an add-in was added to the current document. You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="3cd24-154">Дополнительные сведения см. в разделах [Событие Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) и [Перечисление InitializationReason](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="3cd24-154">For more information, see [Office.initialize Event](https://docs.microsoft.com/javascript/api/office?view=office-js) and [InitializationReason Enumeration](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js).</span></span>

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="3cd24-155">Основные различия между Office.initialize и Office.onReady</span><span class="sxs-lookup"><span data-stu-id="3cd24-155">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="3cd24-p113">Можно назначить только один обработчик для `Office.initialize`, и он вызывается только один раз в инфраструктуре Office. Однако можно вызвать `Office.onReady()` в различных местах вашего кода и использовать различные обратные вызовы. Например, код может вызывать `Office.onReady()` сразу же после того, как ваш пользовательский сценарий загрузится с помощью обратного вызова, на котором выполняется логика инициализации. У вашего кода также может быть кнопка на области задач, сценарий которой вызывает `Office.onReady()` другим обратным вызовом. В этом случае обратный вызов выполняется при нажатии кнопки.</span><span class="sxs-lookup"><span data-stu-id="3cd24-p113">You can assign only one handler to `Office.initialize` and it is called, only once, by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks. For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback. If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="3cd24-p114">Событие  `Office.initialize` запускается в конце внутреннего процесса, в котором инициализируется Office.js. Оно запускается *сразу же* после завершения внутреннего процесса. Если код, в котором вы присвоили обработчика событию, выполняется слишком долго после запуска события, тогда ваш обработчик не запускается. Например, при использовании диспетчера задач WebPack он может настроить домашнюю страницу надстройки для загрузки файлов polyfill после загрузки файла Office.js, но перед загрузкой настраиваемого JavaScript. К моменту загрузки вашего сценария и назначения им обработчика событие инициализации уже произойдет. Но никогда не «слишком поздно» вызвать `Office.onReady()`. Если событие инициализации уже произошло, обратный вызов выполнится немедленно.</span><span class="sxs-lookup"><span data-stu-id="3cd24-p114">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself. And it fires *immediately* after the internal process ends. If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run. For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript. By the time your script loads and assigns the handler, the initialize event has already happened. But it is never "too late" to call `Office.onReady()`. If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="3cd24-p115">Даже если у вас нет логики запуска, рекомендуется либо вызвать `Office.onReady()`, либо назначить пустую функцию для `Office.initialize` при загрузке вашей надстройки JavaScript, поскольку некоторые комбинации ведущего приложения и платформы Office не будут загружать область задач до тех пор, пока не произойдет одно из этих событий. Следующие строки показывают два пути выполнения этой процедуры:</span><span class="sxs-lookup"><span data-stu-id="3cd24-p115">Even if you have no start-up logic, it is a good practice to either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads, because some Office host and platform combinations won't load the task pane until one of these happens. The following lines show the two ways this can be done:</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="3cd24-168">Объектная модель API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="3cd24-168">Office JavaScript API object model</span></span>

<span data-ttu-id="3cd24-p116">После инициализации надстройки могут взаимодействовать с узлом (например, Excel, Outlook). Страница [Office JavaScript API object model](office-javascript-api-object-model.md) содержит более подробные сведения по  определенному использованию шаблонов. Также имеются подробные справочные материалы по [общим API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) и конкретным узлам.</span><span class="sxs-lookup"><span data-stu-id="3cd24-p116">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook). The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns. There is also detailed reference documentation for both [shared APIs](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) and specific hosts.</span></span>

## <a name="api-support-matrix"></a><span data-ttu-id="3cd24-172">Матрица поддержки API</span><span class="sxs-lookup"><span data-stu-id="3cd24-172">API support matrix</span></span>

<span data-ttu-id="3cd24-173">В этой таблице представлены API и функции, поддерживаемые всеми типами надстроек (надстройками содержимого, области задач и Outlook), а также приложения Office, в которых они могут работать, когда вы указываете ведущие приложения Office, поддерживаемые вашей надстройкой, с помощью [схемы манифестов надстроек версии 1.1 и функций, поддерживаемых API JavaScript для Office версии 1.1](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="3cd24-173">This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||<span data-ttu-id="3cd24-174">**Имя узла**</span><span class="sxs-lookup"><span data-stu-id="3cd24-174">**Host name**</span></span>|<span data-ttu-id="3cd24-175">База данных</span><span class="sxs-lookup"><span data-stu-id="3cd24-175">Database</span></span>|<span data-ttu-id="3cd24-176">Книга</span><span class="sxs-lookup"><span data-stu-id="3cd24-176">Workbook</span></span>|<span data-ttu-id="3cd24-177">Почтовый ящик</span><span class="sxs-lookup"><span data-stu-id="3cd24-177">Mailbox</span></span>|<span data-ttu-id="3cd24-178">Презентация</span><span class="sxs-lookup"><span data-stu-id="3cd24-178">Presentation</span></span>|<span data-ttu-id="3cd24-179">Документ</span><span class="sxs-lookup"><span data-stu-id="3cd24-179">Document</span></span>|<span data-ttu-id="3cd24-180">Project</span><span class="sxs-lookup"><span data-stu-id="3cd24-180">Project</span></span>|
||<span data-ttu-id="3cd24-181">**Поддерживаемые** **ведущие приложения**</span><span class="sxs-lookup"><span data-stu-id="3cd24-181">**Supported** **Host applications**</span></span>|<span data-ttu-id="3cd24-182">Веб-приложения Access</span><span class="sxs-lookup"><span data-stu-id="3cd24-182">Access web apps</span></span>|<span data-ttu-id="3cd24-183">Excel,</span><span class="sxs-lookup"><span data-stu-id="3cd24-183">Excel,</span></span><br/><span data-ttu-id="3cd24-184">Excel Online</span><span class="sxs-lookup"><span data-stu-id="3cd24-184">Excel Online</span></span>|<span data-ttu-id="3cd24-185">Outlook,</span><span class="sxs-lookup"><span data-stu-id="3cd24-185">Outlook,</span></span><br/><span data-ttu-id="3cd24-186">веб-приложение Outlook,</span><span class="sxs-lookup"><span data-stu-id="3cd24-186">Outlook Web App,</span></span><br/><span data-ttu-id="3cd24-187">OWA (веб-приложения Outlook) для устройств</span><span class="sxs-lookup"><span data-stu-id="3cd24-187">OWA for Devices</span></span>|<span data-ttu-id="3cd24-188">PowerPoint,</span><span class="sxs-lookup"><span data-stu-id="3cd24-188"> (PowerPoint)</span></span><br/><span data-ttu-id="3cd24-189">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="3cd24-189">PowerPoint Online</span></span>|<span data-ttu-id="3cd24-190">Word</span><span class="sxs-lookup"><span data-stu-id="3cd24-190">Word</span></span>|<span data-ttu-id="3cd24-191">Project</span><span class="sxs-lookup"><span data-stu-id="3cd24-191">Project</span></span>|
|<span data-ttu-id="3cd24-192">**Поддерживаемые типы надстроек**</span><span class="sxs-lookup"><span data-stu-id="3cd24-192">**Supported add-in types**</span></span>|<span data-ttu-id="3cd24-193">Содержимое</span><span class="sxs-lookup"><span data-stu-id="3cd24-193">Content</span></span>|<span data-ttu-id="3cd24-194">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-194">Y</span></span>|<span data-ttu-id="3cd24-195">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-195">Y</span></span>||<span data-ttu-id="3cd24-196">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-196">Y</span></span>|||
||<span data-ttu-id="3cd24-197">Область задач</span><span class="sxs-lookup"><span data-stu-id="3cd24-197">Task pane</span></span>||<span data-ttu-id="3cd24-198">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-198">Y</span></span>||<span data-ttu-id="3cd24-199">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-199">Y</span></span>|<span data-ttu-id="3cd24-200">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-200">Y</span></span>|<span data-ttu-id="3cd24-201">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-201">Y</span></span>|
||<span data-ttu-id="3cd24-202">Outlook</span><span class="sxs-lookup"><span data-stu-id="3cd24-202">Outlook</span></span>|||<span data-ttu-id="3cd24-203">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-203">Y</span></span>||||
|<span data-ttu-id="3cd24-204">**Поддерживаемые функции API**</span><span class="sxs-lookup"><span data-stu-id="3cd24-204">**Supported API features**</span></span>|<span data-ttu-id="3cd24-205">Чтение/запись текста</span><span class="sxs-lookup"><span data-stu-id="3cd24-205">Read/Write Text</span></span>||<span data-ttu-id="3cd24-206">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-206">Y</span></span>||<span data-ttu-id="3cd24-207">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-207">Y</span></span>|<span data-ttu-id="3cd24-208">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-208">Y</span></span>|<span data-ttu-id="3cd24-209">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-209">Y</span></span><br/><span data-ttu-id="3cd24-210">(только для чтения)</span><span class="sxs-lookup"><span data-stu-id="3cd24-210">(Read only)</span></span>|
||<span data-ttu-id="3cd24-211">Чтение/запись матрицы</span><span class="sxs-lookup"><span data-stu-id="3cd24-211">Read/Write Matrix</span></span>||<span data-ttu-id="3cd24-212">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-212">Y</span></span>|||<span data-ttu-id="3cd24-213">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-213">Y</span></span>||
||<span data-ttu-id="3cd24-214">Чтение/запись таблицы</span><span class="sxs-lookup"><span data-stu-id="3cd24-214">Read/Write Table</span></span>||<span data-ttu-id="3cd24-215">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-215">Y</span></span>|||<span data-ttu-id="3cd24-216">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-216">Y</span></span>||
||<span data-ttu-id="3cd24-217">Чтение/запись HTML</span><span class="sxs-lookup"><span data-stu-id="3cd24-217">Read/Write HTML</span></span>|||||<span data-ttu-id="3cd24-218">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-218">Y</span></span>||
||<span data-ttu-id="3cd24-219">Чтение/запись</span><span class="sxs-lookup"><span data-stu-id="3cd24-219">Read/Write</span></span><br/><span data-ttu-id="3cd24-220">Office Open XML</span><span class="sxs-lookup"><span data-stu-id="3cd24-220">Office Open XML</span></span>|||||<span data-ttu-id="3cd24-221">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-221">Y</span></span>||
||<span data-ttu-id="3cd24-222">Чтение свойств task, resource, view и field</span><span class="sxs-lookup"><span data-stu-id="3cd24-222">Read task, resource, view, and field properties</span></span>||||||<span data-ttu-id="3cd24-223">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-223">Y</span></span>|
||<span data-ttu-id="3cd24-224">События изменения выделения</span><span class="sxs-lookup"><span data-stu-id="3cd24-224">Selection changed events</span></span>||<span data-ttu-id="3cd24-225">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-225">Y</span></span>|||<span data-ttu-id="3cd24-226">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-226">Y</span></span>||
||<span data-ttu-id="3cd24-227">Загрузка всего документа</span><span class="sxs-lookup"><span data-stu-id="3cd24-227">Get whole document</span></span>||||<span data-ttu-id="3cd24-228">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-228">Y</span></span>|<span data-ttu-id="3cd24-229">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-229">Y</span></span>||
||<span data-ttu-id="3cd24-230">Привязки и их события</span><span class="sxs-lookup"><span data-stu-id="3cd24-230">Bindings and binding events</span></span>|<span data-ttu-id="3cd24-231">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-231">Y</span></span><br/><span data-ttu-id="3cd24-232">(только полные и частичные привязки таблиц)</span><span class="sxs-lookup"><span data-stu-id="3cd24-232">(Only full and partial table bindings)</span></span>|<span data-ttu-id="3cd24-233">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-233">Y</span></span>|||<span data-ttu-id="3cd24-234">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-234">Y</span></span>||
||<span data-ttu-id="3cd24-235">Чтение/запись настраиваемых XML-частей</span><span class="sxs-lookup"><span data-stu-id="3cd24-235">Read/Write Custom XML Parts</span></span>|||||<span data-ttu-id="3cd24-236">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-236">Y</span></span>||
||<span data-ttu-id="3cd24-237">Сохранение данных состояния надстройки (параметры)</span><span class="sxs-lookup"><span data-stu-id="3cd24-237">Persist add-in state data (settings)</span></span>|<span data-ttu-id="3cd24-238">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-238">Y</span></span><br/><span data-ttu-id="3cd24-239">(на ведущую надстройку)</span><span class="sxs-lookup"><span data-stu-id="3cd24-239">(Per host add-in)</span></span>|<span data-ttu-id="3cd24-240">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-240">Y</span></span><br/><span data-ttu-id="3cd24-241">(на документ)</span><span class="sxs-lookup"><span data-stu-id="3cd24-241">(Per document)</span></span>|<span data-ttu-id="3cd24-242">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-242">Y</span></span><br/><span data-ttu-id="3cd24-243">(на почтовый ящик)</span><span class="sxs-lookup"><span data-stu-id="3cd24-243">(Per mailbox)</span></span>|<span data-ttu-id="3cd24-244">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-244">Y</span></span><br/><span data-ttu-id="3cd24-245">(на документ)</span><span class="sxs-lookup"><span data-stu-id="3cd24-245">(Per document)</span></span>|<span data-ttu-id="3cd24-246">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-246">Y</span></span><br/><span data-ttu-id="3cd24-247">(на документ)</span><span class="sxs-lookup"><span data-stu-id="3cd24-247">(Per document)</span></span>||
||<span data-ttu-id="3cd24-248">События изменения параметров</span><span class="sxs-lookup"><span data-stu-id="3cd24-248">Settings changed events</span></span>|<span data-ttu-id="3cd24-249">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-249">Y</span></span>|<span data-ttu-id="3cd24-250">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-250">Y</span></span>||<span data-ttu-id="3cd24-251">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-251">Y</span></span>|<span data-ttu-id="3cd24-252">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-252">Y</span></span>||
||<span data-ttu-id="3cd24-253">Получение активного режима просмотра</span><span class="sxs-lookup"><span data-stu-id="3cd24-253">Get active view mode</span></span><br/><span data-ttu-id="3cd24-254">и просмотр измененных событий</span><span class="sxs-lookup"><span data-stu-id="3cd24-254">and view changed events</span></span>||||<span data-ttu-id="3cd24-255">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-255">Y</span></span>|||
||<span data-ttu-id="3cd24-256">Переход к расположениям</span><span class="sxs-lookup"><span data-stu-id="3cd24-256">Navigate to locations</span></span><br/><span data-ttu-id="3cd24-257">в документе</span><span class="sxs-lookup"><span data-stu-id="3cd24-257">in the document</span></span>||<span data-ttu-id="3cd24-258">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-258">Y</span></span>||<span data-ttu-id="3cd24-259">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-259">Y</span></span>|<span data-ttu-id="3cd24-260">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-260">Y</span></span>||
||<span data-ttu-id="3cd24-261">Активация в зависимости от контекста</span><span class="sxs-lookup"><span data-stu-id="3cd24-261">Activate contextually</span></span><br/><span data-ttu-id="3cd24-262">с помощью правил и регулярных выражений</span><span class="sxs-lookup"><span data-stu-id="3cd24-262">using rules and RegEx</span></span>|||<span data-ttu-id="3cd24-263">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-263">Y</span></span>||||
||<span data-ttu-id="3cd24-264">Чтение свойств элемента</span><span class="sxs-lookup"><span data-stu-id="3cd24-264">Read Item properties</span></span>|||<span data-ttu-id="3cd24-265">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-265">Y</span></span>||||
||<span data-ttu-id="3cd24-266">Чтение профиля пользователя</span><span class="sxs-lookup"><span data-stu-id="3cd24-266">Read User profile</span></span>|||<span data-ttu-id="3cd24-267">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-267">Y</span></span>||||
||<span data-ttu-id="3cd24-268">Получение вложений</span><span class="sxs-lookup"><span data-stu-id="3cd24-268">Get attachments</span></span>|||<span data-ttu-id="3cd24-269">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-269">Y</span></span>||||
||<span data-ttu-id="3cd24-270">Получение маркера удостоверения пользователя</span><span class="sxs-lookup"><span data-stu-id="3cd24-270">Get User identity token</span></span>|||<span data-ttu-id="3cd24-271">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-271">Y</span></span>||||
||<span data-ttu-id="3cd24-272">Вызов веб-служб Exchange</span><span class="sxs-lookup"><span data-stu-id="3cd24-272">Call Exchange Web Services</span></span>|||<span data-ttu-id="3cd24-273">Да</span><span class="sxs-lookup"><span data-stu-id="3cd24-273">Y</span></span>||||
