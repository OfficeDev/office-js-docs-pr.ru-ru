---
title: Инициализация надстройки Office
description: Узнайте, как инициализировать надстройку Office.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 5dc9d0143ac9eaab18625e280891bd601fa9f899
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293326"
---
# <a name="initialize-your-office-add-in"></a><span data-ttu-id="fccc5-103">Инициализация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="fccc5-103">Initialize your Office Add-in</span></span>

<span data-ttu-id="fccc5-104">Надстройки Office часто поддерживают логику запуска для выполнения следующих действий:</span><span class="sxs-lookup"><span data-stu-id="fccc5-104">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="fccc5-105">Убедитесь, что версия Office пользователя поддерживает все API Office, которые вызывает ваш код.</span><span class="sxs-lookup"><span data-stu-id="fccc5-105">Check that the user's version of Office supports all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="fccc5-106">Обеспечьте наличие определенных артефактов, например листа с определенным именем.</span><span class="sxs-lookup"><span data-stu-id="fccc5-106">Ensure the existence of certain artifacts, such as a worksheet with a specific name.</span></span>

- <span data-ttu-id="fccc5-107">Предлагать пользователю выбрать некоторые ячейки в Excel, а затем вставить диаграмму, инициализированную с выбранными значениями.</span><span class="sxs-lookup"><span data-stu-id="fccc5-107">Prompt the user to select some cells in Excel, and then insert a chart initialized with those selected values.</span></span>

- <span data-ttu-id="fccc5-108">Установите привязки.</span><span class="sxs-lookup"><span data-stu-id="fccc5-108">Establish bindings.</span></span>

- <span data-ttu-id="fccc5-109">Используйте API диалоговых окон Office, чтобы запросить у пользователя значения параметров надстройки по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="fccc5-109">Use the Office Dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="fccc5-110">Однако Надстройка Office не может вызывать все API JavaScript для Office, пока библиотека не будет загружена.</span><span class="sxs-lookup"><span data-stu-id="fccc5-110">However, an Office Add-in cannot successfully call any Office JavaScript APIs until the library has been loaded.</span></span> <span data-ttu-id="fccc5-111">В этой статье описываются два способа, позволяющие коду убедиться в том, что библиотека загружена:</span><span class="sxs-lookup"><span data-stu-id="fccc5-111">This article describes the two ways your code can ensure that the library has been loaded:</span></span>

- <span data-ttu-id="fccc5-112">Инициализация с помощью `Office.onReady()` .</span><span class="sxs-lookup"><span data-stu-id="fccc5-112">Initialize with `Office.onReady()`.</span></span>
- <span data-ttu-id="fccc5-113">Инициализация с помощью `Office.initialize` .</span><span class="sxs-lookup"><span data-stu-id="fccc5-113">Initialize with `Office.initialize`.</span></span>

> [!TIP]
> <span data-ttu-id="fccc5-114">Рекомендуется использовать `Office.onReady()` вместо `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="fccc5-114">We recommend that you use `Office.onReady()` instead of `Office.initialize`.</span></span> <span data-ttu-id="fccc5-115">Хотя `Office.initialize` все еще поддерживаются, `Office.onReady()` обеспечивается большая гибкость.</span><span class="sxs-lookup"><span data-stu-id="fccc5-115">Although `Office.initialize` is still supported, `Office.onReady()` provides more flexibility.</span></span> <span data-ttu-id="fccc5-116">В инфраструктуре Office можно назначить только один обработчик, `Office.initialize` и он будет вызываться только один раз.</span><span class="sxs-lookup"><span data-stu-id="fccc5-116">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure.</span></span> <span data-ttu-id="fccc5-117">Вы можете звонить `Office.onReady()` в различных местах кода и использовать разные обратные вызовы.</span><span class="sxs-lookup"><span data-stu-id="fccc5-117">You can call `Office.onReady()` in different places in your code and use different callbacks.</span></span>
> 
> <span data-ttu-id="fccc5-118">Сведения о различиях описанных ниже приемов см. в статье [Основные различия между Office.initialize и Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span><span class="sxs-lookup"><span data-stu-id="fccc5-118">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span>

<span data-ttu-id="fccc5-119">Дополнительные сведения о последовательности событий при инициализации надстройки см. в статье [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).</span><span class="sxs-lookup"><span data-stu-id="fccc5-119">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

## <a name="initialize-with-officeonready"></a><span data-ttu-id="fccc5-120">Инициализация с использованием Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="fccc5-120">Initialize with Office.onReady()</span></span>

<span data-ttu-id="fccc5-121">`Office.onReady()` — Это асинхронный метод, который возвращает объект [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) при проверке загрузки библиотеки Office.js.</span><span class="sxs-lookup"><span data-stu-id="fccc5-121">`Office.onReady()` is an asynchronous method that returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) object while it checks to see if the Office.js library is loaded.</span></span> <span data-ttu-id="fccc5-122">Когда библиотека будет загружена, она разрешается в качестве объекта, который указывает клиентское приложение Office со `Office.HostType` значением перечисления (,, `Excel` и `Word` т. д.) и платформой со `Office.PlatformType` значением перечисления ( `PC` , `Mac` , `OfficeOnline` , и т. д.).</span><span class="sxs-lookup"><span data-stu-id="fccc5-122">When the library is loaded, it resolves the Promise as an object that specifies the Office client application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="fccc5-123">Объект Promise сопоставляется незамедлительно, если библиотека уже загружена, когда вызывается `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="fccc5-123">The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.</span></span>

<span data-ttu-id="fccc5-124">Один из способов вызова `Office.onReady()` состоит в передаче ему метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fccc5-124">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="fccc5-125">Пример:</span><span class="sxs-lookup"><span data-stu-id="fccc5-125">Here's an example:</span></span>

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

<span data-ttu-id="fccc5-126">Кроме того, вы можете привязать метод `then()` к вызову `Office.onReady()`, вместо того чтобы использовать обратный вызов.</span><span class="sxs-lookup"><span data-stu-id="fccc5-126">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="fccc5-127">Например приведенный ниже код проверяет, поддерживает ли версия Excel пользователя использование API, которые может вызывать надстройка.</span><span class="sxs-lookup"><span data-stu-id="fccc5-127">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="fccc5-128">Вот аналогичный же пример, использующий ключевые слова `async` и `await` в TypeScript:</span><span class="sxs-lookup"><span data-stu-id="fccc5-128">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="fccc5-129">При использовании дополнительных платформ JavaScript, включающих собственный обработчик событий инициализации или тесты, они, *как правило*, должны размещаться внутри ответа для `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="fccc5-129">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="fccc5-130">Например, ссылка на [JQuery](https://jquery.com) функция `$(document).ready()` должна выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="fccc5-130">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="fccc5-131">Однако существуют исключения для таких случаев.</span><span class="sxs-lookup"><span data-stu-id="fccc5-131">However, there are exceptions to this practice.</span></span> <span data-ttu-id="fccc5-132">Например, предположим, что вы хотите открыть надстройку в браузере (вместо Загрузка неопубликованных ее в приложении Office), чтобы выполнить отладку пользовательского интерфейса с помощью средств браузера.</span><span class="sxs-lookup"><span data-stu-id="fccc5-132">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office application) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="fccc5-133">Так как Office.js не загружается в браузер, `onReady` не будет работать, а `$(document).ready` не будет работать при вызове внутри Office `onReady`.</span><span class="sxs-lookup"><span data-stu-id="fccc5-133">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> 

<span data-ttu-id="fccc5-134">Еще одно исключение, если вы хотите, чтобы индикатор хода выполнения отображался в области задач при загрузке надстройки.</span><span class="sxs-lookup"><span data-stu-id="fccc5-134">Another exception would be if you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="fccc5-135">В этом сценарии код должен вызывать jQuery `ready` и использовать обратный вызов для отображения индикатора хода выполнения.</span><span class="sxs-lookup"><span data-stu-id="fccc5-135">In this scenario, your code should call the jQuery `ready` and use its callback to render the progress indicator.</span></span> <span data-ttu-id="fccc5-136">Затем обратный вызов `onReady` Office может заменять индикатор выполнения на окончательный пользовательский интерфейс </span><span class="sxs-lookup"><span data-stu-id="fccc5-136">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

## <a name="initialize-with-officeinitialize"></a><span data-ttu-id="fccc5-137">Инициализация с использованием Office.initialize</span><span class="sxs-lookup"><span data-stu-id="fccc5-137">Initialize with Office.initialize</span></span>

<span data-ttu-id="fccc5-138">Событие инициализации запускается, когда библиотека Office.js будет загружена и готова к взаимодействию с пользователем.</span><span class="sxs-lookup"><span data-stu-id="fccc5-138">An initialize event fires when the Office.js library is loaded and ready for user interaction.</span></span> <span data-ttu-id="fccc5-139">Вы можете назначить обработчик `Office.initialize` для реализации вашей логики инициализации.</span><span class="sxs-lookup"><span data-stu-id="fccc5-139">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="fccc5-140">Например, приведенный ниже код проверяет, поддерживает ли версия Excel пользователя использование API, которые может вызывать надстройка.</span><span class="sxs-lookup"><span data-stu-id="fccc5-140">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="fccc5-141">При использовании дополнительных платформ JavaScript, включающих собственные обработчики инициализации или тесты, они *обычно* должны размещаться внутри `Office.initialize` события (исключения, описанные в разделе **initialize with Office. onreading ()** , ранее применимы в этом случае).</span><span class="sxs-lookup"><span data-stu-id="fccc5-141">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event (the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also).</span></span> <span data-ttu-id="fccc5-142">Например, ссылка на [JQuery](https://jquery.com) функция `$(document).ready()` должна выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="fccc5-142">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="fccc5-143">Для надстроек области задач и контентных надстроек `Office.initialize` предоставляет дополнительный параметр _reason_.</span><span class="sxs-lookup"><span data-stu-id="fccc5-143">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="fccc5-144">Этот параметр определяет порядок добавления надстройки в текущий документ.</span><span class="sxs-lookup"><span data-stu-id="fccc5-144">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="fccc5-145">Это поможет обеспечить разную логику в тех случаях, когда надстройка вставляется впервые или когда она уже существует в документе.</span><span class="sxs-lookup"><span data-stu-id="fccc5-145">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="fccc5-146">Дополнительные сведения см. в статьях [Событие Office.initialize](/javascript/api/office) и [Перечисление InitializationReason](/javascript/api/office/office.initializationreason).</span><span class="sxs-lookup"><span data-stu-id="fccc5-146">For more information, see [Office.initialize Event](/javascript/api/office) and [InitializationReason Enumeration](/javascript/api/office/office.initializationreason).</span></span>

## <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="fccc5-147">Основные различия между Office.initialize и Office.onReady</span><span class="sxs-lookup"><span data-stu-id="fccc5-147">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="fccc5-148">Вы можете назначить только один обработчик для `Office.initialize`, который будет вызываться только один раз инфраструктурой Office, но вы можете вызывать `Office.onReady()` в разных местах вашего кода и использовать разные обратные вызовы.</span><span class="sxs-lookup"><span data-stu-id="fccc5-148">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="fccc5-149">Например, ваш код может вызвать `Office.onReady()` сразу после загрузки настраиваемого скрипта с обратным вызовом, запускающим логику инициализации. В коде также может применяться кнопка в области задач, чей скрипт вызывает `Office.onReady()` с другим обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="fccc5-149">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="fccc5-150">В этом случае второй обратный вызов запускается при нажатии кнопки.</span><span class="sxs-lookup"><span data-stu-id="fccc5-150">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="fccc5-151">Событие `Office.initialize` запускается в конце выполнения внутренних процессов, когда Office.js инициализирует собственное выполнение.</span><span class="sxs-lookup"><span data-stu-id="fccc5-151">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="fccc5-152">И оно срабатывает *сразу же* после окончания внутренних процессов.</span><span class="sxs-lookup"><span data-stu-id="fccc5-152">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="fccc5-153">Если код, в котором вы назначаете обработчик события, выполняется слишком долго после запуска события, тогда обработчик не запускается.</span><span class="sxs-lookup"><span data-stu-id="fccc5-153">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="fccc5-154">Например если вы используете диспетчер задач WebPack, он может настроить домашнюю страницу надстройки для загрузки файлов полизаполнения сразу после загрузки Office.js, но перед загрузкой вашего настраиваемого скрипта JavaScript.</span><span class="sxs-lookup"><span data-stu-id="fccc5-154">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="fccc5-155">К тому моменту, когда ваш скрипт загружается и назначает обработчика, инициализации события уже выполнена.</span><span class="sxs-lookup"><span data-stu-id="fccc5-155">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="fccc5-156">Но никогда не «поздно» выполнить вызов `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="fccc5-156">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="fccc5-157">Если инициализация события уже произошла, обратный вызов выполняется немедленно.</span><span class="sxs-lookup"><span data-stu-id="fccc5-157">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="fccc5-158">Даже если отсутствует логика запуска, следует вызвать `Office.onReady()` или назначить пустую функцию для `Office.initialize`, когда ваша надстройка загружает JavaScript.</span><span class="sxs-lookup"><span data-stu-id="fccc5-158">Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads.</span></span> <span data-ttu-id="fccc5-159">Некоторые сочетания приложений и приложений Office не загружают область задач до тех пор, пока не будет выполняться одно из этих действий.</span><span class="sxs-lookup"><span data-stu-id="fccc5-159">Some Office application and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="fccc5-160">Эти два способа показаны в приведенных ниже примерах.</span><span class="sxs-lookup"><span data-stu-id="fccc5-160">The following examples show these two approaches.</span></span>
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a><span data-ttu-id="fccc5-161">См. также</span><span class="sxs-lookup"><span data-stu-id="fccc5-161">See also</span></span>

- [<span data-ttu-id="fccc5-162">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="fccc5-162">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="fccc5-163">Загрузка модели DOM и среды выполнения</span><span class="sxs-lookup"><span data-stu-id="fccc5-163">Loading the DOM and runtime environment</span></span>](loading-the-dom-and-runtime-environment.md)