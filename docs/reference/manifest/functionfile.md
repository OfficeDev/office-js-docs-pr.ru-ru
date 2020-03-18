---
title: Элемент FunctionFile в файле манифеста
description: Указывает файл исходного кода для операций, предоставляемых надстройкой через команды надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 376ea82f48360d502ea9be05dc5d6b02f9294add
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718197"
---
# <a name="functionfile-element"></a><span data-ttu-id="e968b-103">Элемент FunctionFile</span><span class="sxs-lookup"><span data-stu-id="e968b-103">FunctionFile element</span></span>

<span data-ttu-id="e968b-104">Указывает файл исходного кода для операций, предоставляемых надстройкой через команды надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="e968b-104">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI.</span></span> <span data-ttu-id="e968b-105">`FunctionFile` Элемент является дочерним для элемента [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e968b-105">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="e968b-106">Для `resid` `FunctionFile` атрибута элемента задается значение `id` атрибута `Url` элемента в `Resources` элементе, который содержит URL-адрес HTML-файла, который содержит или загружает все функции JavaScript, которые используются кнопками надстройки без пользовательского интерфейса, как определено [элементом Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="e968b-106">The `resid` attribute of the `FunctionFile` element is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="e968b-107">Ниже приведен пример `FunctionFile` элемента.</span><span class="sxs-lookup"><span data-stu-id="e968b-107">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="e968b-108">JavaScript в HTML-файле, указанном `FunctionFile` элементом, должен вызывать `Office.initialize` и определять именованные функции, которые принимают один параметр `event`:.</span><span class="sxs-lookup"><span data-stu-id="e968b-108">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="e968b-109">Функции должны использовать API `item.notificationMessages`, чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке.</span><span class="sxs-lookup"><span data-stu-id="e968b-109">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="e968b-110">Он также должен вызывать метод `event.completed` после выполнения.</span><span class="sxs-lookup"><span data-stu-id="e968b-110">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="e968b-111">Имя функции используется в `FunctionName` элементе для кнопок без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="e968b-111">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="e968b-112">Ниже приведен пример HTML-файла, определяющего `trackMessage` функцию.</span><span class="sxs-lookup"><span data-stu-id="e968b-112">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

<span data-ttu-id="e968b-113">В приведенном ниже коде показано, как реализовать функцию, `FunctionName`используемую в.</span><span class="sxs-lookup"><span data-stu-id="e968b-113">The following code shows how to implement the function used by `FunctionName`.</span></span>

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> <span data-ttu-id="e968b-114">Вызов, `event.completed` сигнализирующий о том, что событие успешно обработано.</span><span class="sxs-lookup"><span data-stu-id="e968b-114">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="e968b-115">Если функция вызывается несколько раз, например при выборе одной команды надстройки несколько раз, все события автоматически помещаются в очередь.</span><span class="sxs-lookup"><span data-stu-id="e968b-115">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="e968b-116">Первое событие запускается автоматически, тогда как остальные ожидают в очереди.</span><span class="sxs-lookup"><span data-stu-id="e968b-116">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="e968b-117">При вызове `event.completed`функции выполняется следующий вызов этой функции в очереди.</span><span class="sxs-lookup"><span data-stu-id="e968b-117">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="e968b-118">Необходимо вызвать `event.completed`; в противном случае функция не будет выполняться.</span><span class="sxs-lookup"><span data-stu-id="e968b-118">You must call `event.completed`; otherwise your function will not run.</span></span>
