---
title: Элемент FunctionFile в файле манифеста
description: Указывает исходный файл кода для операций, которые надстройка предоставляет с помощью команд надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: f31a1bc7a561305a89f5388102a4985aaa31fe37
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348302"
---
# <a name="functionfile-element"></a><span data-ttu-id="7c4c9-103">Элемент FunctionFile</span><span class="sxs-lookup"><span data-stu-id="7c4c9-103">FunctionFile element</span></span>

<span data-ttu-id="7c4c9-104">Указывает исходный код файла для операций, которые надстройка предоставляет одним из следующих способов.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-104">Specifies the source code file for operations that an add-in exposes in one of the following ways.</span></span>

* <span data-ttu-id="7c4c9-105">Команды надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-105">Add-in commands that execute a JavaScript function instead of displaying UI.</span></span>
* <span data-ttu-id="7c4c9-106">Клавиши, которые выполняют функцию JavaScript.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-106">Keyboard shortcuts that execute a JavaScript function.</span></span>

<span data-ttu-id="7c4c9-107">Элемент `FunctionFile` является детским элементом [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="7c4c9-107">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="7c4c9-108">Атрибут элемента может быть не более 32 символов и задарен значению атрибута элемента в элементе, который содержит `resid` `FunctionFile` `id` `Url` `Resources` URL-адрес HTML-файла, [](control.md)который содержит или загружает все функции JavaScript, используемые кнопками командной команды без пользовательского интерфейса, как это определено элементом Control.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-108">The `resid` attribute of the `FunctionFile` element can be no more than 32 characters and is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="7c4c9-109">Ниже приводится пример `FunctionFile` элемента.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-109">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="7c4c9-110">JavaScript в HTML-файле, указанный элементом, должен вызывать и определять именные функции, которые принимают `FunctionFile` `Office.initialize` один параметр: `event` .</span><span class="sxs-lookup"><span data-stu-id="7c4c9-110">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="7c4c9-111">Функции должны использовать API `item.notificationMessages`, чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-111">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="7c4c9-112">Он также должен вызывать метод `event.completed` после выполнения.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-112">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="7c4c9-113">Имя функций используется в элементе для кнопок без `FunctionName` пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-113">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="7c4c9-114">Ниже приводится пример HTML-файла, определяющий `trackMessage` функцию.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-114">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

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

<span data-ttu-id="7c4c9-115">В следующем коде показано, как реализовать используемую функцию `FunctionName` .</span><span class="sxs-lookup"><span data-stu-id="7c4c9-115">The following code shows how to implement the function used by `FunctionName`.</span></span>

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
> <span data-ttu-id="7c4c9-116">Вызов `event.completed` сигналов об успешном обращении с событием.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-116">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="7c4c9-117">Если функция вызывается несколько раз, например при выборе одной команды надстройки несколько раз, все события автоматически помещаются в очередь.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-117">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="7c4c9-118">Первое событие запускается автоматически, тогда как остальные ожидают в очереди.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-118">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="7c4c9-119">При вызове функции выполняется следующий вызов в очереди на `event.completed` эту функцию.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-119">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="7c4c9-120">Необходимо `event.completed` вызвать; в противном случае функция не будет работать.</span><span class="sxs-lookup"><span data-stu-id="7c4c9-120">You must call `event.completed`; otherwise your function will not run.</span></span>
