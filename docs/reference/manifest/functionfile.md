---
title: Элемент FunctionFile в файле манифеста
description: Указывает файл исходного кода для операций, предоставляемых надстройкой через команды надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 4c47c3e4b824f2b93aaea17cef88e01f748d6f95
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996447"
---
# <a name="functionfile-element"></a><span data-ttu-id="c7f79-103">Элемент FunctionFile</span><span class="sxs-lookup"><span data-stu-id="c7f79-103">FunctionFile element</span></span>

<span data-ttu-id="c7f79-104">Указывает файл исходного кода для операций, предоставляемых надстройкой, одним из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="c7f79-104">Specifies the source code file for operations that an add-in exposes in one of the following ways:</span></span>

* <span data-ttu-id="c7f79-105">Команды надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="c7f79-105">Add-in commands that execute a JavaScript function instead of displaying UI.</span></span>
* <span data-ttu-id="c7f79-106">Сочетания клавиш, которые выполняют функцию JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c7f79-106">Keyboard shortcuts that execute a JavaScript function.</span></span>

<span data-ttu-id="c7f79-107">`FunctionFile`Элемент является дочерним для элемента [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="c7f79-107">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="c7f79-108">`resid` `FunctionFile` Для атрибута элемента задается значение `id` атрибута `Url` элемента в `Resources` элементе, который содержит URL-адрес HTML-файла, который содержит или загружает все функции JavaScript, которые используются КНОПКАМИ надстройки без пользовательского интерфейса, как определено [элементом Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="c7f79-108">The `resid` attribute of the `FunctionFile` element is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="c7f79-109">Ниже приведен пример `FunctionFile` элемента.</span><span class="sxs-lookup"><span data-stu-id="c7f79-109">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="c7f79-110">JavaScript в HTML-файле, указанном `FunctionFile` элементом, должен вызывать `Office.initialize` и определять именованные функции, которые принимают один параметр: `event` .</span><span class="sxs-lookup"><span data-stu-id="c7f79-110">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="c7f79-111">Функции должны использовать API `item.notificationMessages`, чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке.</span><span class="sxs-lookup"><span data-stu-id="c7f79-111">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="c7f79-112">Он также должен вызывать метод `event.completed` после выполнения.</span><span class="sxs-lookup"><span data-stu-id="c7f79-112">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="c7f79-113">Имя функции используется в `FunctionName` элементе для кнопок без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="c7f79-113">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="c7f79-114">Ниже приведен пример HTML-файла, определяющего `trackMessage` функцию.</span><span class="sxs-lookup"><span data-stu-id="c7f79-114">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

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

<span data-ttu-id="c7f79-115">В приведенном ниже коде показано, как реализовать функцию, используемую в `FunctionName` .</span><span class="sxs-lookup"><span data-stu-id="c7f79-115">The following code shows how to implement the function used by `FunctionName`.</span></span>

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
> <span data-ttu-id="c7f79-116">Вызов `event.completed` , сигнализирующий о том, что событие успешно обработано.</span><span class="sxs-lookup"><span data-stu-id="c7f79-116">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="c7f79-117">Если функция вызывается несколько раз, например при выборе одной команды надстройки несколько раз, все события автоматически помещаются в очередь.</span><span class="sxs-lookup"><span data-stu-id="c7f79-117">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="c7f79-118">Первое событие запускается автоматически, тогда как остальные ожидают в очереди.</span><span class="sxs-lookup"><span data-stu-id="c7f79-118">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="c7f79-119">При вызове функции `event.completed` выполняется следующий вызов этой функции в очереди.</span><span class="sxs-lookup"><span data-stu-id="c7f79-119">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="c7f79-120">Необходимо позвонить `event.completed` ; в противном случае функция не будет запускаться.</span><span class="sxs-lookup"><span data-stu-id="c7f79-120">You must call `event.completed`; otherwise your function will not run.</span></span>
