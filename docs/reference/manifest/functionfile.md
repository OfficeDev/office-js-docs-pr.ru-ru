---
title: Элемент FunctionFile в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5f87d10428b58adfb89f1119ba5741599079afba
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450585"
---
# <a name="functionfile-element"></a><span data-ttu-id="8747c-102">Элемент FunctionFile</span><span class="sxs-lookup"><span data-stu-id="8747c-102">FunctionFile element</span></span>

<span data-ttu-id="8747c-p101">Указывает файл с исходным кодом для операций, доступных через те команды надстройки, для выполнения которых используется функция JavaScript, а не отображается пользовательский интерфейс. Элемент **FunctionFile** является дочерним для [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md). Атрибуту **resid** элемента **FunctionFile** присваивается значение атрибута **id** элемента **Url** в элементе **Resources**. Последний содержит URL-адрес HTML-файла, который содержит или загружает все функции JavaScript, используемые для выполнения команд надстройки без пользовательского интерфейса, как определено элементом [Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="8747c-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="8747c-106">Ниже приведен пример элемента **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="8747c-106">The following is an example of the  **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="8747c-107">Код JavaScript в HTML-файле, на который указывает элемент **FunctionFile**, должен вызывать `Office.initialize` и определять именованные функции, принимающие один параметр — `event`.</span><span class="sxs-lookup"><span data-stu-id="8747c-107">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="8747c-108">Функции должны использовать API `item.notificationMessages`, чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке.</span><span class="sxs-lookup"><span data-stu-id="8747c-108">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="8747c-109">Он также должен вызывать метод `event.completed` после выполнения.</span><span class="sxs-lookup"><span data-stu-id="8747c-109">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="8747c-110">Имена функций используются в элементе **FunctionName** для кнопок без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="8747c-110">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="8747c-111">Ниже приведен пример HTML-файла для определения функции **trackMessage**.</span><span class="sxs-lookup"><span data-stu-id="8747c-111">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

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

<span data-ttu-id="8747c-112">В примере кода ниже показано, как внедрить функцию, используемую элементом **FunctionName**.</span><span class="sxs-lookup"><span data-stu-id="8747c-112">The following code shows how to implement the function used by  **FunctionName**.</span></span>

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
> <span data-ttu-id="8747c-113">Вызов метода **event.completed** означает, что событие успешно обработано.</span><span class="sxs-lookup"><span data-stu-id="8747c-113">The call to  **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="8747c-114">Если функция вызывается несколько раз, например при многократном выборе одной команды надстройки, все события автоматически помещаются в очередь.</span><span class="sxs-lookup"><span data-stu-id="8747c-114">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="8747c-115">Первое событие запускается автоматически, тогда как остальные ожидают в очереди.</span><span class="sxs-lookup"><span data-stu-id="8747c-115">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="8747c-116">Когда функция вызывает метод **event.completed**, для нее запускается следующий вызов в очереди.</span><span class="sxs-lookup"><span data-stu-id="8747c-116">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="8747c-117">Если вы не реализуете вызов **event.completed**, функция не будет работать.</span><span class="sxs-lookup"><span data-stu-id="8747c-117">You must call **event.completed**; otherwise your function will not run.</span></span>
