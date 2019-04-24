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
# <a name="functionfile-element"></a>Элемент FunctionFile

Указывает файл с исходным кодом для операций, доступных через те команды надстройки, для выполнения которых используется функция JavaScript, а не отображается пользовательский интерфейс. Элемент **FunctionFile** является дочерним для [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md). Атрибуту **resid** элемента **FunctionFile** присваивается значение атрибута **id** элемента **Url** в элементе **Resources**. Последний содержит URL-адрес HTML-файла, который содержит или загружает все функции JavaScript, используемые для выполнения команд надстройки без пользовательского интерфейса, как определено элементом [Control](control.md).

Ниже приведен пример элемента **FunctionFile**.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

Код JavaScript в HTML-файле, на который указывает элемент **FunctionFile**, должен вызывать `Office.initialize` и определять именованные функции, принимающие один параметр — `event`. Функции должны использовать API `item.notificationMessages`, чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке. Он также должен вызывать метод `event.completed` после выполнения. Имена функций используются в элементе **FunctionName** для кнопок без пользовательского интерфейса.

Ниже приведен пример HTML-файла для определения функции **trackMessage**.

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

В примере кода ниже показано, как внедрить функцию, используемую элементом **FunctionName**.

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
> Вызов метода **event.completed** означает, что событие успешно обработано. Если функция вызывается несколько раз, например при многократном выборе одной команды надстройки, все события автоматически помещаются в очередь. Первое событие запускается автоматически, тогда как остальные ожидают в очереди. Когда функция вызывает метод **event.completed**, для нее запускается следующий вызов в очереди. Если вы не реализуете вызов **event.completed**, функция не будет работать.
