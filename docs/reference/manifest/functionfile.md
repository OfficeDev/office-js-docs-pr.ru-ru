---
title: Элемент FunctionFile в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eec1dc8eb2e099670469af6ef300592fc4a31e64
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324871"
---
# <a name="functionfile-element"></a>Элемент FunctionFile

Указывает файл исходного кода для операций, предоставляемых надстройкой через команды надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса. `FunctionFile` Элемент является дочерним для элемента [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md). Для `resid` `FunctionFile` атрибута элемента задается значение `id` атрибута `Url` элемента в `Resources` элементе, который содержит URL-адрес HTML-файла, который содержит или загружает все функции JavaScript, которые используются кнопками надстройки без пользовательского интерфейса, как определено [элементом Control](control.md).

Ниже приведен пример `FunctionFile` элемента.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

JavaScript в HTML-файле, указанном `FunctionFile` элементом, должен вызывать `Office.initialize` и определять именованные функции, которые принимают один параметр `event`:. Функции должны использовать API `item.notificationMessages`, чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке. Он также должен вызывать метод `event.completed` после выполнения. Имя функции используется в `FunctionName` элементе для кнопок без пользовательского интерфейса.

Ниже приведен пример HTML-файла, определяющего `trackMessage` функцию.

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

В приведенном ниже коде показано, как реализовать функцию, `FunctionName`используемую в.

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
> Вызов, `event.completed` сигнализирующий о том, что событие успешно обработано. Если функция вызывается несколько раз, например при выборе одной команды надстройки несколько раз, все события автоматически помещаются в очередь. Первое событие запускается автоматически, тогда как остальные ожидают в очереди. При вызове `event.completed`функции выполняется следующий вызов этой функции в очереди. Необходимо вызвать `event.completed`; в противном случае функция не будет выполняться.
