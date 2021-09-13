---
title: Элемент FunctionFile в файле манифеста
description: Указывает исходный файл кода для операций, которые надстройка предоставляет с помощью команд надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: 443fde5cc5456508556962254ecceb6bd717e8a8
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154799"
---
# <a name="functionfile-element"></a>Элемент FunctionFile

Указывает исходный код файла для операций, которые надстройка предоставляет одним из следующих способов.

* Команды надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.
* Клавиши, которые выполняют функцию JavaScript.

Элемент `FunctionFile` является детским элементом [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md). Атрибут элемента может быть не более 32 символов и задарен значению атрибута элемента в элементе, который содержит `resid` `FunctionFile` `id` `Url` `Resources` URL-адрес HTML-файла, [](control.md)который содержит или загружает все функции JavaScript, используемые кнопками командной команды без пользовательского интерфейса, как это определено элементом Control.

Ниже приводится пример `FunctionFile` элемента.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

JavaScript в HTML-файле, указанный элементом, должен вызывать и определять именные функции, которые принимают `FunctionFile` `Office.initialize` один параметр: `event` . Функции должны использовать API `item.notificationMessages`, чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке. Он также должен вызывать метод `event.completed` после выполнения. Имя функций используется в элементе для кнопок без `FunctionName` пользовательского интерфейса.

Ниже приводится пример HTML-файла, определяющий `trackMessage` функцию.

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

В следующем коде показано, как реализовать используемую функцию `FunctionName` .

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
> Вызов `event.completed` сигналов об успешном обращении с событием. Если функция вызывается несколько раз, например при выборе одной команды надстройки несколько раз, все события автоматически помещаются в очередь. Первое событие запускается автоматически, тогда как остальные ожидают в очереди. При вызове функции выполняется следующий вызов в очереди на `event.completed` эту функцию. Необходимо `event.completed` вызвать; в противном случае функция не будет работать.
