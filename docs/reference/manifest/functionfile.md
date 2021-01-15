---
title: Элемент FunctionFile в файле манифеста
description: Указывает файл исходных кодов для операций, которые надстройка предоставляет с помощью команд надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 44bfd514025b8a23f4f6acdf3fec004485ca4c5a
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771402"
---
# <a name="functionfile-element"></a>Элемент FunctionFile

Указывает файл исходных кодов для операций, которые предоставляет надстройка одним из следующих способов:

* Команды надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.
* Сочетания клавиш, которые выполняют функцию JavaScript.

Этот `FunctionFile` элемент является элементом [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor.](mobileformfactor.md) Атрибут элемента не может быть больше 32 символов и имеет значение атрибута элемента, который содержит `resid` `FunctionFile` URL-адрес `id` `Url` `Resources` HTML-файла, [](control.md)который содержит или загружает все функции JavaScript, используемые кнопками команд надстройки без пользовательского интерфейса, как определено элементом Control.

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

Код JavaScript в HTML-файле, определенном элементом, должен вызывать и определять именующие функции, которые `FunctionFile` `Office.initialize` принимают один параметр: `event` . Функции должны использовать API `item.notificationMessages`, чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке. Он также должен вызывать метод `event.completed` после выполнения. Имя функций используется в элементе для кнопок без пользовательского `FunctionName` интерфейса.

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

В следующем коде показано, как реализовать функцию, используемую `FunctionName` .

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
> Вызов `event.completed` сигнализирует о том, что событие успешно обработано. Если функция вызывается несколько раз, например при выборе одной команды надстройки несколько раз, все события автоматически помещаются в очередь. Первое событие запускается автоматически, тогда как остальные ожидают в очереди. При вызове функции запускается следующий вызов в очереди `event.completed` для этой функции. Необходимо `event.completed` вызвать; в противном случае функция не будет работать.
