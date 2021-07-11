---
title: Обработка ошибок и событий в диалоговом окне "Office"
description: Описывает, как улавливать и обрабатывать ошибки при открытии Office диалоговом окне
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: be1fb8bcd30b47ac6399657d928d3cad7f857f39
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349898"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a>Обработка ошибок и событий в диалоговом окне "Office"

В этой статье описывается, как улавливать и обрабатывать ошибки при открытии диалоговое окно и ошибки, которые происходят в диалоговом окне.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с основами использования API диалогового Office, как описано в Статье [Использование API](dialog-api-in-office-add-ins.md)диалогов Office в Office надстройки .
> 
> См. также лучшие практики и правила для API Office [диалоговом ок.](dialog-best-practices.md)

Код должен обрабатывать две категории событий:

- Ошибки, возвращаемые при вызове метода `displayDialogAsync`, так как не удается создать диалоговое окно.
- Ошибки и другие события в диалоговом окне.

## <a name="errors-from-displaydialogasync"></a>Ошибки метода displayDialogAsync

Помимо общих ошибок платформы и системы, четыре ошибки являются специфическими для вызова `displayDialogAsync` .

|Цифровой код|Значение|
|:-----|:-----|
|12004|Домен URL-адреса, передаваемого в метод `displayDialogAsync`, не является доверенным. Домен должен быть таким же, как и для главной страницы (а также протокол и номер порта).|
|12005|URL-адрес, передаваемый в метод `displayDialogAsync`, использует протокол HTTP. Необходим протокол HTTPS. (В некоторых версиях Office текст сообщения об ошибке, возвращенный с 12005, является тем же, что и для 12004.)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|Диалоговое окно уже открыто из этого главного окна. Для главного окна, например области задач, невозможно открыть сразу несколько диалоговых окон.|
|12009|Пользователь проигнорировал диалоговое окно. Эта ошибка может возникнуть в Office в Интернете, когда пользователи могут не разрешать надстройке представлять диалоговое окно. Дополнительные сведения см. в ссылке [Обработка всплывающих](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)блокаторов с помощью Office в Интернете .|

При `displayDialogAsync` вызове он передает объект [AsyncResult](/javascript/api/office/office.asyncresult) функцию вызова. При успешном вызове открывается диалоговое окно, а свойством объекта является `value` `AsyncResult` объект [Dialog.](/javascript/api/office/office.dialog) В этом примере см. статью [Отправка сведений из диалогового окна на хост-страницу.](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page) При сбой вызова диалоговое окно не создается, задается свойство объекта `displayDialogAsync` `status` и `AsyncResult` `Office.AsyncResultStatus.Failed` `error` заполняется свойство объекта. Всегда необходимо предоставить вызов, который проверяет ошибку и отвечает на `status` нее. Пример сообщения об ошибке независимо от номера кода см. в следующем коде. `showNotification`(Функция, не заданная в этой статье, отображает или регистрит ошибку. Пример реализации этой функции в надстройки см. в Office примере [API диалогов надстройки.)](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

## <a name="errors-and-events-in-the-dialog-box"></a>Ошибки и события в диалоговом окне

Три ошибки и события в диалоговом окне поднимут `DialogEventReceived` событие на хост-странице. Напоминая о том, что такое хост-страница, см. в странице Откройте диалоговое [окно с хост-страницы.](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)

|Цифровой код|Значение|
|:-----|:-----|
|12002|Одно из следующих:<br> – По URL-адресу, переданному в `displayDialogAsync`, не существует страницы.<br> - Страница, которая была передана для загрузки, но диалоговое окно было перенаправлено на страницу, которую она не может найти или загрузить, или она была направлена на URL-адрес с недействительным `displayDialogAsync` синтаксис.|
|12003|Выполнена попытка открыть из диалогового окна страницу, для URL-адреса которой используется протокол HTTP. Необходим протокол HTTPS.|
|12006|Диалоговое окно было закрыто, как правило, из-за того, что пользователь выбрал кнопку **Закрыть** **X**.|

Код может назначить обработчик для события `DialogEventReceived` при вызове `displayDialogAsync`. Ниже приведен простой пример.

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

Пример обработки события, создав настраиваемые сообщения об ошибке для каждого кода ошибки, см. `DialogEventReceived` в следующем примере.

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

Надстройку с такой обработкой ошибок см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).
