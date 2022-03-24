---
title: Обработка ошибок и событий в диалоговом окне "Office"
description: Узнайте, как улавливать и обрабатывать ошибки при открытии Office диалоговом окне.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 96bb2189ccf9b9ef6c976bb154746368c5bde69a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743814"
---
# <a name="handle-errors-and-events-in-the-office-dialog-box"></a>Обработка ошибок и событий в диалоговом Office диалоговом окне

В этой статье описывается, как улавливать и обрабатывать ошибки при открытии диалоговое окно и ошибки, которые происходят в диалоговом окне.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с основами использования API диалогового Office, как описано в статье [Использование API](dialog-api-in-office-add-ins.md) диалогового Office в Office надстройки.
> 
> См. [также лучшие практики и правила для API Office диалоговом ок.](dialog-best-practices.md)

Код должен обрабатывать две категории событий.

- Ошибки, возвращаемые при вызове метода `displayDialogAsync`, так как не удается создать диалоговое окно.
- Ошибки и другие события в диалоговом окне.

## <a name="errors-from-displaydialogasync"></a>Ошибки метода displayDialogAsync

Помимо общих ошибок платформы и системы, четыре ошибки являются специфическими для вызова `displayDialogAsync`.

|Цифровой код|Значение|
|:-----|:-----|
|12004|Домен URL-адреса, передаваемого в метод `displayDialogAsync`, не является доверенным. Домен должен быть таким же, как и для главной страницы (а также протокол и номер порта).|
|12005|URL-адрес, передаваемый в метод `displayDialogAsync`, использует протокол HTTP. Необходим протокол HTTPS. (В некоторых версиях Office текст сообщения об ошибке, возвращенный с 12005, является тем же, что и для 12004 года.)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|Диалоговое окно уже открыто из этого главного окна. Для главного окна, например области задач, невозможно открыть сразу несколько диалоговых окон.|
|12009|Пользователь проигнорировал диалоговое окно. Эта ошибка может возникнуть в Office в Интернете, когда пользователи могут не разрешать надстройке представлять диалоговое окно. Дополнительные сведения см. в [ссылке Обработка](dialog-best-practices.md#handle-pop-up-blockers-with-office-on-the-web) всплывающих блокаторов с помощью Office в Интернете.|

При `displayDialogAsync` вызове он передает объект [AsyncResult](/javascript/api/office/office.asyncresult) функцию вызова. При успешном вызове открывается диалоговое окно, `value` `AsyncResult` а свойством объекта является [объект Dialog](/javascript/api/office/office.dialog) . В этом примере см. статью Отправка сведений из [диалогового окна на хост-страницу](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page). При сбой `displayDialogAsync` вызова диалоговое окно не создается, `Office.AsyncResultStatus.Failed``status` `AsyncResult` задается свойство объекта и `error` заполняется свойство объекта. Всегда необходимо предоставить вызов, который проверяет `status` ошибку и отвечает на нее. Пример сообщения об ошибке независимо от номера кода см. в следующем коде. (Функция `showNotification` , не заданная в этой статье, отображает или регистрит ошибку. Пример реализации этой функции в надстройки см. в Office [API надстройки](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) диалогов.)

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

Три ошибки и события в диалоговом окне поднимут `DialogEventReceived` событие на хост-странице. Чтобы напомнить о том, что такое хост-страница, см. в странице Откройте диалоговое [окно с хост-страницы](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).

|Цифровой код|Значение|
|:-----|:-----|
|12002|Одно из следующих:<br> – По URL-адресу, переданному в `displayDialogAsync`, не существует страницы.<br> - Страница, `displayDialogAsync` которая была передана для загрузки, но диалоговое окно было перенаправлено на страницу, которую она не может найти или загрузить, или она была направлена на URL-адрес с недействительным синтаксис.|
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

Пример обработки `DialogEventReceived` события, создав настраиваемые сообщения об ошибке для каждого кода ошибки, см. в следующем примере.

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

## <a name="see-also"></a>См. также

Надстройку с такой обработкой ошибок см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).
