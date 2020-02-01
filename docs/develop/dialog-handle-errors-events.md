---
title: Обработка ошибок и событий в диалоговом окне "Office"
description: Описывает перехват и обработку ошибок при открытии и использовании диалогового окна "Office"
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: a35131a46dc9f5edc18df37495abe5d8c2c5ad2a
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650121"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a>Обработка ошибок и событий в диалоговом окне "Office"

В этой статье описывается, как выполнять перехват и обработку ошибок при открытии диалогового окна и ошибок, происходящих в диалоговом окне.

> [!NOTE]
> В этой статье предсказано, что вы знакомы с основами использования API диалоговых окон Office, описанных в статье [Использование API диалоговых окон Office в](dialog-api-in-office-add-ins.md)надстройках Office.
> 
> Кроме того, вы можете ознакомиться [с рекомендациями и правилами для API диалоговых окон Office](dialog-best-practices.md).

Код должен обрабатывать две категории событий:

- Ошибки, возвращаемые при вызове метода `displayDialogAsync`, так как не удается создать диалоговое окно.
- Ошибки и другие события в диалоговом окне.

## <a name="errors-from-displaydialogasync"></a>Ошибки метода displayDialogAsync

В дополнение к общим ошибкам платформы и системы, четыре ошибки относятся к вызову `displayDialogAsync`.

|Цифровой код|Значение|
|:-----|:-----|
|12004|Домен URL-адреса, передаваемого в метод `displayDialogAsync`, не является доверенным. Домен должен быть таким же, как и для главной страницы (а также протокол и номер порта).|
|12005|URL-адрес, передаваемый в метод `displayDialogAsync`, использует протокол HTTP. Необходим протокол HTTPS. (В некоторых версиях Office текст сообщения об ошибке, возвращенный с 12005, совпадает с указанным для 12004.)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|Диалоговое окно уже открыто из этого главного окна. Для главного окна, например области задач, невозможно открыть сразу несколько диалоговых окон.|
|12009|Пользователь проигнорировал диалоговое окно. Эта ошибка может возникать в Office в Интернете, где пользователи могут отказаться от того, чтобы надстройка не могла показать диалоговое окно. Дополнительные сведения см в разделе [Обработка блокирования всплывающих окон с помощью Office в Интернете](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).|

Когда `displayDialogAsync` вызывается, объект [asyncResult](/javascript/api/office/office.asyncresult) передается в функцию обратного вызова. При успешном вызове открывается диалоговое окно, и `value` свойство `AsyncResult` объекта является объектом [диалогового окна](/javascript/api/office/office.dialog) . Например, в [диалоговом окне "отправить сведения" на страницу узла](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page). Когда вызов завершается `displayDialogAsync` с ошибкой, диалоговое окно не создается `status` , свойству `AsyncResult` объекта присваивается значение `Office.AsyncResultStatus.Failed`, и `error` свойство объекта заполняется. Всегда следует предоставлять обратный вызов, который проверяет `status` и отвечает на сообщение об ошибке. Пример, в котором сообщается о сообщении об ошибке независимо от его кода, представлен в приведенном ниже коде. ( `showNotification` Функция, не определенная в этой статье, либо отображает ошибку, либо заносит ее в журнал. Пример реализации этой функции в надстройке приведен в статье [Пример использования API диалоговых окон надстроек Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

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

Три ошибки и события в диалоговом окне вызывают `DialogEventReceived` событие на главной странице. Напоминание о странице ведущего приложения можно узнать в разделе [Открытие диалогового окна на странице узла](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).

|Цифровой код|Значение|
|:-----|:-----|
|12002|Одно из следующих:<br> – По URL-адресу, переданному в `displayDialogAsync`, не существует страницы.<br> — Страница, которая была `displayDialogAsync` перезагружена, но диалоговое окно было перенаправлено на страницу, которая не может быть найдена или загружена, или она направлена на URL-адрес с недопустимым синтаксисом.|
|12003|Выполнена попытка открыть из диалогового окна страницу, для URL-адреса которой используется протокол HTTP. Необходим протокол HTTPS.|
|12006|Диалоговое окно было закрыто, как правило, потому что пользователь выбрал кнопку **закрытия** **X**.|

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

Ниже приведен пример обработчика для события `DialogEventReceived`, который создает особые сообщения об ошибках для каждого кода ошибки.

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
