---
title: Обработка ошибок и событий в диалоговом окне "Office"
description: Узнайте, как перехватывать и обрабатывать ошибки при открытии и использовании диалогового окна Office.
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: d3bdae7d4dddcd92a54a46fec0d5854a1a18a0bc
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616037"
---
# <a name="handle-errors-and-events-in-the-office-dialog-box"></a>Обработка ошибок и событий в диалоговом окне Office

В этой статье описывается, как перехватывать и обрабатывать ошибки при открытии диалогового окна и ошибки, которые происходят в диалоговом окне.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с основами использования API диалогов Office, как описано в статье "Использование API диалогов Office в надстройке [Office"](dialog-api-in-office-add-ins.md).
>
> См. [также рекомендации и правила для API диалогового окна Office](dialog-best-practices.md).

Код должен обрабатывать две категории событий.

- Ошибки, возвращаемые при вызове метода `displayDialogAsync`, так как не удается создать диалоговое окно.
- Ошибки и другие события в диалоговом окне.

## <a name="errors-from-displaydialogasync"></a>Ошибки метода displayDialogAsync

Помимо общих ошибок платформы и системы, четыре ошибки относятся к вызову `displayDialogAsync`.

|Цифровой код|Значение|
|:-----|:-----|
|12004|Домен URL-адреса, передаваемого в метод `displayDialogAsync`, не является доверенным. Домен должен быть таким же, как и для главной страницы (а также протокол и номер порта).|
|12005|URL-адрес, передаваемый в метод `displayDialogAsync`, использует протокол HTTP. Необходим протокол HTTPS. (В некоторых версиях Office текст сообщения об ошибке, возвращаемый с кодом 12005, совпадает с текстом 12004.)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|Диалоговое окно уже открыто из этого главного окна. Для главного окна, например области задач, невозможно открыть сразу несколько диалоговых окон.|
|12009|Пользователь проигнорировал диалоговое окно. Эта ошибка может возникнуть в Office в Интернете, когда пользователи могут не разрешать надстройке отображать диалоговое окно. Дополнительные сведения см. в [разделе "Обработка блокаторов](dialog-best-practices.md#handle-pop-up-blockers-with-office-on-the-web) всплывающих элементов с помощью Office в Интернете".|
|12011| Надстройка работает в Office в Интернете и конфигурация браузера пользователя блокирует всплывающие окна. Чаще всего это происходит, когда браузер является устаревшей версией Edge, а домен надстройки находится в зоне безопасности, которая отличается от домена, который диалоговое окно пытается открыть. Другой сценарий, который вызывает эту ошибку, — браузер Safari и настроенный для блокировки всех всплывающих элементов. Рассмотрите возможность ответа на эту ошибку с запросом на изменение конфигурации браузера или использование другого браузера.|

При `displayDialogAsync` вызове объект [AsyncResult](/javascript/api/office/office.asyncresult) передается в функцию обратного вызова. При успешном вызове открывается диалоговое окно, `value` `AsyncResult` а свойством объекта является [объект Dialog](/javascript/api/office/office.dialog) . Пример этого см. в разделе ["Отправка сведений из диалогового окна на страницу узла"](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page). При сбое `displayDialogAsync` вызова диалоговое окно не создается, `Office.AsyncResultStatus.Failed``status` `AsyncResult` свойство объекта устанавливается в значение и `error` свойство объекта заполняется. Всегда следует предоставлять обратный вызов, который проверяет `status` ошибку и отвечает на нее. Пример сообщения об ошибке независимо от номера кода см. в следующем коде. (Функция `showNotification` , не определенная в этой статье, отображает или регистрирует ошибку. Пример реализации этой функции в надстройке см. в примере [API диалогового окна надстройки Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)

```js
let dialog;
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

Три ошибки и события в диалоговом окне `DialogEventReceived` вызывают событие на странице узла. Напоминание о том, что такое хост-страница, см. в разделе ["Открытие диалогового окна с хост-страницы"](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).

|Цифровой код|Значение|
|:-----|:-----|
|12002|Одно из следующих:<br> – По URL-адресу, переданному в `displayDialogAsync`, не существует страницы.<br> — Страница, `displayDialogAsync` которая была передана для загрузки, но диалоговое окно было перенаправлено на страницу, которую не удается найти или загрузить, или она была направлена на URL-адрес с недопустимым синтаксисом.|
|12003|Выполнена попытка открыть из диалогового окна страницу, для URL-адреса которой используется протокол HTTP. Необходим протокол HTTPS.|
|12006|Диалоговое окно было закрыто, обычно потому, что пользователь нажмет кнопку **"Закрыть****" X**.|

Код может назначить обработчик для события `DialogEventReceived` при вызове `displayDialogAsync`. Ниже приведен простой пример.

```js
let dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

Пример обработчика `DialogEventReceived` события, создающий пользовательские сообщения об ошибках для каждого кода ошибки, см. в следующем примере.

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
