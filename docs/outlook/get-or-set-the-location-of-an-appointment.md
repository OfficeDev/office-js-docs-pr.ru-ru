---
title: Просмотр или изменение места встречи в надстройке
description: Узнайте, как просмотреть и изменить место проведения встречи в надстройке Outlook.
ms.date: 10/31/2019
ms.localizationpriority: medium
ms.openlocfilehash: d409019dd5d99a27c030bea825c108f65fc35567
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484599"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Просмотр или изменение расположения при создании встречи в Outlook

API Office JavaScript предоставляет свойства и методы для управления расположением записи, которую создает пользователь. В настоящее время существует два свойства, которые предоставляют расположение встречи:

- [item.location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties). Базовый API, который позволяет получать и устанавливать расположение.
- [item.enhancedLocation](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties). Расширенный API, который позволяет получать и устанавливать расположение, а также включает указание [типа расположения](/javascript/api/outlook/office.mailboxenums.locationtype). Тип, если `LocationType.Custom` вы установите расположение с помощью `item.location`.

В следующей таблице перечислены API расположения и режимы (например, Compose или Read), где они доступны.

| API | Применимые режимы назначения |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member) | Чел/чтение |
| [item.location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) | Организатор/композит |
| [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) | Организатор/композит |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) | Организатор/композит,<br>Чел/чтение |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) | Организатор/композит |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) | Организатор/композит |

Чтобы использовать методы, доступные только для составить надстройки, настройте манифест надстройки, чтобы активировать надстройку в режиме Organis/Compose. [Дополнительные сведения см. в Outlook Create](compose-scenario.md) Outlook надстройки для создания форм.

## <a name="use-the-enhancedlocation-api"></a>`enhancedLocation` Использование API

API можно использовать `enhancedLocation` для получения и определения расположения встречи. Поле расположения поддерживает несколько расположений, и для каждого расположения можно установить имя отображения, тип и адрес электронной почты конференц-зала (если это применимо). См [. в этой ленте LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) для поддерживаемых типов расположения.

### <a name="add-location"></a>Добавление расположения

В следующем примере показано, как добавить расположение, позвонив [addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) на [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member).

```js
var item;
var locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a>Расположение

В следующем примере показано, как получить расположение, позвонив [в getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) на [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-enhancedlocation-member).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (place) {
        console.log("Display name: " + place.displayName);
        console.log("Type: " + place.locationIdentifier.type);
        if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
            console.log("Email address: " + place.emailAddress);
        }
    });
}
```

### <a name="remove-location"></a>Удаление расположения

В следующем примере показано, как удалить расположение, позвонив [removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) в [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        Office.context.mailbox.item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a>`location` Использование API

API можно использовать `location` для получения и определения расположения встречи.

### <a name="get-the-location"></a>Получение места проведения

В этом разделе представлен пример кода, который получает и отображает место проведения создаваемой пользователем встречи.

Чтобы использовать метод `item.location.getAsync`, создайте метод обратного вызова, который проверяет состояние и результат асинхронного вызова. Вы можете указать все необходимые аргументы метода обратного вызова с помощью необязательного параметра `asyncContext`. Вы можете получить состояние, результаты и любую ошибку с помощью параметра `asyncResult` вывода вызова. Если асинхронный вызов успешно выполнен, вы можете получить место проведения в строковом формате с помощью свойства [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="set-the-location"></a>Установка места проведения

В этом разделе показан пример кода, который устанавливает место проведения создаваемой пользователем встречи.

Чтобы использовать метод `item.location.setAsync`, укажите строку длиной до 255 символов в параметре data. При желании вы можете указать метод обратного вызова и его аргументы в параметре `asyncContext`. Необходимо проверить состояние, результат и любое сообщение об ошибке в параметре `asyncResult` вывода вызова. Если асинхронный вызов успешно выполнен, `setAsync` вставляет указанную строку в виде обычного текста, заменяя существующее место проведения.

> [!NOTE]
> Вы можете установить несколько местоположений, используя полу-двоеточие в качестве сепаратора (например, "Конференц-зал A; Конференц-зал B').

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever is appropriate for your scenario,
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

## <a name="see-also"></a>См. также

- [Создание первой Outlook надстройки](../quickstarts/outlook-quickstart.md)
- [Асинхронное программирование в случае надстроек Office](../develop/asynchronous-programming-in-office-add-ins.md)
