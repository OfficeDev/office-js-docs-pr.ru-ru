---
title: Просмотр или изменение места встречи в надстройке
description: Узнайте, как просмотреть и изменить место проведения встречи в надстройке Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 79cf5ebe029d2b95b1501b6f9066a2c8f9013ef3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609185"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Просмотр или изменение расположения при создании встречи в Outlook

API JavaScript для Office предоставляет свойства и методы для управления расположением встречи, создаваемой пользователем. В настоящее время существует два свойства, которые предоставляют место встречи:

- [Item. Location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): базовый API, с помощью которого можно получить и задать расположение.
- [Item. енханцедлокатион](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Расширенный API, позволяющий получить и задать расположение, а также указать [тип расположения](/javascript/api/outlook/office.mailboxenums.locationtype). Тип `LocationType.Custom` задается в том случае, если вы задаете расположение с помощью `item.location` .

В следующей таблице перечислены API расположения и режимы (например, создание или чтение), где они доступны.

| API | Применяемые режимы встреч |
|---|---|
| [Item. Location](/javascript/api/outlook/office.appointmentread#location) | Участник или чтение |
| [Item. Location. Async](/javascript/api/outlook/office.location#getasync-options--callback-) | Организатор/создание |
| [item.location.setAsync](/javascript/api/outlook/office.location#setasync-location--options--callback-) | Организатор/создание |
| [Item. Енханцедлокатион. Async](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) | Органайзер/создание,<br>Участник или чтение |
| [Item. Енханцедлокатион. addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) | Организатор/создание |
| [Item. Енханцедлокатион. removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) | Организатор/создание |

Чтобы использовать методы, доступные только для создания надстроек, настройте манифест надстройки, чтобы активировать надстройку в режиме органайзера или создания. Более подробную информацию можно найти в статье [Создание надстроек Outlook для форм создания](compose-scenario.md) .

## <a name="use-the-enhancedlocation-api"></a>Использование `enhancedLocation` API

Вы можете использовать `enhancedLocation` API для получения и задания места встречи. Поле Location поддерживает несколько расположений, и для каждого местоположения можно задать отображаемое имя, тип и адрес электронной почты комнаты конференц-связи (если это возможно). Поддерживаемые типы расположений представлены в [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) .

### <a name="add-location"></a>Добавление расположения

В приведенном ниже примере показано, как добавить расположение, вызвав [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) для [Mailbox. Item. енханцедлокатион](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).

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

### <a name="get-location"></a>Получение расположения

В приведенном ниже примере показано, как получить расположение, вызвав метод [Async](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) для [Mailbox. Item. енханцедлокатион](/javascript/api/outlook/office.appointmentread#enhancedlocation).

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

В приведенном ниже примере показано, как удалить расположение, вызвав [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) для [Mailbox. Item. енханцедлокатион](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).

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

## <a name="use-the-location-api"></a>Использование `location` API

Вы можете использовать `location` API для получения и задания места встречи.

### <a name="get-the-location"></a>Получение места проведения

В этом разделе представлен пример кода, который получает и отображает место проведения создаваемой пользователем встречи.

Чтобы использовать метод `item.location.getAsync`, создайте метод обратного вызова, который проверяет состояние и результат асинхронного вызова. Вы можете указать все необходимые аргументы метода обратного вызова с помощью необязательного параметра `asyncContext`. Вы можете получать состояние, результаты и любые ошибки, используя выходной параметр `asyncResult` обратного вызова. Если асинхронный вызов успешно выполнен, вы можете получить место проведения в строковом формате с помощью свойства [AsyncResult.value](/javascript/api/office/office.asyncresult#value).

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

Чтобы использовать метод `item.location.setAsync`, укажите строку длиной до 255 символов в параметре data. При желании вы можете указать метод обратного вызова и его аргументы в параметре `asyncContext`. Необходимо проверить состояние, результат и любое сообщение об ошибке в `asyncResult` выходном параметре обратного вызова. Если асинхронный вызов успешно выполнен, `setAsync` вставляет указанную строку в виде обычного текста, заменяя существующее место проведения.

> [!NOTE]
> Можно задать несколько расположений, используя точку с запятой в качестве разделителя (например, "Конференц-зал A; Конференц-зал B ').

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

- [Создание первой надстройки Outlook](../quickstarts/outlook-quickstart.md)
- [Асинхронное программирование в случае надстроек Office](../develop/asynchronous-programming-in-office-add-ins.md)
