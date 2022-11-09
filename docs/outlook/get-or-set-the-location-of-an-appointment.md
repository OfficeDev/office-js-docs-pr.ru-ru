---
title: Просмотр или изменение места встречи в надстройке
description: Узнайте, как просмотреть и изменить место проведения встречи в надстройке Outlook.
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: d88e2494592d9b261945ecdaf0ca27ae79c73ba8
ms.sourcegitcommit: cae583433e489a3b71418ea270a90db72ad1e838
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/09/2022
ms.locfileid: "68892366"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Просмотр или изменение расположения при создании встречи в Outlook

API JavaScript для Office предоставляет свойства и методы для управления расположением встречи, которую создает пользователь. В настоящее время существует два свойства, которые предоставляют расположение встречи:

- [item.location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties): базовый API, позволяющий получить и задать расположение.
- [item.enhancedLocation](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties): расширенный API, который позволяет получить и задать расположение и включает указание [типа расположения](/javascript/api/outlook/office.mailboxenums.locationtype). Тип имеет значение , `LocationType.Custom` если вы задали расположение с помощью `item.location`.

В следующей таблице перечислены API расположения и режимы (т. е. Создание или Чтение), где они доступны.

| API | Применимые режимы встреч |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member) | Участник/чтение |
| [item.location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) | Организатор или создание |
| [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) | Организатор или создание |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) | Организатор или создание,<br>Участник/чтение |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) | Организатор или создание |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) | Организатор или создание |

Чтобы использовать методы, доступные только для создания надстроек, настройте XML-манифест надстройки для активации надстройки в режиме организатора или создания. Дополнительные сведения см. [в статье Создание надстроек Outlook для форм создания](compose-scenario.md) . Правила активации не поддерживаются в надстройках, использующих [манифест Teams для надстроек Office (предварительная версия).](../develop/json-manifest-overview.md)

## <a name="use-the-enhancedlocation-api"></a>`enhancedLocation` Использование API

С помощью `enhancedLocation` API можно получить и задать расположение встречи. Поле location поддерживает несколько расположений. Для каждого расположения можно задать отображаемое имя, тип и адрес электронной почты конференц-зала (если применимо). Поддерживаемые типы расположений см. в разделе [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) .

### <a name="add-location"></a>Добавление расположения

В следующем примере показано, как добавить расположение путем вызова [addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) в [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member).

```js
let item;
const locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a>Получение расположения

В следующем примере показано, как получить расположение, вызвав [метод getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) в [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-enhancedlocation-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

> [!NOTE]
> [Личные группы контактов](https://support.microsoft.com/office/88ff6c60-0a1d-4b54-8c9d-9e1a71bc3023) , добавленные в качестве расположений встреч, не возвращаются методом [enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) .

### <a name="remove-location"></a>Удалить расположение

В следующем примере показано, как удалить расположение, вызвав [метод removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) в [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a>`location` Использование API

С помощью `location` API можно получить и задать расположение встречи.

### <a name="get-the-location"></a>Получение места проведения

В этом разделе представлен пример кода, который получает и отображает место проведения создаваемой пользователем встречи.

Чтобы использовать `item.location.getAsync`, предоставьте функцию обратного вызова, которая проверяет состояние и результат асинхронного вызова. Вы можете указать любые необходимые аргументы для функции обратного вызова с помощью необязательного `asyncContext` параметра. Состояние, результаты и любую ошибку можно получить с помощью выходного параметра `asyncResult` обратного вызова. Если асинхронный вызов успешно выполнен, вы можете получить место проведения в строковом формате с помощью свойства [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

Чтобы использовать метод `item.location.setAsync`, укажите строку длиной до 255 символов в параметре data. При необходимости можно указать функцию обратного вызова и любые аргументы для функции обратного вызова в параметре `asyncContext` . Необходимо проверить состояние, результат и любое сообщение об ошибке в выходном `asyncResult` параметре обратного вызова. Если асинхронный вызов успешно выполнен, `setAsync` вставляет указанную строку в виде обычного текста, заменяя существующее место проведения.

> [!NOTE]
> Можно задать несколько расположений, используя точку с запятой в качестве разделителя (например, "Конференц-зал A; Конференц-зал B').

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready method.
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
- [Асинхронное программирование надстроек Office](../develop/asynchronous-programming-in-office-add-ins.md)
