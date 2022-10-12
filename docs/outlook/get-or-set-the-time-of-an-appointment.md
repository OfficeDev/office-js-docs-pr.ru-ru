---
title: Просмотр или изменение времени встречи в надстройке Outlook
description: Узнайте, как просмотреть и изменить время начала и окончания встречи в надстройке Outlook.
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: c7aa40fda15c613aca869af8b277d4deb6fbf833
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541235"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>Просмотр или изменение времени при создании встречи в Outlook

API JavaScript для Office предоставляет асинхронные методы ([Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) и [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))) для получения и задания времени начала или окончания встречи, которую создает пользователь. Эти асинхронные методы доступны только для создания надстроек. Чтобы использовать эти методы, убедитесь, что вы правильно настроили [XML-манифест](compose-scenario.md) надстройки для Outlook, чтобы активировать надстройку в формах создания, как описано в разделе "Создание надстроек Outlook для форм создания". Правила активации не поддерживаются в надстройки, использующие манифест [Teams для надстроек Office (предварительная версия).](../develop/json-manifest-overview.md)

The [start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:

```js
item.start
```

И в этом примере:

```js
item.end
```

Но так как в форме создания и пользователь, и ваша надстройка могут вставлять или изменять сведения о времени одновременно, для получения времени начала и окончания необходимо использовать асинхронный метод **getAsync**, как показано ниже:

```js
item.start.getAsync
```

И в следующем примере:

```js
item.end.getAsync
```

Как и большинство асинхронных методов в API JavaScript для Office, **getAsync** и **setAsync** принимают необязательные входные параметры. Дополнительные сведения об указании последних см. в разделе [Передача дополнительных параметров в асинхронные методы](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) статьи [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).

## <a name="get-the-start-or-end-time"></a>Получение времени начала или окончания

This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.

```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
```

Чтобы использовать **item.start.getAsync** или **item.end.getAsync**, укажите функцию обратного вызова, которая проверяет состояние и результат асинхронного вызова. Вы можете указать любые необходимые аргументы функции обратного вызова с помощью  _необязательного параметра asyncContext_ . Состояние, результаты и сообщения об ошибках можно получить с помощью выходного параметра  _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, вы можете получить начальное время как объект **Date** в формате UTC, используя свойство [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="set-the-start-or-end-time"></a>Установка времени начала или окончания

This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.

Как и в предыдущем примере, здесь предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи.

Чтобы использовать **item.start.setAsync** или **item.end.setAsync**, укажите  значение даты в формате UTC в _параметре dateTime_. Если вы получаете дату на основе данных, введенных пользователем в клиенте, с помощью [mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) можно преобразовать полученное значение в объект **Date** в формате UTC. В параметре _asyncContext_ можно указать необязательную функцию обратного вызова и любые аргументы для функции обратного вызова. Следует проверить состояние, результат и наличие ошибок в выходном параметре  _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, **setAsync** вставляет указанное строку времени начала или окончания как обычный текст, перезаписывая существующее время начала или окончания для этого элемента.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    const startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
                // Do whatever appropriate for your scenario
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

- [Просмотр и изменение данных элемента в форме создания элементов Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Просмотр и изменение данных элемента Outlook в формах просмотра и создания](item-data.md)
- [Создание надстроек Outlook для форм создания](compose-scenario.md)
- [Асинхронное программирование надстроек Office](../develop/asynchronous-programming-in-office-add-ins.md)
- [Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook](get-set-or-add-recipients.md)  
- [Просмотр или изменение темы при создании встречи или сообщения в Outlook](get-or-set-the-subject.md)
- [Вставка данных в текст при создании встречи или сообщения в Outlook](insert-data-in-the-body.md)
- [Считывание и запись расположения при создании встречи в Outlook](get-or-set-the-location-of-an-appointment.md)
