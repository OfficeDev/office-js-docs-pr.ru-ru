---
title: Просмотр или изменение времени встречи в надстройке Outlook
description: Узнайте, как просмотреть и изменить время начала и окончания встречи в надстройке Outlook.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: ab4016923a883a259a3c9c478639ae288b1ebdf7
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154121"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>Просмотр или изменение времени при создании встречи в Outlook

API Office JavaScript предоставляет асинхронные методы[(Time.getAsync](/javascript/api/outlook/office.time#getAsync_options__callback_) и [Time.setAsync),](/javascript/api/outlook/office.time#setAsync_dateTime__options__callback_)чтобы получить и установить время начала или окончания встречи, которую создает пользователь. Эти асинхронные методы доступны только для составить надстройки. Чтобы использовать эти методы, убедитесь, что манифест надстройки был настроен соответствующим образом для Outlook, чтобы активировать надстройки в композитных формах, как описано в create [Outlook](compose-scenario.md)надстройки для создания форм .

Свойства [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) доступны для встреч в формах создания и чтения. в форме чтения доступ к свойствам можно получить напрямую из родительского объекта, как в следующем примере:

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

Как и большинство асинхронных методов в API javaScript Office JavaScript, **getAsync** и **setAsync** принимают необязательные параметры ввода. Дополнительные сведения об указании последних см. в разделе [Передача дополнительных параметров в асинхронные методы](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) статьи [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-the-start-or-end-time"></a>Получение времени начала или окончания

В этом разделе показан пример кода, который получает время начала встречи, создаваемой пользователем, и отображает его. Вы можете использовать тот же код, заменив свойство **start** на **end**, чтобы получить время окончания. В этом примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи, как показано ниже.


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Чтобы использовать методы **item.start.getAsync** и **item.end.getAsync**, предоставьте метод обратного вызова, который проверяет состояние и результат асинхронного вызова. Вы можете указать любые необходимые аргументы метода обратного вызова с помощью дополнительного параметра _asyncContext_. Состояние, результаты и сообщения об ошибках можно получить с помощью выходного параметра _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, вы можете получить начальное время как объект **Date** в формате UTC, используя свойство [AsyncResult.value](/javascript/api/office/office.asyncresult#value).


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
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

В этом разделе показан пример кода, получающий время начало встречи, создаваемой пользователем. Можно использовать тот же код, заменив свойство **start** на **end**, чтобы получить время начала. Обратите внимание, что если у формы создания уже есть время начала, последующая установка времени начала приведет к изменению времени окончания, чтобы сохранить предыдущую длительность встречи. Если у формы создания уже есть время окончания, последующая установка времени окончания приведет к изменению длительности и времени окончания. Если встреча создана как событие на весь день, установки времени начала приведет к смещению времени окончания на 24 часа и отмены выбора параметра события на весь день в форме создания.

Как и в предыдущем примере, здесь предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи.

Чтобы использовать методы **item.start.setAsync** и **item.end.setAsync**, укажите значение **Date** в формате UTC в параметре _dateTime_. Если вы получаете дату на основе данных, введенных пользователем в клиенте, с помощью [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) можно преобразовать полученное значение в объект **Date** в формате UTC. Можно предоставить необязательный метод обратного вызова и все его аргументы в параметре _asyncContext_. Следует проверить состояние, результат и наличие ошибок в выходном параметре _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, **setAsync** вставляет указанное строку времени начала или окончания как обычный текст, перезаписывая существующее время начала или окончания для этого элемента.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    var startDate = new Date("September 27, 2012 12:30:00");
    
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
- [Просмотр или изменение расположения при создании встречи в Outlook](get-or-set-the-location-of-an-appointment.md)
    
