---
title: Просмотр или изменение темы в надстройке Outlook
description: Узнайте, как просмотреть и изменить тему сообщения или встречи в надстройке Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 93864aee005af61d9648c39402a843d9105bb021
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325442"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>Просмотр или изменение темы при создании встречи или сообщения в Outlook

API JavaScript для Office предоставляет асинхронные методы ([subject. Async](/javascript/api/outlook/office.Subject#getasync-options--callback-) и [subject. setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)) для получения и задания темы встречи или сообщения, создаваемого пользователем. Эти асинхронные методы доступны только для создания надстроек. Чтобы использовать эти методы, убедитесь, что вы правильно настроили манифест надстройки в Outlook для активации надстройки в формах создания.

Свойство **subject** доступно для чтения в формах создания и формах чтения встреч и сообщений. В форме чтения доступ к свойству можно получить напрямую из родительского объекта, например:

```js
item.subject
```

Но так как в форме создания и пользователь, и ваша надстройка могут вставлять или изменять тему одновременно, для получения темы необходимо использовать асинхронный метод **getAsync**, как показано ниже:

```js
item.subject.getAsync
```

Свойство **subject** доступно для записи только в формах создания, но не в формах чтения.

Как и в случае с большинством асинхронных методов в API JavaScript для Office, методы SetAsync и- **Async** и **** принимают необязательные входные параметры. Дополнительные сведения об указании этих дополнительных входных параметров можно найти в разделе "Передача необязательных параметров асинхронным методам" в надстройках [Office в асинхронном программировании](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-the-subject"></a>Получение темы

В этом разделе показан пример кода, получающий и отображающий тему создаваемой встречи или сообщения. В примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

Чтобы использовать метод **item.subject.getAsync**, предоставьте метод обратного вызова, который проверяет состояние и результат асинхронного вызова. Вы можете указать любые необходимые аргументы метода обратного вызова с помощью дополнительного параметра  _asyncContext_. Состояние, результаты и сообщения об ошибках можно получить с помощью выходного параметра _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, вы можете получить тему как текстовую строку, используя свойство [AsyncResult.value](/javascript/api/office/office.asyncresult#value).


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-the-subject"></a>Установка темы


В этом разделе показан пример кода, задающий тему создаваемой встречи или сообщения. Как и в предыдущем примере, предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения.

Чтобы использовать метод **item.subject.setAsync**, укажите строку длиной до 255 символов в параметре data. При необходимости можно предоставить метод обратного вызова и все его аргументы в параметре _asyncContext_. Следует проверить состояние, результат и наличие ошибок в выходном параметре _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, **setAsync** вставляет указанную строку темы как обычный текст, перезаписывая существующую тему этого элемента.

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
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
- [Вставка данных в текст при создании встречи или сообщения в Outlook](insert-data-in-the-body.md)   
- [Просмотр или изменение расположения при создании встречи в Outlook](get-or-set-the-location-of-an-appointment.md) 
- [Просмотр или изменение времени при создании встречи в Outlook](get-or-set-the-time-of-an-appointment.md)
    
