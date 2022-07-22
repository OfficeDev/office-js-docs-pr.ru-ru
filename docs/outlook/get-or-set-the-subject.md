---
title: Просмотр или изменение темы в надстройке Outlook
description: Узнайте, как просмотреть и изменить тему сообщения или встречи в надстройке Outlook.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: cf221b03753fd76966eb5c6270da68e94abfe0f9
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66959072"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>Просмотр или изменение темы при создании встречи или сообщения в Outlook

API JavaScript для Office предоставляет асинхронные методы ([subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1)) и [subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))) для получения и задания темы встречи или сообщения, которые создает пользователь. Эти асинхронные методы доступны только для создания надстроек. Чтобы использовать эти методы, убедитесь, что манифест надстройки настроен соответствующим образом, чтобы Outlook активирует надстройку в формах создания.

Свойство **subject** доступно для чтения в формах создания и формах чтения встреч и сообщений. В форме чтения доступ к свойству можно получить напрямую из родительского объекта, например:

```js
item.subject
```

Но так как в форме создания и пользователь, и ваша надстройка могут вставлять или изменять тему одновременно, для получения темы необходимо использовать асинхронный метод **getAsync**, как показано ниже:

```js
item.subject.getAsync
```

Свойство **subject** доступно для записи только в формах создания, но не в формах чтения.

Как и большинство асинхронных методов в API JavaScript для Office, **getAsync** и **setAsync** принимают необязательные входные параметры. Дополнительные сведения об указании этих необязательных входных параметров см. в разделе "Передача необязательных параметров асинхронным методам" в асинхронном программировании в надстройки [Office](../develop/asynchronous-programming-in-office-add-ins.md).

## <a name="get-the-subject"></a>Получение темы

В этом разделе показан пример кода, получающий и отображающий тему создаваемой встречи или сообщения. В примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Чтобы использовать **item.subject.getAsync**, укажите функцию обратного вызова, которая проверяет состояние и результат асинхронного вызова. Вы можете указать любые необходимые аргументы функции обратного вызова с помощью  _необязательного параметра asyncContext_ . Состояние, результаты и сообщения об ошибках можно получить с помощью выходного параметра  _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, вы можете получить тему как текстовую строку, используя свойство [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

Чтобы использовать **item.subject.setAsync**, укажите строку до 255 символов в параметре данных. При необходимости можно указать функцию обратного вызова и любые аргументы для функции обратного вызова в  _параметре asyncContext_ . Следует проверить состояние, результат и наличие ошибок в выходном параметре  _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, **setAsync** вставляет указанную строку темы как обычный текст, перезаписывая существующую тему этого элемента.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    const today = new Date();
    let subject;

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
