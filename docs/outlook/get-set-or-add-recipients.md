---
title: Просмотр и изменение получателей в надстройке Outlook
description: Узнайте, как просмотреть, изменить или добавить получателей сообщения или встречи в надстройке Outlook.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 396f425f639c0d7043154ccfe1ddea16a236f993
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325435"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook


API JavaScript для Office предоставляет асинхронные методы ([Recipients. Async](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients. setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)или [Recipients. addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)), чтобы соответственно получать, задавать и добавлять получателей в форме создания встречи или сообщения. Эти асинхронные методы доступны только для создания надстроек. Чтобы использовать эти методы, убедитесь, что вы правильно настроили манифест надстройки в Outlook для активации надстройки в формах создания, как описано в статье [Создание надстроек Outlook для форм создания](compose-scenario.md).

Некоторые свойства, представляющие получателей в сообщении или встрече, доступны для чтения в формах создания и чтения. Это свойства [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) для встреч, а также свойства [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) для сообщений.

В форме чтения доступ к свойству можно получить напрямую из родительского объекта, например:

```js
item.cc
```

Но в форме создания, так как пользователь и надстройка могут одновременно вставлять или изменять получателя, необходимо использовать асинхронный метод `getAsync` для получения этих свойств, как показано в следующем примере:


```js
item.cc.getAsync
```

Эти свойства доступны для записи только в формах создания, но не формах чтения.

Как и в случае с большинством асинхронных методов в API `getAsync`JavaScript `setAsync`для Office `addAsync` , и использовать необязательные входные параметры. Дополнительные сведения об указании последних см. в разделе [Передача дополнительных параметров в асинхронные методы](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) статьи [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-recipients"></a>Извлечение получателей


В этом разделе показан пример кода, извлекающий получателей создаваемой встречи или сообщения, и показывающий адреса электронной почты получателей. В примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

В API JavaScript для Office, так как свойства, представляющие получателей встречи ( **optionalAttendees** и **requiredAttendees**), отличаются от тех, которые относятся к сообщению ([BCC](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **CC**и **to**), необходимо сначала использовать свойство [Item. ItemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) , чтобы определить, является ли состав создаваемого элемента встречей или сообщением. В режиме создания все эти свойства встреч и сообщений являются объектами [Recipients](/javascript/api/outlook/office.Recipients) , поэтому можно применить асинхронный метод, `Recipients.getAsync`чтобы получить соответствующих получателей.

Для использования `getAsync` Предоставьте метод обратного вызова для проверки состояния, результатов и любой ошибки, возвращенной асинхронным `getAsync` вызовом. Методу обратного вызова можно передать любые аргументы, используя дополнительный параметр  _asyncContext_. Метод обратного вызова возвращает выходной параметр  _asyncResult_. Можно `status` использовать свойства `error` и объекта параметра [asyncResult](/javascript/api/office/office.asyncresult) , чтобы проверить состояние и все сообщения об ошибках асинхронного вызова, а также `value` свойство для получения фактических получателей. Они представляются как массив объектов [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails).

Обратите внимание на `getAsync` то, что метод является асинхронным, если существуют последующие действия, которые зависят от успешного извлечения получателей, необходимо организовать код для запуска таких действий только в соответствующем методе обратного вызова после успешного завершения асинхронного вызова.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-recipients"></a>Установка получателей


В этом разделе показан пример кода, устанавливающий получателей встречи или сообщения, создаваемых пользователем. При установке получателей все существующие получателей перезаписываются. Как и предыдущий пример, извлекающий получателей в форме создания, этот пример предполагает, что надстройка активируется в формах создания встреч и сообщений. В этом примере сначала проверяется, является ли состав созданного элемента встречей или сообщением, чтобы применить асинхронный метод, чтобы применить асинхронный метод `Recipients.setAsync`, в соответствующих свойствах, представляющих получателей встречи или сообщения.

При вызове `setAsync`укажите массив в качестве входного аргумента для параметра _Recipients_ в одном из следующих форматов:


- массив строк, являющихся SMTP-адресами;
    
- массив словарей, каждый из которых содержит отображаемое имя и адрес электронной почты, как показано в следующем примере кода;
    
- Массив `EmailAddressDetails` объектов, аналогичный элементу, возвращаемому `getAsync` методом.
    
При необходимости можно предоставить метод обратного вызова в качестве аргумента ввода для `setAsync` метода, чтобы убедиться, что код, зависящий от успешной настройки получателей, будет выполняться только в том случае, если это произойдет. Кроме того, методу обратного вызова можно передать любые аргументы, используя дополнительный параметр  _asyncContext_. При использовании метода обратного вызова можно получить доступ к выходному параметру _asyncResult_ и использовать свойства **Status** и **Error** объекта `AsyncResult` Parameter для проверки состояния и любых сообщений об ошибках асинхронного вызова.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## <a name="add-recipients"></a>Добавление получателей

Если вы не хотите перезаписать существующие получатели в встрече или сообщении, вместо использования `Recipients.setAsync`можно использовать `Recipients.addAsync` асинхронный метод для добавления получателей. `addAsync`работает так же `setAsync` , как и в случае, если требуется аргумент _Recipients_ . При необходимости можно предоставить метод обратного вызова и все аргументы для обратного вызова с помощью параметра asyncContext. Затем можно проверить состояние, результат и любую ошибку асинхронного `addAsync` вызова, используя выходной параметр _asyncResult_ метода обратного вызова. В следующем примере проверяется, является ли состав создаваемого элемента встречей, и добавляет в встречу двух обязательных участников.


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## <a name="see-also"></a>См. также

- [Просмотр и изменение данных элемента в форме создания элементов Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Просмотр и изменение данных элемента Outlook в формах просмотра и создания](item-data.md)
- [Создание надстроек Outlook для форм создания](compose-scenario.md)
- [Асинхронное программирование надстроек Office](../develop/asynchronous-programming-in-office-add-ins.md)
- [Просмотр или изменение темы при создании встречи или сообщения в Outlook](get-or-set-the-subject.md)
- [Вставка данных в текст при создании встречи или сообщения в Outlook](insert-data-in-the-body.md)
- [Просмотр или изменение расположения при создании встречи в Outlook](get-or-set-the-location-of-an-appointment.md)
- [Просмотр или изменение времени при создании встречи в Outlook](get-or-set-the-time-of-an-appointment.md)
    
