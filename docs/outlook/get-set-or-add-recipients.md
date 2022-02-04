---
title: Просмотр и изменение получателей в надстройке Outlook
description: 'Узнайте, как просмотреть, изменить или добавить получателей сообщения или встречи в надстройке Outlook.'
ms.date: 10/15/2021
ms.localizationpriority: medium
---

# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook


API Office JavaScript предоставляет асинхронные методы ([Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1)) или [Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))) для получения, набора или добавления получателей в форме записи встречи или сообщения. Эти асинхронные методы доступны только для составить надстройки. Чтобы использовать эти методы, убедитесь, что для Outlook надстройки необходимо настроить манифест надстройки, как описано в create [Outlook](compose-scenario.md) надстройки для создания форм.

Некоторые свойства, представляющие получателей в сообщении или встрече, доступны для чтения в формах создания и чтения. Это свойства [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) для встреч, а также свойства [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) для сообщений.

В форме чтения доступ к свойству можно получить напрямую из родительского объекта, например:

```js
item.cc
```

Но в форме составить, так как пользователь и ваша надстройка могут одновременно вставлять или изменять получателя, `getAsync` для получения этих свойств необходимо использовать асинхронный метод, как в следующем примере.


```js
item.cc.getAsync
```

Эти свойства доступны для записи только в формах создания, но не формах чтения.

Как и большинство асинхронных методов в API JavaScript для Office, `getAsync`и `setAsync``addAsync` принимать необязательные параметры ввода. Дополнительные сведения об указании последних см. в разделе [Передача дополнительных параметров в асинхронные методы](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) статьи [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-recipients"></a>Извлечение получателей


В этом разделе показан пример кода, извлекающий получателей создаваемой встречи или сообщения, и показывающий адреса электронной почты получателей. В примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

В API Office JavaScript, так как свойства, которые представляют получателей встречи (**необязательныйAttendees** и **requiredAttendees**), отличаются от свойств сообщения ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **cc** и **to**), сначала следует использовать свойство [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), чтобы определить, является ли составный элемент назначением или сообщением. В режиме композитации все эти свойства назначений и сообщений [](/javascript/api/outlook/office.recipients) являются объектами получателей, поэтому можно применить асинхронный метод, `Recipients.getAsync`чтобы получить соответствующих получателей.

Чтобы использовать `getAsync` метод обратного вызова для проверки состояния, результатов и любой ошибки, возвращаемой асинхронным вызовом `getAsync` . Методу обратного вызова можно передать любые аргументы, используя дополнительный параметр  _asyncContext_. Метод обратного вызова возвращает выходной параметр  _asyncResult_. Вы можете `status` `error` использовать свойства объекта параметра [AsyncResult](/javascript/api/office/office.asyncresult) для проверки состояния и любых сообщений об ошибках асинхронного вызова, `value` а также свойства для получения фактических получателей. Они представляются как массив объектов [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails).

`getAsync` Обратите внимание, что поскольку метод асинхронный, если существуют последующие действия, зависят от успешного получения получателей, необходимо организовать код для запуска таких действий только в соответствующем методе вызова после успешного завершения асинхронного вызова.

> [!IMPORTANT]
> В Outlook в Интернете случае, если пользователь создал новое сообщение, активировав ссылку на электронный адрес контакта или карточку профиля, `Recipients.getAsync` `displayName` вызов надстройки в настоящее время не возвращает значение в свойстве связанного `EmailAddressDetails` объекта.
> Дополнительные сведения обратитесь к [связанной GitHub проблеме](https://github.com/OfficeDev/office-js-docs-pr/issues/2962).

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


В этом разделе показан пример кода, устанавливающий получателей встречи или сообщения, создаваемых пользователем. При установке получателей все существующие получателей перезаписываются. Как и предыдущий пример, извлекающий получателей в форме создания, этот пример предполагает, что надстройка активируется в формах создания встреч и сообщений. В этом примере сначала проверяется, является ли составленный элемент назначением или сообщением, `Recipients.setAsync`поэтому для применения асинхронного метода применяются соответствующие свойства, которые представляют получателей встречи или сообщения.

При вызове `setAsync`устрой массив в качестве аргумента ввода для параметра  _получателей_ в одном из следующих форматов.


- массив строк, являющихся SMTP-адресами;
    
- массив словарей, каждый из которых содержит отображаемое имя и адрес электронной почты, как показано в следующем примере кода;
    
- Массив объектов `EmailAddressDetails` , аналогичных возвращаемой методом `getAsync` .
    
Вы можете дополнительно `setAsync` предоставить метод вызова в качестве аргумента ввода для метода, чтобы убедиться, что любой код, который зависит от успешного настройки получателей будет выполняться только тогда, когда это произойдет. Кроме того, методу обратного вызова можно передать любые аргументы, используя дополнительный параметр  _asyncContext_. При использовании метода вызова можно получить доступ к выходному параметру _asyncResult_ и  `AsyncResult` использовать свойства состояния и ошибки объекта параметра для проверки состояния и сообщений об ошибках асинхронного вызова.




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

Если вы не хотите перезаписывать существующие получатели в записи или сообщении, `Recipients.setAsync`вместо использования, `Recipients.addAsync` вы можете использовать асинхронный метод для приложения получателей. `addAsync` работает так же, как `setAsync` и в том, что для этого требуется аргумент _ввода_ получателей. Можно дополнительно предоставить метод вызова и любые аргументы для вызова с помощью параметра asyncContext. Затем можно проверить состояние, `addAsync` результат и любую ошибку асинхронного вызова с помощью параметра _вывода asyncResult_ метода вызова. В следующем примере проверяется, является ли составленный элемент назначением, и приложены два необходимых участника к встрече.


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
    
