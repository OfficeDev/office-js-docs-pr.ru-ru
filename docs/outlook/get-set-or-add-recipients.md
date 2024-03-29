---
title: Просмотр и изменение получателей в надстройке Outlook
description: Узнайте, как просмотреть, изменить или добавить получателей сообщения или встречи в надстройке Outlook.
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: de47d2ee238ffe55ab0b5ee460096717557e4dba
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541270"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook

API JavaScript для Office предоставляет асинхронные методы ([Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1)) или [Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))) для получения, задания или добавления получателей в форме создания встречи или сообщения. Эти асинхронные методы доступны только для создания надстроек. Чтобы использовать эти методы, убедитесь, что манифест надстройки настроен соответствующим образом для Outlook, чтобы активировать надстройку в формах создания, как описано в разделе "Создание надстроек [Outlook](compose-scenario.md) для форм создания". Правила активации не поддерживаются в надстройки, использующие манифест [Teams для надстроек Office (предварительная версия).](../develop/json-manifest-overview.md)

Some of the properties that represent recipients in an appointment or message are available for read access in a compose form and in a read form. These properties include  [optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) for appointments, and [cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), and  [to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) for messages.

В форме чтения доступ к свойству можно получить напрямую из родительского объекта, например:

```js
item.cc
```

Но в форме создания, так как пользователь и ваша надстройка могут вставлять или изменять получателя одновременно, `getAsync` для получения этих свойств необходимо использовать асинхронный метод, как показано в следующем примере.

```js
item.cc.getAsync
```

Эти свойства доступны для записи только в формах создания, но не формах чтения.

Как и большинство асинхронных методов в API JavaScript для Office, `getAsync``setAsync``addAsync` и примите необязательные входные параметры. Дополнительные сведения об указании последних см. в разделе [Передача дополнительных параметров в асинхронные методы](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) статьи [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).

## <a name="get-recipients"></a>Извлечение получателей

This section shows a code sample that gets the recipients of the appointment or message that is being composed, and displays the email addresses of the recipients. The code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

В API JavaScript для Office свойства, которые представляют получателей встречи ( **optionalAttendees** и **requiredAttendees**), отличаются от свойств сообщения ([bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), **cc** и **to**), сначала следует использовать свойство [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) , чтобы определить, является ли создаваемый элемент встречей или сообщением. В режиме создания все эти свойства встреч и сообщений являются объектами [Recipients](/javascript/api/outlook/office.recipients) , поэтому вы можете применить асинхронный метод, `Recipients.getAsync`чтобы получить соответствующих получателей.

Для использования `getAsync`предоставьте функцию обратного вызова для проверки состояния, результатов и любой ошибки, возвращаемой асинхронным вызовом `getAsync` . Вы можете указать любые аргументы функции обратного вызова, используя _необязательный параметр asyncContext_ . Функция обратного вызова возвращает _выходной параметр asyncResult_ . Вы можете `status` `error` использовать свойства объекта параметра [AsyncResult](/javascript/api/office/office.asyncresult) для проверки состояния и любых сообщений об ошибках асинхронного вызова, `value` а также свойство для получения фактических получателей. Они представляются как массив объектов [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails).

`getAsync` Обратите внимание, что, так как метод является асинхронным, при наличии последующих действий, которые зависят от успешного получения получателей, следует упорядочить код для запуска таких действий только в соответствующей функции обратного вызова после успешного завершения асинхронного вызова.

> [!IMPORTANT]
> Метод `getAsync` возвращает только получателей, разрешенных клиентом Outlook. Разрешенный получатель имеет следующие характеристики.
>
> - Если у получателя есть сохраненная запись в адресной книге отправителя, Outlook разрешает адрес электронной почты в сохраненное отображаемое имя получателя.
> - Значок состояния собрания Teams отображается перед именем или адресом электронной почты получателя.
> - После имени или адреса электронной почты получателя появляется точка с запятой.
> - Имя или адрес электронной почты получателя подчеркнуты или заключены в поле.
>
> Чтобы разрешить адрес электронной почты после его добавления в почтовый элемент, отправитель должен использовать клавишу **TAB** или выбрать предлагаемый контакт или адрес электронной почты из списка автозавершание.

> [!NOTE]
> В Outlook в Интернете и Windows, если пользователь создает новое сообщение путем активации ссылки на адрес электронной почты контакта из карточки контакта или профиля, `Recipients.getAsync` `displayName` `EmailAddressDetails` вызов надстройки возвращает адрес электронной почты контакта в свойстве связанного объекта, а не сохраненное имя контакта.
>
> Дополнительные сведения см. в [связанной проблеме gitHub](https://github.com/OfficeDev/office-js/issues/2201).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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
    let toRecipients, ccRecipients, bccRecipients;
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
    for (let i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="set-recipients"></a>Установка получателей

В этом разделе показан пример кода, устанавливающий получателей встречи или сообщения, создаваемых пользователем. При установке получателей все существующие получателей перезаписываются. Как и предыдущий пример, извлекающий получателей в форме создания, этот пример предполагает, что надстройка активируется в формах создания встреч и сообщений. В этом примере сначала проверяется, является ли составной элемент встречей или сообщением, поэтому для применения асинхронного метода к соответствующим свойствам, `Recipients.setAsync`которые представляют получателей встречи или сообщения.

При вызове `setAsync`укажите массив в качестве входного аргумента для параметра  _получателя_ в одном из следующих форматов.

- массив строк, являющихся SMTP-адресами;
- массив словарей, каждый из которых содержит отображаемое имя и адрес электронной почты, как показано в следующем примере кода;
- Массив объектов `EmailAddressDetails` , аналогичный возвращаемой методом `getAsync` .
  
При необходимости можно `setAsync` указать функцию обратного вызова в качестве входного аргумента для метода, чтобы убедиться, что любой код, который зависит от успешной настройки получателей, будет выполняться только в этом случае. Можно также указать любые аргументы для функции обратного вызова, используя _необязательный параметр asyncContext_ . При использовании функции обратного вызова можно получить доступ к выходному параметру _asyncResult_ и  `AsyncResult` использовать свойства  состояния и ошибки объекта параметра для проверки состояния и сообщений об ошибках асинхронного вызова.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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
    let toRecipients, ccRecipients, bccRecipients;

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

Если вы не хотите перезаписывать существующих получателей во встрече или сообщении, `Recipients.setAsync``Recipients.addAsync` вместо использования можно использовать асинхронный метод для добавления получателей. `addAsync` работает так же, как `setAsync` и в том, что для него требуется _входной аргумент_ получателя. При необходимости можно указать функцию обратного вызова и любые аргументы для обратного вызова с помощью параметра asyncContext. Затем можно проверить состояние, результат `addAsync` и любую ошибку асинхронного вызова с помощью выходного параметра _asyncResult_ функции обратного вызова. В следующем примере проверяется, является ли составленный элемент встречей, и добавляет двух обязательных участников к встрече.

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
- [Асинхронное программирование в случае надстроек Office](../develop/asynchronous-programming-in-office-add-ins.md)
- [Считывание и запись темы при создании встречи или сообщения в Outlook](get-or-set-the-subject.md)
- [Вставка данных в текст при создании встречи или сообщения в Outlook](insert-data-in-the-body.md)
- [Просмотр или изменение расположения при создании встречи в Outlook](get-or-set-the-location-of-an-appointment.md)
- [Просмотр или изменение времени при создании встречи в Outlook](get-or-set-the-time-of-an-appointment.md)
