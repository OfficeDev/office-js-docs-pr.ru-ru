---
title: Добавление и удаление вложений в надстройке Outlook
description: Вы можете использовать различные API вложений для управления файлами Outlook элементов, присоединенных к элементу, который создает пользователь.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 87076965d600cbbcfe88d6711ea3acfb2b3c1fdd
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/05/2021
ms.locfileid: "53774457"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Управление вложениями элемента в форме композиции в Outlook

API Office JavaScript предоставляет несколько API, которые можно использовать для управления вложениями элемента при его сочинении.

## <a name="attach-a-file-or-outlook-item"></a>Прикрепить файл или Outlook элемент

Вы можете прикрепить файл или Outlook элемент к форме композиции с помощью метода, подходящего для типа вложения.

- [addFileAttachmentAsync:](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)Прикрепить файл
- [addFileAttachmentFromBase64Async:](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)Прикрепить файл с помощью строки base64
- [addItemAttachmentAsync:](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)прикрепите элемент Outlook

Это асинхронные методы, что означает, что выполнение можно выполнить, не дожидаясь завершения действия. В зависимости от исходного расположения и размера добавляемого вложения асинхронный вызов может занять некоторое время.

Если какие-то задачи зависят от завершения действия, их следует выполнять в методе обратного вызова. Этот метод вызова необязателен и вызывается после завершения загрузки вложения. Метод обратного вызова принимает объект [AsyncResult](/javascript/api/office/office.asyncresult) как выходной параметр, который содержит состояние, ошибку и возвращаемое значение при добавлении вложения. Если для обратного вызова требуются дополнительные параметры, их можно указать в необязательном параметре `options.asyncContext`. Параметр `options.asyncContext` может относиться к любому типу, поддерживаемому методом обратного вызова.

Например, можно определить объект JSON, содержащий одну или несколько пар значений `options.asyncContext` ключа. Дополнительные примеры передачи необязательных параметров асинхронным методам можно найти на платформе Office надстройки в Асинхронном программировании в [Office](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-to-asynchronous-methods)надстройки . В следующем примере показано, как использовать параметр для `asyncContext` передачи 2 аргументов методу вызова.

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

Успешность обратного вызова асинхронного метода можно проверить с помощью свойств `status` и `error` объекта `AsyncResult`. Если операция вложения завершается успешно, вы можете использовать свойство `AsyncResult.value`, чтобы получить идентификатор вложения. Это целое число, которое можно использовать в дальнейшем, чтобы удалить вложение.

> [!NOTE]
> ID вложения действителен только в пределах одного сеанса и не гарантируется для привязки к одному и том же вложению во всех сеансах. Примеры того, когда сеанс более, включают, когда пользователь закрывает надстройки, или если пользователь начинает сочинять в форме, а затем выскакивая в линию форму, чтобы продолжить в отдельном окне.

### <a name="attach-a-file"></a>Присоединение файла

Вы можете прикрепить файл к сообщению или встрече в форме записи с помощью метода и указания `addFileAttachmentAsync` URI файла. Вы также можете использовать `addFileAttachmentFromBase64Async` метод, но укажите строку base64 в качестве ввода. Если файл защищен, можно добавить соответствующее удостоверение или токен проверки подлинности как параметр строки запроса URI. Exchange вызовет URI, чтобы получить вложение, а веб-службе, которая защищает файл, потребуется использовать токен для проверки подлинности.

Следующий пример JavaScript — это надстройка создания, которая прикрепляет файл picture.png с веб-сервера к создаваемому сообщению или встрече. Метод обратного вызова принимает `asyncResult` в качестве параметра, проверяет состояние результата и получает его идентификатор, если метод выполнен успешно.

```js
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID.
        // You can optionally pass any object that you would
        // access in the callback method as an argument to
        // the asyncContext parameter.
        Office.context.mailbox.item.addFileAttachmentAsync(
            `https://webserver/picture.png`,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                } else {
                    // Get the ID of the attached file.
                    var attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="attach-an-outlook-item"></a>Прикрепить элемент Outlook

Вы можете прикрепить элемент Outlook (например, электронную почту, календарь или контактный элемент) к сообщению или встрече в форме записи, указав Exchange веб-службы (EWS) ИД элемента и используя `addItemAttachmentAsync` метод. Вы можете получить EWS-ID элемента электронной почты, календаря, контакта или задачи в почтовом ящике пользователя с помощью метода [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) и доступа к операции EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). Свойство [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) также предоставляет идентификатор EWS существующего элемента в форме чтения.

Следующая функция JavaScript расширяет первый пример выше и добавляет элемент в качестве вложения в составную электронную почту или `addItemAttachment` встречу. В качестве параметра функция принимает идентификатор EWS прикрепляемого элемента. Если присоединение успешно, он получает ID вложения для дальнейшей обработки, включая удаление этого вложения в том же сеансе.

```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            } else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> Надстройку можно использовать для прикрепления экземпляра повторяющейся встречи в Outlook в Интернете или на мобильных устройствах. Однако в клиенте Outlook настольного компьютера попытка прикрепить экземпляр приведет к присоединению повторяющейся серии (родительского назначения).

## <a name="get-attachments"></a>Получение вложений

API для получения вложений в режиме композитации доступны из набора требований [1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).

- [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

Вы можете использовать [метод getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) для получения вложений сообщения или записи на прием.

Чтобы получить содержимое вложения, можно использовать [метод getAttachmentContentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) Поддерживаемые форматы перечислены в перечислении [AttachmentContentFormat.](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

Необходимо предоставить метод вызова, чтобы проверить состояние и любую ошибку с помощью объекта `AsyncResult` параметра вывода. Вы также можете передать все дополнительные параметры методу вызова, используя необязательный `asyncContext` параметр.

В следующем примере JavaScript вы можете получить вложения и настроить отдельную обработку для каждого поддерживаемого формата вложений.

```js
var item = Office.context.mailbox.item;
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

## <a name="remove-an-attachment"></a>Удаление вложения

Вы можете удалить вложение файла или элемента из сообщения или элемента встречи в форме записи, указав соответствующий ID вложения при использовании метода [removeAttachmentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

> [!IMPORTANT]
> Если используется набор требований 1.7 или ранее, следует удалять только вложения, добавленные той же надстройкой в том же сеансе.

Аналогично методу `addFileAttachmentAsync` `addItemAttachmentAsync` , и `getAttachmentsAsync` методам, является `removeAttachmentAsync` асинхронным методом. Необходимо предоставить метод вызова, чтобы проверить состояние и любую ошибку с помощью объекта `AsyncResult` параметра вывода. Вы также можете передать все дополнительные параметры методу вызова, используя необязательный `asyncContext` параметр.

Следующая функция JavaScript продолжает расширять указанные выше примеры и удаляет указанное вложение из записи электронной почты или записи. `removeAttachment` В качестве аргумента функция принимает идентификатор вложения, которое требуется удалить. Вы можете получить ID вложения после успешного или методного вызова и использовать его в `addFileAttachmentAsync` `addFileAttachmentFromBase64Async` `addItemAttachmentAsync` последующем `removeAttachmentAsync` вызове метода. Вы также можете вызвать `getAttachmentsAsync` (вводится в наборе требований 1.8), чтобы получить вложения и их ID для этого сеанса надстройки.

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback method is invoked.
    // Here, the callback method uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback method as an argument to the asyncContext parameter.
    Office.context.mailbox.item.removeAttachmentAsync(
        attachmentId,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            } else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```

## <a name="see-also"></a>См. также

- [Создание надстроек Outlook для форм создания](compose-scenario.md)
- [Асинхронное программирование надстроек Office](../develop/asynchronous-programming-in-office-add-ins.md)
