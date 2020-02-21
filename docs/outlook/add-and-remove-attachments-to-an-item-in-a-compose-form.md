---
title: Добавление и удаление вложений в надстройке Outlook
description: Можно использовать различные API вложений для управления файлами или элементами Outlook, связанными с элементом, создаваемым пользователем.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 977b8fa814a251c76aabc64345762a3a9556a60b
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166789"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Управление вложениями элемента в форме создания в Outlook

API JavaScript для Office предоставляет несколько интерфейсов API, которые можно использовать для управления вложениями элементов при создании пользователем.

## <a name="attach-a-file-or-outlook-item"></a>Вложение файла или элемента Outlook

Вы можете прикрепить файл или элемент Outlook к форме создания, используя метод, который соответствует типу вложения.

- [addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Прикрепление файла
- [addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): присоединение файла с помощью строки Base64
- [addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): присоединение элемента Outlook

Это асинхронные методы, что означает, что выполнение может пройти без ожидания завершения действия. В зависимости от исходного расположения и размера добавляемого вложения для завершения асинхронного вызова может потребоваться некоторое время.

Если какие-то задачи зависят от завершения действия, их следует выполнять в методе обратного вызова. Этот метод обратного вызова является необязательным и вызывается после завершения отправки вложения. Метод обратного вызова принимает объект [AsyncResult](/javascript/api/office/office.asyncresult) как выходной параметр, который содержит состояние, ошибку и возвращаемое значение при добавлении вложения. Если для обратного вызова требуются дополнительные параметры, их можно указать в необязательном параметре `options.asyncContext`. Параметр `options.asyncContext` может относиться к любому типу, поддерживаемому методом обратного вызова.

Например, можно определить параметр `options.asyncContext` как объект JSON, содержащий одну или несколько пар "ключ-значение". Дополнительные примеры передачи необязательных параметров в асинхронные методы для платформы надстроек Office см. в статье [Асинхронное программирование в надстройках Office](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). Ниже показано, как использовать параметр `asyncContext` для передачи двух аргументов методу обратного вызова.

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

Успешность обратного вызова асинхронного метода можно проверить с помощью свойств `status` и `error` объекта `AsyncResult`. Если операция вложения завершается успешно, вы можете использовать свойство `AsyncResult.value`, чтобы получить идентификатор вложения. Это целое число, которое можно использовать в дальнейшем, чтобы удалить вложение.

> [!NOTE]
> Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено той же надстройкой в ходе текущего сеанса. В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает надстройку или начинает создавать элемент во встроенной форме и затем продолжает работу в отдельном окне.

### <a name="attach-a-file"></a>Вложение файла

Вы можете прикрепить файл к сообщению или встрече в форме создания, используя `addFileAttachmentAsync` метод и указав универсальный код ресурса (URI) файла. Кроме того, можно использовать `addFileAttachmentFromBase64Async` метод, но указать строку Base64 в качестве входных данных. Если файл защищен, можно добавить соответствующее удостоверение или токен проверки подлинности как параметр строки запроса URI. Exchange вызовет URI, чтобы получить вложение, а веб-службе, которая защищает файл, потребуется использовать токен для проверки подлинности.

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
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
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

### <a name="attach-an-outlook-item"></a>Вложение элемента Outlook

Вы можете прикрепить элемент Outlook (например, электронное сообщение, элемент календаря или контакт) к сообщению или встрече в форме создания, указав идентификатор элемента в веб-службах Exchange (EWS) и вызвав метод `addItemAttachmentAsync`. Вы можете получить идентификатор EWS для элемента сообщения, календаря, контакта или задачи в почтовом ящике пользователя, вызвав метод [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) и используя операцию EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). Свойство [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) также предоставляет идентификатор EWS существующего элемента в форме чтения.

Приведенная ниже функция `addItemAttachment`JavaScript, которая расширяет первый пример выше и добавляет элемент в качестве вложения в создаваемую электронную почту или встречу. В качестве параметра функция принимает идентификатор EWS прикрепляемого элемента. В случае успешного присоединения он получает идентификатор вложения для дальнейшей обработки, включая удаление этого вложения в том же сеансе.

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
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> Вы можете использовать надстройку создания, чтобы вложить экземпляр повторяющейся встречи в Outlook в Интернете или на мобильных устройствах. Однако в расширенном клиенте Outlook попытка прикрепить такой экземпляр приведет к прикреплению ряда повторений (основной встречи).

## <a name="get-attachments"></a>Получение вложений

Вы можете использовать метод [жетаттачментсасинк](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) для получения вложений создаваемого сообщения или встречи.

Чтобы получить содержимое вложения, можно использовать метод [жетаттачментконтентасинк](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) . Поддерживаемые форматы перечислены в перечислении [аттачментконтентформат](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) .

Необходимо предоставить метод обратного вызова для проверки состояния и любой ошибки с помощью объекта `AsyncResult` Output Parameter. Кроме того, можно передать дополнительные параметры в метод обратного вызова, используя необязательный `asyncContext` параметр.

В приведенном ниже примере JavaScript показано, как получить вложения и настроить индивидуальную обработку для каждого поддерживаемого формата вложений.

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

Вы можете удалить вложение из элемента сообщения или встречи в форме создания, указав соответствующий идентификатор вложения и вызвав метод [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods). Удалять можно только вложения, добавленные одной и той же надстройкой в одном сеансе. Аналогично `addFileAttachmentAsync` методам `addItemAttachmentAsync` и, `removeAttachmentAsync` — это асинхронный метод. Необходимо предоставить метод обратного вызова для проверки состояния и любой ошибки с помощью объекта `AsyncResult` Output Parameter. Кроме того, можно передать дополнительные параметры в метод обратного вызова, используя необязательный `asyncContext` параметр.

Приведенная ниже функция `removeAttachment`JavaScript продолжает расширять приведенные выше примеры и удаляет указанное вложение из создаваемого сообщения электронной почты или встречи. В качестве аргумента функция принимает идентификатор вложения, которое требуется удалить. Вы можете получить идентификатор `addFileAttachmentAsync`вложения после успешного `addFileAttachmentFromBase64Async`вызова, или `addItemAttachmentAsync` вызова метода и сохранить его для последующего `removeAttachmentAsync` вызова метода.

```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be
// removed.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
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
