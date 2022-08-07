---
title: Добавление и удаление вложений в надстройке Outlook
description: Используйте различные API вложений для управления файлами или элементами Outlook, прикрепленными к элементу, который создает пользователь.
ms.date: 08/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: af3b44814fd11c5e2006dbb921130c15c7535385
ms.sourcegitcommit: 76b8c79cba707c771ae25df57df14b6445f9b8fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2022
ms.locfileid: "67274171"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Управление вложениями элемента в форме создания в Outlook

API JavaScript для Office предоставляет несколько API, которые можно использовать для управления вложениями элемента при создании пользователем.

## <a name="attach-a-file-or-outlook-item"></a>Вложение файла или элемента Outlook

Вы можете вложить файл или элемент Outlook в форму создания с помощью метода, подходящего для типа вложения.

- [addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): вложение файла
- [addFileAttachmentFromBase64Async](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): присоединение файла с помощью строки base64
- [addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): присоединение элемента Outlook

Это асинхронные методы, то есть выполнение может выполняться без ожидания завершения действия. В зависимости от исходного расположения и размера добавляемого вложения выполнение асинхронного вызова может занять некоторое время.

Если существуют задачи, которые зависят от выполняемого действия, эти задачи следует выполнять в функции обратного вызова. Эта функция обратного вызова является необязательной и вызывается после завершения отправки вложения. Функция обратного вызова принимает объект [AsyncResult](/javascript/api/office/office.asyncresult) в качестве выходного параметра, который предоставляет любое состояние, ошибку и возвращаемое значение при добавлении вложения. Если для обратного вызова требуются дополнительные параметры, их можно указать в необязательном параметре `options.asyncContext`. `options.asyncContext` может иметь любой тип, ожидаемый функцией обратного вызова.

Например, можно определить как `options.asyncContext` объект JSON, содержащий одну или несколько пар "ключ-значение". Дополнительные примеры передачи необязательных параметров в асинхронные методы можно найти на платформе надстроек Office в асинхронном программировании в надстройки [Office](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-to-asynchronous-methods). В следующем примере показано, как использовать параметр `asyncContext` для передачи 2 аргументов функции обратного вызова.

```js
const options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

Вы можете проверить успешность или `status` `error` ошибку асинхронного вызова метода в функции обратного вызова, используя свойства объекта `AsyncResult` . Если присоединение завершается успешно, можно `AsyncResult.value` использовать свойство для получения идентификатора вложения. Это целое число, которое можно использовать в дальнейшем, чтобы удалить вложение.

> [!NOTE]
> Идентификатор вложения действителен только в пределах одного сеанса и не гарантируется сопоставление одного и того же вложения между сеансами. Примеры завершения сеанса включают, когда пользователь закрывает надстройку или начинает составление во встроенной форме, а затем выводит встроенную форму, чтобы продолжить в отдельном окне.

### <a name="attach-a-file"></a>Вложение файла

Вы можете вложить файл `addFileAttachmentAsync` в сообщение или встречу в форме создания с помощью метода и указать универсальный код ресурса (URI) файла. Этот метод также можно использовать `addFileAttachmentFromBase64Async` , но в качестве входных данных указать строку base64. Если файл защищен, можно добавить соответствующее удостоверение или токен проверки подлинности как параметр строки запроса URI. Exchange вызовет URI, чтобы получить вложение, а веб-службе, которая защищает файл, потребуется использовать токен для проверки подлинности.

Следующий пример JavaScript — это надстройка создания, которая прикрепляет файл picture.png с веб-сервера к создаваемому сообщению или встрече. Функция обратного вызова принимает `asyncResult` в качестве параметра, проверяет состояние результата и получает идентификатор вложения, если метод выполнен успешно.

```js
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback function is invoked and gets the attachment ID.
        // You can optionally pass any object that you would
        // access in the callback function as an argument to
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
                    const attachmentID = asyncResult.value;
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

Чтобы добавить встроенное изображение base64 в текст создаваемого сообщения или встречи, `Office.context.mailbox.item.body.getAsync` `addFileAttachmentFromBase64Async` необходимо сначала получить текущий текст элемента с помощью метода перед вставкой изображения с помощью метода. В противном случае изображение не будет отображаться в теле после вставки. Инструкции см. в следующем примере JavaScript, который добавляет встроенный образ base64 в начало текста элемента.

```js
const mailItem = Office.context.mailbox.item;
const base64String =
  "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAnUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAN0S+bUAAAAMdFJOUwAQIDBAUI+fr7/P7yEupu8AAAAJcEhZcwAADsMAAA7DAcdvqGQAAAF8SURBVGhD7dfLdoMwDEVR6Cspzf9/b20QYOthS5Zn0Z2kVdY6O2WULrFYLBaLxd5ur4mDZD14b8ogWS/dtxV+dmx9ysA2QUj9TQRWv5D7HyKwuIW9n0vc8tkpHP0W4BOg3wQ8wtlvA+PC1e8Ao8Ld7wFjQtHvAiNC2e8DdqHqKwCrUPc1gE1AfRVgEXBfB+gF0lcCWoH2tYBOYPpqQCNwfT3QF9i+AegJfN8CtAWhbwJagtS3AbIg9o2AJMh9M5C+SVGBvx6zAfmT0r+Bv8JMwP4kyFPir+cswF5KL3WLv14zAFBCLf56Tw9cparFX4upgaJUtPhrOS1QlY5W+vWTXrGgBFB/b72ev3/0igUdQPppP/nfowfKUUEFcP207y/yxKmgAYQ+PywoAFOfCH3A2MdCFzD3kdADBvq10AGG+pXQBgb7pdAEhvuF0AIc/VtoAK7+JciAs38KIuDugyAC/v4hiMCE/i7IwLRBsh68N2WQjMVisVgs9i5bln8LGScNcCrONQAAAABJRU5ErkJggg==";

// Get the current body of the message or appointment.
mailItem.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
  if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
    // Insert the base64 image to the beginning of the body.
    const options = { isInline: true, asyncContext: bodyResult.value };
    mailItem.addFileAttachmentFromBase64Async(base64String, "sample.png", options, (attachResult) => {
      if (attachResult.status === Office.AsyncResultStatus.Succeeded) {
        let body = attachResult.asyncContext;
        body = body.replace("<p class=MsoNormal>", `<p class=MsoNormal><img src="cid:sample.png">`);
        mailItem.body.setAsync(body, { coercionType: Office.CoercionType.Html }, (setResult) => {
          if (setResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Inline base64 image added to the body.");
          } else {
            console.log(setResult.error.message);
          }
        });
      } else {
        console.log(attachResult.error.message);
      }
    });
  } else {
    console.log(bodyResult.error.message);
  }
});
```

### <a name="attach-an-outlook-item"></a>Присоединение элемента Outlook

Вы можете вложить элемент Outlook (например, электронную почту, календарь или элемент контакта) в сообщение или встречу в форме создания, указав идентификатор веб-служб Exchange (EWS) `addItemAttachmentAsync` элемента и используя метод. Вы можете получить идентификатор EWS сообщения электронной почты, календаря, контакта или элемента задачи в почтовом ящике пользователя с помощью метода [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) и доступа к операции EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). Свойство [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) также предоставляет идентификатор EWS существующего элемента в форме чтения.

Приведенная ниже функция `addItemAttachment`JavaScript расширяет первый приведенный выше пример и добавляет элемент в качестве вложения в создаваемую электронную почту или встречу. В качестве параметра функция принимает идентификатор EWS прикрепляемого элемента. Если подключение выполнено успешно, он получает идентификатор вложения для дальнейшей обработки, включая удаление этого вложения в том же сеансе.

```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback function is invoked. Here, the callback
    // function uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback function as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            } else {
                const attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> Надстройку compose можно использовать для подключения экземпляра повторяющейся встречи в Outlook в Интернете или на мобильных устройствах. Однако в классическом клиенте Outlook попытка подключения экземпляра приведет к присоединению повторяющегося ряда (родительской встречи).

## <a name="get-attachments"></a>Получение вложений

API для получения вложений в режиме создания доступны из набора обязательных [элементов 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8).

- [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

Метод [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) можно использовать для получения вложений сообщения или создаваемой встречи.

Чтобы получить содержимое вложения, можно использовать метод [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) . Поддерживаемые форматы перечислены в [перечислении AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) .

Необходимо предоставить функцию обратного вызова для проверки состояния и любой ошибки с помощью объекта выходного `AsyncResult` параметра. Вы также можете передать любые дополнительные параметры функции обратного вызова с помощью необязательного параметра `asyncContext` .

В следующем примере Кода JavaScript показано, как получить вложения и настроить различные обработки для каждого поддерживаемого формата вложений.

```js
const item = Office.context.mailbox.item;
const options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (let i = 0 ; i < result.value.length ; i++) {
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

Вы можете удалить вложение файла или элемента из сообщения или элемента встречи в форме создания, указав соответствующий идентификатор вложения при использовании метода [removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) .

> [!IMPORTANT]
> Если вы используете набор обязательных элементов 1.7 или более ранней версии, следует удалять только вложения, добавленные той же надстройке в том же сеансе.

Аналогично методу `addFileAttachmentAsync`, `addItemAttachmentAsync`и `getAttachmentsAsync` методам, `removeAttachmentAsync` является асинхронным методом. Необходимо предоставить функцию обратного вызова для проверки состояния и любой ошибки с помощью объекта выходного `AsyncResult` параметра. Вы также можете передать любые дополнительные параметры функции обратного вызова с помощью необязательного параметра `asyncContext` .

Приведенная ниже функция JavaScript `removeAttachment`продолжает расширять приведенные выше примеры и удаляет указанное вложение из создаваемого сообщения электронной почты или встречи. В качестве аргумента функция принимает идентификатор вложения, которое требуется удалить. Идентификатор вложения можно получить после `addFileAttachmentAsync``addFileAttachmentFromBase64Async``addItemAttachmentAsync` успешного вызова метода или метода и использовать его в последующем вызове `removeAttachmentAsync` метода. Вы также можете вызвать `getAttachmentsAsync` (представленный в наборе обязательных элементов 1.8), чтобы получить вложения и их идентификаторы для этого сеанса надстройки.

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback function is invoked.
    // Here, the callback function uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback function as an argument to the asyncContext parameter.
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
