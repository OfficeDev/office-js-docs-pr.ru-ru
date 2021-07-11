---
title: Добавление и удаление вложений в надстройке Outlook
description: Вы можете использовать различные API вложений для управления файлами Outlook элементов, присоединенных к элементу, который создает пользователь.
ms.date: 02/24/2021
localization_priority: Normal
ms.openlocfilehash: 0ba142bb1e8fb5f324d2bb6460bc8325a4800d2d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348589"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a><span data-ttu-id="d056b-103">Управление вложениями элемента в форме композиции в Outlook</span><span class="sxs-lookup"><span data-stu-id="d056b-103">Manage an item's attachments in a compose form in Outlook</span></span>

<span data-ttu-id="d056b-104">API Office JavaScript предоставляет несколько API, которые можно использовать для управления вложениями элемента при его сочинении.</span><span class="sxs-lookup"><span data-stu-id="d056b-104">The Office JavaScript API provides several APIs you can use to manage an item's attachments when the user is composing.</span></span>

## <a name="attach-a-file-or-outlook-item"></a><span data-ttu-id="d056b-105">Прикрепить файл или Outlook элемент</span><span class="sxs-lookup"><span data-stu-id="d056b-105">Attach a file or Outlook item</span></span>

<span data-ttu-id="d056b-106">Вы можете прикрепить файл или Outlook элемент к форме композиции с помощью метода, подходящего для типа вложения.</span><span class="sxs-lookup"><span data-stu-id="d056b-106">You can attach a file or Outlook item to a compose form by using the method that's appropriate for the type of attachment.</span></span>

- <span data-ttu-id="d056b-107">[addFileAttachmentAsync:](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)Прикрепить файл</span><span class="sxs-lookup"><span data-stu-id="d056b-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file</span></span>
- <span data-ttu-id="d056b-108">[addFileAttachmentFromBase64Async:](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)Прикрепить файл с помощью строки base64</span><span class="sxs-lookup"><span data-stu-id="d056b-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file using its base64 string</span></span>
- <span data-ttu-id="d056b-109">[addItemAttachmentAsync:](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)прикрепите элемент Outlook</span><span class="sxs-lookup"><span data-stu-id="d056b-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach an Outlook item</span></span>

<span data-ttu-id="d056b-110">Это асинхронные методы, что означает, что выполнение можно выполнить, не дожидаясь завершения действия.</span><span class="sxs-lookup"><span data-stu-id="d056b-110">These are asynchronous methods, which means execution can go on without waiting for the action to complete.</span></span> <span data-ttu-id="d056b-111">В зависимости от исходного расположения и размера добавляемого вложения асинхронный вызов может занять некоторое время.</span><span class="sxs-lookup"><span data-stu-id="d056b-111">Depending on the original location and size of the attachment being added, the asynchronous call may take a while to complete.</span></span>

<span data-ttu-id="d056b-112">Если какие-то задачи зависят от завершения действия, их следует выполнять в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d056b-112">If there are tasks that depend on the action to complete, you should carry out those tasks in a callback method.</span></span> <span data-ttu-id="d056b-113">Этот метод вызова необязателен и вызывается после завершения загрузки вложения.</span><span class="sxs-lookup"><span data-stu-id="d056b-113">This callback method is optional and is invoked when the attachment upload has completed.</span></span> <span data-ttu-id="d056b-114">Метод обратного вызова принимает объект [AsyncResult](/javascript/api/office/office.asyncresult) как выходной параметр, который содержит состояние, ошибку и возвращаемое значение при добавлении вложения.</span><span class="sxs-lookup"><span data-stu-id="d056b-114">The callback method takes an [AsyncResult](/javascript/api/office/office.asyncresult) object as an output parameter that provides any status, error, and returned value from adding the attachment.</span></span> <span data-ttu-id="d056b-115">Если для обратного вызова требуются дополнительные параметры, их можно указать в необязательном параметре `options.asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="d056b-115">If the callback requires any extra parameters, you can specify them in the optional `options.asyncContext` parameter.</span></span> <span data-ttu-id="d056b-116">Параметр `options.asyncContext` может относиться к любому типу, поддерживаемому методом обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d056b-116">`options.asyncContext` can be of any type that your callback method expects.</span></span>

<span data-ttu-id="d056b-117">Например, можно определить объект JSON, содержащий одну или несколько пар значений `options.asyncContext` ключа.</span><span class="sxs-lookup"><span data-stu-id="d056b-117">For example, you can define `options.asyncContext` as a JSON object that contains one or more key-value pairs.</span></span> <span data-ttu-id="d056b-118">Дополнительные примеры передачи необязательных параметров асинхронным методам можно найти на платформе Office надстройки в Асинхронном программировании в [Office](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)надстройки . В следующем примере показано, как использовать параметр для `asyncContext` передачи 2 аргументов методу вызова.</span><span class="sxs-lookup"><span data-stu-id="d056b-118">You can find more examples about passing optional parameters to asynchronous methods in the Office Add-ins platform in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). The following example shows how to use the `asyncContext` parameter to pass 2 arguments to a callback method.</span></span>

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

<span data-ttu-id="d056b-p104">Успешность обратного вызова асинхронного метода можно проверить с помощью свойств `status` и `error` объекта `AsyncResult`. Если операция вложения завершается успешно, вы можете использовать свойство `AsyncResult.value`, чтобы получить идентификатор вложения. Это целое число, которое можно использовать в дальнейшем, чтобы удалить вложение.</span><span class="sxs-lookup"><span data-stu-id="d056b-p104">You can check for success or error of an asynchronous method call in the callback method using the `status` and `error` properties of the `AsyncResult` object. If the attaching completes successfully, you can use the `AsyncResult.value` property to get the attachment ID. The attachment ID is an integer which you can subsequently use to remove the attachment.</span></span>

> [!NOTE]
> <span data-ttu-id="d056b-122">ID вложения действителен только в пределах одного сеанса и не гарантируется для привязки к одному и том же вложению во всех сеансах.</span><span class="sxs-lookup"><span data-stu-id="d056b-122">The attachment ID is valid only within the same session and isn't guaranteed to map to the same attachment across sessions.</span></span> <span data-ttu-id="d056b-123">Примеры того, когда сеанс более, включают, когда пользователь закрывает надстройки, или если пользователь начинает сочинять в форме, а затем выскакивая в линию форму, чтобы продолжить в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="d056b-123">Examples of when a session is over include when the user closes the add-in, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

### <a name="attach-a-file"></a><span data-ttu-id="d056b-124">Присоединение файла</span><span class="sxs-lookup"><span data-stu-id="d056b-124">Attach a file</span></span>

<span data-ttu-id="d056b-125">Вы можете прикрепить файл к сообщению или встрече в форме записи с помощью метода и указания `addFileAttachmentAsync` URI файла.</span><span class="sxs-lookup"><span data-stu-id="d056b-125">You can attach a file to a message or appointment in a compose form by using the `addFileAttachmentAsync` method and specifying the URI of the file.</span></span> <span data-ttu-id="d056b-126">Вы также можете использовать `addFileAttachmentFromBase64Async` метод, но укажите строку base64 в качестве ввода.</span><span class="sxs-lookup"><span data-stu-id="d056b-126">You can also use the `addFileAttachmentFromBase64Async` method but specify the base64 string as input.</span></span> <span data-ttu-id="d056b-127">Если файл защищен, можно добавить соответствующее удостоверение или токен проверки подлинности как параметр строки запроса URI.</span><span class="sxs-lookup"><span data-stu-id="d056b-127">If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter.</span></span> <span data-ttu-id="d056b-128">Exchange вызовет URI, чтобы получить вложение, а веб-службе, которая защищает файл, потребуется использовать токен для проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="d056b-128">Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.</span></span>

<span data-ttu-id="d056b-p107">Следующий пример JavaScript — это надстройка создания, которая прикрепляет файл picture.png с веб-сервера к создаваемому сообщению или встрече. Метод обратного вызова принимает `asyncResult` в качестве параметра, проверяет состояние результата и получает его идентификатор, если метод выполнен успешно.</span><span class="sxs-lookup"><span data-stu-id="d056b-p107">The following JavaScript example is a compose add-in that attaches a file, picture.png, from a web server to the message or appointment being composed. The callback method takes `asyncResult` as a parameter, checks for the result status, and gets the attachment ID if the method succeeds.</span></span>

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

### <a name="attach-an-outlook-item"></a><span data-ttu-id="d056b-131">Прикрепить элемент Outlook</span><span class="sxs-lookup"><span data-stu-id="d056b-131">Attach an Outlook item</span></span>

<span data-ttu-id="d056b-132">Вы можете прикрепить элемент Outlook (например, электронную почту, календарь или контактный элемент) к сообщению или встрече в форме записи, указав Exchange веб-службы (EWS) ИД элемента и используя `addItemAttachmentAsync` метод.</span><span class="sxs-lookup"><span data-stu-id="d056b-132">You can attach an Outlook item (for example, email, calendar, or contact item) to a message or appointment in a compose form by specifying the Exchange Web Services (EWS) ID of the item and using the `addItemAttachmentAsync` method.</span></span> <span data-ttu-id="d056b-133">Вы можете получить EWS-ID элемента электронной почты, календаря, контакта или задачи в почтовом ящике пользователя с помощью метода [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) и доступа к операции EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation).</span><span class="sxs-lookup"><span data-stu-id="d056b-133">You can get the EWS ID of an email, calendar, contact, or task item in the user's mailbox by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method and accessing the EWS operation [FindItem](/exchange/client-developer/web-service-reference/finditem-operation).</span></span> <span data-ttu-id="d056b-134">Свойство [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) также предоставляет идентификатор EWS существующего элемента в форме чтения.</span><span class="sxs-lookup"><span data-stu-id="d056b-134">The [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property also provides the EWS ID of an existing item in a read form.</span></span>

<span data-ttu-id="d056b-135">Следующая функция JavaScript расширяет первый пример выше и добавляет элемент в качестве вложения в составную электронную почту или `addItemAttachment` встречу.</span><span class="sxs-lookup"><span data-stu-id="d056b-135">The following JavaScript function, `addItemAttachment`, extends the first example above, and adds an item as an attachment to the email or appointment that is being composed.</span></span> <span data-ttu-id="d056b-136">В качестве параметра функция принимает идентификатор EWS прикрепляемого элемента.</span><span class="sxs-lookup"><span data-stu-id="d056b-136">The function takes as an argument the EWS ID of the item that is to be attached.</span></span> <span data-ttu-id="d056b-137">Если присоединение успешно, он получает ID вложения для дальнейшей обработки, включая удаление этого вложения в том же сеансе.</span><span class="sxs-lookup"><span data-stu-id="d056b-137">If attaching succeeds, it gets the attachment ID for further processing, including removing that attachment in the same session.</span></span>

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
> <span data-ttu-id="d056b-138">Надстройку можно использовать для прикрепления экземпляра повторяющейся встречи в Outlook в Интернете или на мобильных устройствах.</span><span class="sxs-lookup"><span data-stu-id="d056b-138">You can use a compose add-in to attach an instance of a recurring appointment in Outlook on the web or on mobile devices.</span></span> <span data-ttu-id="d056b-139">Однако в клиенте Outlook настольного компьютера попытка прикрепить экземпляр приведет к присоединению повторяющейся серии (родительского назначения).</span><span class="sxs-lookup"><span data-stu-id="d056b-139">However, in a supporting Outlook desktop client, attempting to attach an instance would result in attaching the recurring series (the parent appointment).</span></span>

## <a name="get-attachments"></a><span data-ttu-id="d056b-140">Получение вложений</span><span class="sxs-lookup"><span data-stu-id="d056b-140">Get attachments</span></span>

<span data-ttu-id="d056b-141">API для получения вложений в режиме композитации доступны из набора требований [1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="d056b-141">APIs to get attachments in compose mode are available from [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

- [<span data-ttu-id="d056b-142">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="d056b-142">getAttachmentsAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="d056b-143">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="d056b-143">getAttachmentContentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

<span data-ttu-id="d056b-144">Вы можете использовать [метод getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) для получения вложений сообщения или записи на прием.</span><span class="sxs-lookup"><span data-stu-id="d056b-144">You can use the [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method to get the attachments of the message or appointment being composed.</span></span>

<span data-ttu-id="d056b-145">Чтобы получить содержимое вложения, можно использовать [метод getAttachmentContentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="d056b-145">To get an attachment's content, you can use the [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="d056b-146">Поддерживаемые форматы перечислены в перечислении [AttachmentContentFormat.](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)</span><span class="sxs-lookup"><span data-stu-id="d056b-146">The supported formats are listed in the [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) enum.</span></span>

<span data-ttu-id="d056b-147">Необходимо предоставить метод вызова, чтобы проверить состояние и любую ошибку с помощью объекта `AsyncResult` параметра вывода.</span><span class="sxs-lookup"><span data-stu-id="d056b-147">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="d056b-148">Вы также можете передать все дополнительные параметры методу вызова, используя необязательный `asyncContext` параметр.</span><span class="sxs-lookup"><span data-stu-id="d056b-148">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="d056b-149">В следующем примере JavaScript вы можете получить вложения и настроить отдельную обработку для каждого поддерживаемого формата вложений.</span><span class="sxs-lookup"><span data-stu-id="d056b-149">The following JavaScript example gets the attachments and allows you to set up distinct handling for each supported attachment format.</span></span>

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

## <a name="remove-an-attachment"></a><span data-ttu-id="d056b-150">Удаление вложения</span><span class="sxs-lookup"><span data-stu-id="d056b-150">Remove an attachment</span></span>

<span data-ttu-id="d056b-151">Вы можете удалить вложение файла или элемента из сообщения или элемента встречи в форме записи, указав соответствующий ID вложения при использовании метода [removeAttachmentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="d056b-151">You can remove a file or item attachment from a message or appointment item in a compose form by specifying the corresponding attachment ID when using the [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d056b-152">Если используется набор требований 1.7 или ранее, следует удалять только вложения, добавленные той же надстройкой в том же сеансе.</span><span class="sxs-lookup"><span data-stu-id="d056b-152">If you're using requirement set 1.7 or earlier, you should only remove attachments that the same add-in has added in the same session.</span></span>

<span data-ttu-id="d056b-153">Аналогично методу `addFileAttachmentAsync` `addItemAttachmentAsync` , и `getAttachmentsAsync` методам, является `removeAttachmentAsync` асинхронным методом.</span><span class="sxs-lookup"><span data-stu-id="d056b-153">Similar to the `addFileAttachmentAsync`, `addItemAttachmentAsync`, and `getAttachmentsAsync` methods, `removeAttachmentAsync` is an asynchronous method.</span></span> <span data-ttu-id="d056b-154">Необходимо предоставить метод вызова, чтобы проверить состояние и любую ошибку с помощью объекта `AsyncResult` параметра вывода.</span><span class="sxs-lookup"><span data-stu-id="d056b-154">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="d056b-155">Вы также можете передать все дополнительные параметры методу вызова, используя необязательный `asyncContext` параметр.</span><span class="sxs-lookup"><span data-stu-id="d056b-155">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="d056b-156">Следующая функция JavaScript продолжает расширять указанные выше примеры и удаляет указанное вложение из записи электронной почты или записи. `removeAttachment`</span><span class="sxs-lookup"><span data-stu-id="d056b-156">The following JavaScript function, `removeAttachment`, continues to extend the examples above, and removes the specified attachment from the email or appointment that is being composed.</span></span> <span data-ttu-id="d056b-157">В качестве аргумента функция принимает идентификатор вложения, которое требуется удалить.</span><span class="sxs-lookup"><span data-stu-id="d056b-157">The function takes as an argument the ID of the attachment to be removed.</span></span> <span data-ttu-id="d056b-158">Вы можете получить ID вложения после успешного или методного вызова и использовать его в `addFileAttachmentAsync` `addFileAttachmentFromBase64Async` `addItemAttachmentAsync` последующем `removeAttachmentAsync` вызове метода.</span><span class="sxs-lookup"><span data-stu-id="d056b-158">You can obtain the ID of an attachment after a successful `addFileAttachmentAsync`, `addFileAttachmentFromBase64Async`, or `addItemAttachmentAsync` method call, and use it in a subsequent `removeAttachmentAsync` method call.</span></span> <span data-ttu-id="d056b-159">Вы также можете вызвать `getAttachmentsAsync` (вводится в наборе требований 1.8), чтобы получить вложения и их ID для этого сеанса надстройки.</span><span class="sxs-lookup"><span data-stu-id="d056b-159">You can also call `getAttachmentsAsync` (introduced in requirement set 1.8) to get the attachments and their IDs for that add-in session.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="d056b-160">См. также</span><span class="sxs-lookup"><span data-stu-id="d056b-160">See also</span></span>

- [<span data-ttu-id="d056b-161">Создание надстроек Outlook для форм создания</span><span class="sxs-lookup"><span data-stu-id="d056b-161">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="d056b-162">Асинхронное программирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d056b-162">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
