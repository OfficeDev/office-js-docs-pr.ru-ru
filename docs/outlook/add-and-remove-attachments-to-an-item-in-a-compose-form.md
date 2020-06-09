---
title: Добавление и удаление вложений в надстройке Outlook
description: Можно использовать различные API вложений для управления файлами или элементами Outlook, связанными с элементом, создаваемым пользователем.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d162ae4c0fa8059376a3c55463080e38679d9a01
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611675"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a><span data-ttu-id="89ae5-103">Управление вложениями элемента в форме создания в Outlook</span><span class="sxs-lookup"><span data-stu-id="89ae5-103">Manage an item's attachments in a compose form in Outlook</span></span>

<span data-ttu-id="89ae5-104">API JavaScript для Office предоставляет несколько интерфейсов API, которые можно использовать для управления вложениями элементов при создании пользователем.</span><span class="sxs-lookup"><span data-stu-id="89ae5-104">The Office JavaScript API provides several APIs you can use to manage an item's attachments when the user is composing.</span></span>

## <a name="attach-a-file-or-outlook-item"></a><span data-ttu-id="89ae5-105">Вложение файла или элемента Outlook</span><span class="sxs-lookup"><span data-stu-id="89ae5-105">Attach a file or Outlook item</span></span>

<span data-ttu-id="89ae5-106">Вы можете прикрепить файл или элемент Outlook к форме создания, используя метод, который соответствует типу вложения.</span><span class="sxs-lookup"><span data-stu-id="89ae5-106">You can attach a file or Outlook item to a compose form by using the method that's appropriate for the type of attachment.</span></span>

- <span data-ttu-id="89ae5-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Прикрепление файла</span><span class="sxs-lookup"><span data-stu-id="89ae5-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file</span></span>
- <span data-ttu-id="89ae5-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): присоединение файла с помощью строки Base64</span><span class="sxs-lookup"><span data-stu-id="89ae5-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file using its base64 string</span></span>
- <span data-ttu-id="89ae5-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): присоединение элемента Outlook</span><span class="sxs-lookup"><span data-stu-id="89ae5-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach an Outlook item</span></span>

<span data-ttu-id="89ae5-110">Это асинхронные методы, что означает, что выполнение может пройти без ожидания завершения действия.</span><span class="sxs-lookup"><span data-stu-id="89ae5-110">These are asynchronous methods, which means execution can go on without waiting for the action to complete.</span></span> <span data-ttu-id="89ae5-111">В зависимости от исходного расположения и размера добавляемого вложения для завершения асинхронного вызова может потребоваться некоторое время.</span><span class="sxs-lookup"><span data-stu-id="89ae5-111">Depending on the original location and size of the attachment being added, the asynchronous call may take a while to complete.</span></span>

<span data-ttu-id="89ae5-112">Если какие-то задачи зависят от завершения действия, их следует выполнять в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="89ae5-112">If there are tasks that depend on the action to complete, you should carry out those tasks in a callback method.</span></span> <span data-ttu-id="89ae5-113">Этот метод обратного вызова является необязательным и вызывается после завершения отправки вложения.</span><span class="sxs-lookup"><span data-stu-id="89ae5-113">This callback method is optional and is invoked when the attachment upload has completed.</span></span> <span data-ttu-id="89ae5-114">Метод обратного вызова принимает объект [AsyncResult](/javascript/api/office/office.asyncresult) как выходной параметр, который содержит состояние, ошибку и возвращаемое значение при добавлении вложения.</span><span class="sxs-lookup"><span data-stu-id="89ae5-114">The callback method takes an [AsyncResult](/javascript/api/office/office.asyncresult) object as an output parameter that provides any status, error, and returned value from adding the attachment.</span></span> <span data-ttu-id="89ae5-115">Если для обратного вызова требуются дополнительные параметры, их можно указать в необязательном параметре `options.asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="89ae5-115">If the callback requires any extra parameters, you can specify them in the optional `options.asyncContext` parameter.</span></span> <span data-ttu-id="89ae5-116">Параметр `options.asyncContext` может относиться к любому типу, поддерживаемому методом обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="89ae5-116">`options.asyncContext` can be of any type that your callback method expects.</span></span>

<span data-ttu-id="89ae5-p103">Например, можно определить параметр `options.asyncContext` как объект JSON, содержащий одну или несколько пар "ключ-значение". Дополнительные примеры передачи необязательных параметров в асинхронные методы для платформы надстроек Office см. в статье [Асинхронное программирование в надстройках Office](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). Ниже показано, как использовать параметр `asyncContext` для передачи двух аргументов методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="89ae5-p103">For example, you can define `options.asyncContext` as a JSON object that contains one or more key-value pairs. You can find more examples about passing optional parameters to asynchronous methods in the Office Add-ins platform in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). The following example shows how to use the `asyncContext` parameter to pass 2 arguments to a callback method:</span></span>

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

<span data-ttu-id="89ae5-p104">Успешность обратного вызова асинхронного метода можно проверить с помощью свойств `status` и `error` объекта `AsyncResult`. Если операция вложения завершается успешно, вы можете использовать свойство `AsyncResult.value`, чтобы получить идентификатор вложения. Это целое число, которое можно использовать в дальнейшем, чтобы удалить вложение.</span><span class="sxs-lookup"><span data-stu-id="89ae5-p104">You can check for success or error of an asynchronous method call in the callback method using the `status` and `error` properties of the `AsyncResult` object. If the attaching completes successfully, you can use the `AsyncResult.value` property to get the attachment ID. The attachment ID is an integer which you can subsequently use to remove the attachment.</span></span>

> [!NOTE]
> <span data-ttu-id="89ae5-122">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено той же надстройкой в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="89ae5-122">As a best practice, you should use the attachment ID to remove an attachment only if the same add-in has added that attachment in the same session.</span></span> <span data-ttu-id="89ae5-123">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="89ae5-123">In Outlook on the web and mobile devices, the attachment ID is valid only within the same session.</span></span> <span data-ttu-id="89ae5-124">Сеанс завершается, когда пользователь закрывает надстройку или начинает создавать элемент во встроенной форме и затем продолжает работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="89ae5-124">A session is over when the user closes the add-in, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

### <a name="attach-a-file"></a><span data-ttu-id="89ae5-125">Вложение файла</span><span class="sxs-lookup"><span data-stu-id="89ae5-125">Attach a file</span></span>

<span data-ttu-id="89ae5-126">Вы можете прикрепить файл к сообщению или встрече в форме создания, используя `addFileAttachmentAsync` метод и указав универсальный код ресурса (URI) файла.</span><span class="sxs-lookup"><span data-stu-id="89ae5-126">You can attach a file to a message or appointment in a compose form by using the `addFileAttachmentAsync` method and specifying the URI of the file.</span></span> <span data-ttu-id="89ae5-127">Кроме того, можно использовать `addFileAttachmentFromBase64Async` метод, но указать строку Base64 в качестве входных данных.</span><span class="sxs-lookup"><span data-stu-id="89ae5-127">You can also use the `addFileAttachmentFromBase64Async` method but specify the base64 string as input.</span></span> <span data-ttu-id="89ae5-128">Если файл защищен, можно добавить соответствующее удостоверение или токен проверки подлинности как параметр строки запроса URI.</span><span class="sxs-lookup"><span data-stu-id="89ae5-128">If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter.</span></span> <span data-ttu-id="89ae5-129">Exchange вызовет URI, чтобы получить вложение, а веб-службе, которая защищает файл, потребуется использовать токен для проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="89ae5-129">Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.</span></span>

<span data-ttu-id="89ae5-p107">Следующий пример JavaScript — это надстройка создания, которая прикрепляет файл picture.png с веб-сервера к создаваемому сообщению или встрече. Метод обратного вызова принимает `asyncResult` в качестве параметра, проверяет состояние результата и получает его идентификатор, если метод выполнен успешно.</span><span class="sxs-lookup"><span data-stu-id="89ae5-p107">The following JavaScript example is a compose add-in that attaches a file, picture.png, from a web server to the message or appointment being composed. The callback method takes `asyncResult` as a parameter, checks for the result status, and gets the attachment ID if the method succeeds.</span></span>

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

### <a name="attach-an-outlook-item"></a><span data-ttu-id="89ae5-132">Вложение элемента Outlook</span><span class="sxs-lookup"><span data-stu-id="89ae5-132">Attach an Outlook item</span></span>

<span data-ttu-id="89ae5-p108">Вы можете прикрепить элемент Outlook (например, электронное сообщение, элемент календаря или контакт) к сообщению или встрече в форме создания, указав идентификатор элемента в веб-службах Exchange (EWS) и вызвав метод `addItemAttachmentAsync`. Вы можете получить идентификатор EWS для элемента сообщения, календаря, контакта или задачи в почтовом ящике пользователя, вызвав метод [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) и используя операцию EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). Свойство [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) также предоставляет идентификатор EWS существующего элемента в форме чтения.</span><span class="sxs-lookup"><span data-stu-id="89ae5-p108">You can attach an Outlook item (for example, email, calendar, or contact item) to a message or appointment in a compose form by specifying the Exchange Web Services (EWS) ID of the item and using the `addItemAttachmentAsync` method. You can get the EWS ID of an email, calendar, contact or task item in the user's mailbox by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method and accessing the EWS operation [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). The [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property also provides the EWS ID of an existing item in a read form.</span></span>

<span data-ttu-id="89ae5-136">Приведенная ниже функция JavaScript, `addItemAttachment` которая расширяет первый пример выше и добавляет элемент в качестве вложения в создаваемую электронную почту или встречу.</span><span class="sxs-lookup"><span data-stu-id="89ae5-136">The following JavaScript function, `addItemAttachment`, extends the first example above, and adds an item as an attachment to the email or appointment that is being composed.</span></span> <span data-ttu-id="89ae5-137">В качестве параметра функция принимает идентификатор EWS прикрепляемого элемента.</span><span class="sxs-lookup"><span data-stu-id="89ae5-137">The function takes as an argument the EWS ID of the item that is to be attached.</span></span> <span data-ttu-id="89ae5-138">В случае успешного присоединения он получает идентификатор вложения для дальнейшей обработки, включая удаление этого вложения в том же сеансе.</span><span class="sxs-lookup"><span data-stu-id="89ae5-138">If attaching succeeds, it gets the attachment ID for further processing, including removing that attachment in the same session.</span></span>

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
> <span data-ttu-id="89ae5-139">Вы можете использовать надстройку создания, чтобы вложить экземпляр повторяющейся встречи в Outlook в Интернете или на мобильных устройствах.</span><span class="sxs-lookup"><span data-stu-id="89ae5-139">You can use a compose add-in to attach an instance of a recurring appointment in Outlook on the web or mobile devices.</span></span> <span data-ttu-id="89ae5-140">Однако в расширенном клиенте Outlook попытка прикрепить такой экземпляр приведет к прикреплению ряда повторений (основной встречи).</span><span class="sxs-lookup"><span data-stu-id="89ae5-140">However, in a supporting Outlook rich client, attempting to attach an instance would result in attaching the recurring series (the master appointment).</span></span>

## <a name="get-attachments"></a><span data-ttu-id="89ae5-141">Получение вложений</span><span class="sxs-lookup"><span data-stu-id="89ae5-141">Get attachments</span></span>

<span data-ttu-id="89ae5-142">Вы можете использовать метод [жетаттачментсасинк](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) для получения вложений создаваемого сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="89ae5-142">You can use the [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method to get the attachments of the message or appointment being composed.</span></span>

<span data-ttu-id="89ae5-143">Чтобы получить содержимое вложения, можно использовать метод [жетаттачментконтентасинк](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="89ae5-143">To get an attachment's content, you can use the [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="89ae5-144">Поддерживаемые форматы перечислены в перечислении [аттачментконтентформат](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) .</span><span class="sxs-lookup"><span data-stu-id="89ae5-144">The supported formats are listed in the [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) enum.</span></span>

<span data-ttu-id="89ae5-145">Необходимо предоставить метод обратного вызова для проверки состояния и любой ошибки с помощью `AsyncResult` объекта Output Parameter.</span><span class="sxs-lookup"><span data-stu-id="89ae5-145">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="89ae5-146">Кроме того, можно передать дополнительные параметры в метод обратного вызова, используя необязательный `asyncContext` параметр.</span><span class="sxs-lookup"><span data-stu-id="89ae5-146">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="89ae5-147">В приведенном ниже примере JavaScript показано, как получить вложения и настроить индивидуальную обработку для каждого поддерживаемого формата вложений.</span><span class="sxs-lookup"><span data-stu-id="89ae5-147">The following JavaScript example gets the attachments and allows you to set up distinct handling for each supported attachment format.</span></span>

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

## <a name="remove-an-attachment"></a><span data-ttu-id="89ae5-148">Удаление вложения</span><span class="sxs-lookup"><span data-stu-id="89ae5-148">Remove an attachment</span></span>

<span data-ttu-id="89ae5-149">Вы можете удалить вложение из элемента сообщения или встречи в форме создания, указав соответствующий идентификатор вложения и вызвав метод [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods).</span><span class="sxs-lookup"><span data-stu-id="89ae5-149">You can remove a file or item attachment from a message or appointment item in a compose form by specifying the corresponding attachment ID and using the [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="89ae5-150">Удалять можно только вложения, добавленные одной и той же надстройкой в одном сеансе.</span><span class="sxs-lookup"><span data-stu-id="89ae5-150">You should only remove attachments that the same add-in has added in the same session.</span></span> <span data-ttu-id="89ae5-151">Аналогично `addFileAttachmentAsync` `addItemAttachmentAsync` методам и, `removeAttachmentAsync` — это асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="89ae5-151">Similar to the `addFileAttachmentAsync` and `addItemAttachmentAsync` methods, `removeAttachmentAsync` is an asynchronous method.</span></span> <span data-ttu-id="89ae5-152">Необходимо предоставить метод обратного вызова для проверки состояния и любой ошибки с помощью `AsyncResult` объекта Output Parameter.</span><span class="sxs-lookup"><span data-stu-id="89ae5-152">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="89ae5-153">Кроме того, можно передать дополнительные параметры в метод обратного вызова, используя необязательный `asyncContext` параметр.</span><span class="sxs-lookup"><span data-stu-id="89ae5-153">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="89ae5-154">Приведенная ниже функция JavaScript `removeAttachment` продолжает расширять приведенные выше примеры и удаляет указанное вложение из создаваемого сообщения электронной почты или встречи.</span><span class="sxs-lookup"><span data-stu-id="89ae5-154">The following JavaScript function, `removeAttachment`, continues to extend the examples above, and removes the specified attachment from the email or appointment that is being composed.</span></span> <span data-ttu-id="89ae5-155">В качестве аргумента функция принимает идентификатор вложения, которое требуется удалить.</span><span class="sxs-lookup"><span data-stu-id="89ae5-155">The function takes as an argument the ID of the attachment to be removed.</span></span> <span data-ttu-id="89ae5-156">Вы можете получить идентификатор вложения после успешного `addFileAttachmentAsync` `addFileAttachmentFromBase64Async` вызова, или `addItemAttachmentAsync` вызова метода и сохранить его для последующего `removeAttachmentAsync` вызова метода.</span><span class="sxs-lookup"><span data-stu-id="89ae5-156">You can obtain the ID of an attachment after a successful `addFileAttachmentAsync`, `addFileAttachmentFromBase64Async`, or `addItemAttachmentAsync` method call, and store it for a subsequent `removeAttachmentAsync` method call.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="89ae5-157">См. также</span><span class="sxs-lookup"><span data-stu-id="89ae5-157">See also</span></span>

- [<span data-ttu-id="89ae5-158">Создание надстроек Outlook для форм создания</span><span class="sxs-lookup"><span data-stu-id="89ae5-158">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="89ae5-159">Асинхронное программирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="89ae5-159">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
