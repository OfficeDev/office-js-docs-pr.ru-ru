---
title: Office. Context. Mailbox. Item — набор требований 1,1
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: fa5ff45e7a01f7f1286527eec74a70dc45fd59f6
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064405"
---
# <a name="item"></a><span data-ttu-id="110a2-102">item</span><span class="sxs-lookup"><span data-stu-id="110a2-102">item</span></span>

### <span data-ttu-id="110a2-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="110a2-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="110a2-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="110a2-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="110a2-107">Requirements</span></span>

|<span data-ttu-id="110a2-108">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-108">Requirement</span></span>| <span data-ttu-id="110a2-109">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-110">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-111">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-111">1.0</span></span>|
|[<span data-ttu-id="110a2-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="110a2-113">Restricted</span></span>|
|[<span data-ttu-id="110a2-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="110a2-116">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-116">Example</span></span>

<span data-ttu-id="110a2-117">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="110a2-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```

### <a name="members"></a><span data-ttu-id="110a2-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="110a2-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="110a2-119">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="110a2-119">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="110a2-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-122">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="110a2-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="110a2-123">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="110a2-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-124">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-124">Type</span></span>

*   <span data-ttu-id="110a2-125">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="110a2-125">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-126">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-126">Requirements</span></span>

|<span data-ttu-id="110a2-127">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-127">Requirement</span></span>| <span data-ttu-id="110a2-128">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-129">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-130">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-130">1.0</span></span>|
|[<span data-ttu-id="110a2-131">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-132">ReadItem</span></span>|
|[<span data-ttu-id="110a2-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-134">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-135">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-135">Example</span></span>

<span data-ttu-id="110a2-136">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="110a2-137">СК: [получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-137">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-138">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="110a2-139">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="110a2-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-140">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-140">Type</span></span>

*   [<span data-ttu-id="110a2-141">Получатели</span><span class="sxs-lookup"><span data-stu-id="110a2-141">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="110a2-142">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-142">Requirements</span></span>

|<span data-ttu-id="110a2-143">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-143">Requirement</span></span>| <span data-ttu-id="110a2-144">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-145">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-146">1.1</span><span class="sxs-lookup"><span data-stu-id="110a2-146">1.1</span></span>|
|[<span data-ttu-id="110a2-147">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-148">ReadItem</span></span>|
|[<span data-ttu-id="110a2-149">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-150">Создание</span><span class="sxs-lookup"><span data-stu-id="110a2-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-151">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="110a2-152">основной текст: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-152">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-153">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-154">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-154">Type</span></span>

*   [<span data-ttu-id="110a2-155">Body</span><span class="sxs-lookup"><span data-stu-id="110a2-155">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="110a2-156">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-156">Requirements</span></span>

|<span data-ttu-id="110a2-157">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-157">Requirement</span></span>| <span data-ttu-id="110a2-158">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-160">1.1</span><span class="sxs-lookup"><span data-stu-id="110a2-160">1.1</span></span>|
|[<span data-ttu-id="110a2-161">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-162">ReadItem</span></span>|
|[<span data-ttu-id="110a2-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-165">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-165">Example</span></span>

<span data-ttu-id="110a2-166">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="110a2-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="110a2-167">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="110a2-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="110a2-168">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-168">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-169">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="110a2-170">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="110a2-171">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="110a2-171">Read mode</span></span>

<span data-ttu-id="110a2-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="110a2-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="110a2-174">Режим создания</span><span class="sxs-lookup"><span data-stu-id="110a2-174">Compose mode</span></span>

<span data-ttu-id="110a2-175">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="110a2-176">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-176">Type</span></span>

*   <span data-ttu-id="110a2-177">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-177">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-178">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-178">Requirements</span></span>

|<span data-ttu-id="110a2-179">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-179">Requirement</span></span>| <span data-ttu-id="110a2-180">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-181">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-182">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-182">1.0</span></span>|
|[<span data-ttu-id="110a2-183">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-184">ReadItem</span></span>|
|[<span data-ttu-id="110a2-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-186">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="110a2-187">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="110a2-187">(nullable) conversationId: String</span></span>

<span data-ttu-id="110a2-188">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="110a2-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="110a2-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="110a2-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="110a2-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="110a2-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-193">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-193">Type</span></span>

*   <span data-ttu-id="110a2-194">String</span><span class="sxs-lookup"><span data-stu-id="110a2-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-195">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-195">Requirements</span></span>

|<span data-ttu-id="110a2-196">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-196">Requirement</span></span>| <span data-ttu-id="110a2-197">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-198">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-199">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-199">1.0</span></span>|
|[<span data-ttu-id="110a2-200">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-201">ReadItem</span></span>|
|[<span data-ttu-id="110a2-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-203">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-204">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="110a2-205">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="110a2-205">dateTimeCreated: Date</span></span>

<span data-ttu-id="110a2-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-208">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-208">Type</span></span>

*   <span data-ttu-id="110a2-209">Дата</span><span class="sxs-lookup"><span data-stu-id="110a2-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-210">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-210">Requirements</span></span>

|<span data-ttu-id="110a2-211">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-211">Requirement</span></span>| <span data-ttu-id="110a2-212">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-213">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-214">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-214">1.0</span></span>|
|[<span data-ttu-id="110a2-215">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-216">ReadItem</span></span>|
|[<span data-ttu-id="110a2-217">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-218">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-219">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="110a2-220">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="110a2-220">dateTimeModified: Date</span></span>

<span data-ttu-id="110a2-221">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-221">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="110a2-222">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-223">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="110a2-223">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-224">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-224">Type</span></span>

*   <span data-ttu-id="110a2-225">Дата</span><span class="sxs-lookup"><span data-stu-id="110a2-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-226">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-226">Requirements</span></span>

|<span data-ttu-id="110a2-227">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-227">Requirement</span></span>| <span data-ttu-id="110a2-228">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-229">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-230">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-230">1.0</span></span>|
|[<span data-ttu-id="110a2-231">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-232">ReadItem</span></span>|
|[<span data-ttu-id="110a2-233">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-234">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-235">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="110a2-236">конец: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="110a2-236">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-237">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="110a2-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="110a2-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="110a2-240">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="110a2-240">Read mode</span></span>

<span data-ttu-id="110a2-241">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="110a2-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="110a2-242">Режим создания</span><span class="sxs-lookup"><span data-stu-id="110a2-242">Compose mode</span></span>

<span data-ttu-id="110a2-243">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="110a2-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="110a2-244">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="110a2-244">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="110a2-245">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="110a2-246">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-246">Type</span></span>

*   <span data-ttu-id="110a2-247">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-247">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-248">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-248">Requirements</span></span>

|<span data-ttu-id="110a2-249">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-249">Requirement</span></span>| <span data-ttu-id="110a2-250">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-251">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-252">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-252">1.0</span></span>|
|[<span data-ttu-id="110a2-253">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-254">ReadItem</span></span>|
|[<span data-ttu-id="110a2-255">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-256">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="110a2-257">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-257">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="110a2-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="110a2-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-262">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="110a2-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-263">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-263">Type</span></span>

*   [<span data-ttu-id="110a2-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="110a2-264">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="110a2-265">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-265">Requirements</span></span>

|<span data-ttu-id="110a2-266">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-266">Requirement</span></span>| <span data-ttu-id="110a2-267">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-268">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-269">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-269">1.0</span></span>|
|[<span data-ttu-id="110a2-270">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-271">ReadItem</span></span>|
|[<span data-ttu-id="110a2-272">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-273">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-274">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="110a2-275">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="110a2-275">internetMessageId: String</span></span>

<span data-ttu-id="110a2-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-278">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-278">Type</span></span>

*   <span data-ttu-id="110a2-279">String</span><span class="sxs-lookup"><span data-stu-id="110a2-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-280">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-280">Requirements</span></span>

|<span data-ttu-id="110a2-281">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-281">Requirement</span></span>| <span data-ttu-id="110a2-282">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-283">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-284">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-284">1.0</span></span>|
|[<span data-ttu-id="110a2-285">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-286">ReadItem</span></span>|
|[<span data-ttu-id="110a2-287">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-288">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-289">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="110a2-290">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="110a2-290">itemClass: String</span></span>

<span data-ttu-id="110a2-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="110a2-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="110a2-295">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-295">Type</span></span> | <span data-ttu-id="110a2-296">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-296">Description</span></span> | <span data-ttu-id="110a2-297">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="110a2-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="110a2-298">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="110a2-298">Appointment items</span></span> | <span data-ttu-id="110a2-299">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="110a2-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="110a2-300">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="110a2-300">Message items</span></span> | <span data-ttu-id="110a2-301">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="110a2-302">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="110a2-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-303">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-303">Type</span></span>

*   <span data-ttu-id="110a2-304">String</span><span class="sxs-lookup"><span data-stu-id="110a2-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-305">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-305">Requirements</span></span>

|<span data-ttu-id="110a2-306">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-306">Requirement</span></span>| <span data-ttu-id="110a2-307">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-308">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-309">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-309">1.0</span></span>|
|[<span data-ttu-id="110a2-310">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-311">ReadItem</span></span>|
|[<span data-ttu-id="110a2-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-313">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-314">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="110a2-315">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="110a2-315">(nullable) itemId: String</span></span>

<span data-ttu-id="110a2-316">Получает идентификатор элемента веб-служб Exchange для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-316">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="110a2-317">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-317">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-318">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="110a2-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="110a2-319">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="110a2-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="110a2-320">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="110a2-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="110a2-321">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="110a2-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-322">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-322">Type</span></span>

*   <span data-ttu-id="110a2-323">String</span><span class="sxs-lookup"><span data-stu-id="110a2-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-324">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-324">Requirements</span></span>

|<span data-ttu-id="110a2-325">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-325">Requirement</span></span>| <span data-ttu-id="110a2-326">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-327">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-328">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-328">1.0</span></span>|
|[<span data-ttu-id="110a2-329">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-330">ReadItem</span></span>|
|[<span data-ttu-id="110a2-331">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-332">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-333">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-333">Example</span></span>

<span data-ttu-id="110a2-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="110a2-336">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-336">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-337">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="110a2-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="110a2-338">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="110a2-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-339">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-339">Type</span></span>

*   [<span data-ttu-id="110a2-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="110a2-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="110a2-341">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-341">Requirements</span></span>

|<span data-ttu-id="110a2-342">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-342">Requirement</span></span>| <span data-ttu-id="110a2-343">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-344">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-345">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-345">1.0</span></span>|
|[<span data-ttu-id="110a2-346">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-347">ReadItem</span></span>|
|[<span data-ttu-id="110a2-348">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-349">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-350">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="110a2-351">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="110a2-351">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-352">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="110a2-353">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="110a2-353">Read mode</span></span>

<span data-ttu-id="110a2-354">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="110a2-355">Режим создания</span><span class="sxs-lookup"><span data-stu-id="110a2-355">Compose mode</span></span>

<span data-ttu-id="110a2-356">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="110a2-357">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-357">Type</span></span>

*   <span data-ttu-id="110a2-358">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-358">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-359">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-359">Requirements</span></span>

|<span data-ttu-id="110a2-360">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-360">Requirement</span></span>| <span data-ttu-id="110a2-361">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-363">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-363">1.0</span></span>|
|[<span data-ttu-id="110a2-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-365">ReadItem</span></span>|
|[<span data-ttu-id="110a2-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="110a2-368">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="110a2-368">normalizedSubject: String</span></span>

<span data-ttu-id="110a2-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="110a2-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="110a2-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-373">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-373">Type</span></span>

*   <span data-ttu-id="110a2-374">String</span><span class="sxs-lookup"><span data-stu-id="110a2-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-375">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-375">Requirements</span></span>

|<span data-ttu-id="110a2-376">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-376">Requirement</span></span>| <span data-ttu-id="110a2-377">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-378">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-379">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-379">1.0</span></span>|
|[<span data-ttu-id="110a2-380">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-381">ReadItem</span></span>|
|[<span data-ttu-id="110a2-382">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-383">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-384">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="110a2-385">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-385">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-386">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="110a2-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="110a2-387">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="110a2-388">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="110a2-388">Read mode</span></span>

<span data-ttu-id="110a2-389">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="110a2-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="110a2-390">Режим создания</span><span class="sxs-lookup"><span data-stu-id="110a2-390">Compose mode</span></span>

<span data-ttu-id="110a2-391">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="110a2-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="110a2-392">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-392">Type</span></span>

*   <span data-ttu-id="110a2-393">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-393">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-394">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-394">Requirements</span></span>

|<span data-ttu-id="110a2-395">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-395">Requirement</span></span>| <span data-ttu-id="110a2-396">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-397">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-398">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-398">1.0</span></span>|
|[<span data-ttu-id="110a2-399">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-400">ReadItem</span></span>|
|[<span data-ttu-id="110a2-401">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-402">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="110a2-403">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-403">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-406">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-406">Type</span></span>

*   [<span data-ttu-id="110a2-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="110a2-407">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="110a2-408">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-408">Requirements</span></span>

|<span data-ttu-id="110a2-409">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-409">Requirement</span></span>| <span data-ttu-id="110a2-410">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-412">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-412">1.0</span></span>|
|[<span data-ttu-id="110a2-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-414">ReadItem</span></span>|
|[<span data-ttu-id="110a2-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-417">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="110a2-418">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-418">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-419">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="110a2-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="110a2-420">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="110a2-421">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="110a2-421">Read mode</span></span>

<span data-ttu-id="110a2-422">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="110a2-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="110a2-423">Режим создания</span><span class="sxs-lookup"><span data-stu-id="110a2-423">Compose mode</span></span>

<span data-ttu-id="110a2-424">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="110a2-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="110a2-425">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-425">Type</span></span>

*   <span data-ttu-id="110a2-426">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-426">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-427">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-427">Requirements</span></span>

|<span data-ttu-id="110a2-428">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-428">Requirement</span></span>| <span data-ttu-id="110a2-429">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-430">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-431">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-431">1.0</span></span>|
|[<span data-ttu-id="110a2-432">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-433">ReadItem</span></span>|
|[<span data-ttu-id="110a2-434">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-435">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="110a2-436">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-436">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="110a2-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="110a2-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="110a2-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-441">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="110a2-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="110a2-442">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-442">Type</span></span>

*   [<span data-ttu-id="110a2-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="110a2-443">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="110a2-444">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-444">Requirements</span></span>

|<span data-ttu-id="110a2-445">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-445">Requirement</span></span>| <span data-ttu-id="110a2-446">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-447">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-448">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-448">1.0</span></span>|
|[<span data-ttu-id="110a2-449">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-450">ReadItem</span></span>|
|[<span data-ttu-id="110a2-451">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-452">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-453">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="110a2-454">Начало: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="110a2-454">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-455">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="110a2-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="110a2-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="110a2-458">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="110a2-458">Read mode</span></span>

<span data-ttu-id="110a2-459">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="110a2-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="110a2-460">Режим создания</span><span class="sxs-lookup"><span data-stu-id="110a2-460">Compose mode</span></span>

<span data-ttu-id="110a2-461">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="110a2-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="110a2-462">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="110a2-462">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="110a2-463">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="110a2-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="110a2-464">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-464">Type</span></span>

*   <span data-ttu-id="110a2-465">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-465">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-466">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-466">Requirements</span></span>

|<span data-ttu-id="110a2-467">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-467">Requirement</span></span>| <span data-ttu-id="110a2-468">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-469">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-470">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-470">1.0</span></span>|
|[<span data-ttu-id="110a2-471">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-472">ReadItem</span></span>|
|[<span data-ttu-id="110a2-473">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-474">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-474">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="110a2-475">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="110a2-475">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-476">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="110a2-477">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="110a2-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="110a2-478">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="110a2-478">Read mode</span></span>

<span data-ttu-id="110a2-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="110a2-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="110a2-481">Режим создания</span><span class="sxs-lookup"><span data-stu-id="110a2-481">Compose mode</span></span>

<span data-ttu-id="110a2-482">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="110a2-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="110a2-483">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-483">Type</span></span>

*   <span data-ttu-id="110a2-484">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-484">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-485">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-485">Requirements</span></span>

|<span data-ttu-id="110a2-486">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-486">Requirement</span></span>| <span data-ttu-id="110a2-487">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-488">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-489">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-489">1.0</span></span>|
|[<span data-ttu-id="110a2-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-491">ReadItem</span></span>|
|[<span data-ttu-id="110a2-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-493">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-493">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="110a2-494">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-494">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="110a2-495">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="110a2-496">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="110a2-497">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="110a2-497">Read mode</span></span>

<span data-ttu-id="110a2-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="110a2-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="110a2-500">Режим создания</span><span class="sxs-lookup"><span data-stu-id="110a2-500">Compose mode</span></span>

<span data-ttu-id="110a2-501">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="110a2-502">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-502">Type</span></span>

*   <span data-ttu-id="110a2-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-504">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-504">Requirements</span></span>

|<span data-ttu-id="110a2-505">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-505">Requirement</span></span>| <span data-ttu-id="110a2-506">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-507">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-508">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-508">1.0</span></span>|
|[<span data-ttu-id="110a2-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-510">ReadItem</span></span>|
|[<span data-ttu-id="110a2-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="110a2-513">Методы</span><span class="sxs-lookup"><span data-stu-id="110a2-513">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="110a2-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="110a2-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="110a2-515">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="110a2-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="110a2-516">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="110a2-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="110a2-517">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="110a2-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="110a2-518">Параметры</span><span class="sxs-lookup"><span data-stu-id="110a2-518">Parameters</span></span>

|<span data-ttu-id="110a2-519">Имя</span><span class="sxs-lookup"><span data-stu-id="110a2-519">Name</span></span>| <span data-ttu-id="110a2-520">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-520">Type</span></span>| <span data-ttu-id="110a2-521">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="110a2-521">Attributes</span></span>| <span data-ttu-id="110a2-522">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="110a2-523">String</span><span class="sxs-lookup"><span data-stu-id="110a2-523">String</span></span>||<span data-ttu-id="110a2-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="110a2-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="110a2-526">String</span><span class="sxs-lookup"><span data-stu-id="110a2-526">String</span></span>||<span data-ttu-id="110a2-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="110a2-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="110a2-529">Объект</span><span class="sxs-lookup"><span data-stu-id="110a2-529">Object</span></span>| <span data-ttu-id="110a2-530">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-530">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-531">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="110a2-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="110a2-532">Object</span><span class="sxs-lookup"><span data-stu-id="110a2-532">Object</span></span>| <span data-ttu-id="110a2-533">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-533">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-534">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="110a2-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="110a2-535">функция</span><span class="sxs-lookup"><span data-stu-id="110a2-535">function</span></span>| <span data-ttu-id="110a2-536">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-536">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-537">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="110a2-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="110a2-538">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="110a2-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="110a2-539">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="110a2-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="110a2-540">Ошибки</span><span class="sxs-lookup"><span data-stu-id="110a2-540">Errors</span></span>

| <span data-ttu-id="110a2-541">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="110a2-541">Error code</span></span> | <span data-ttu-id="110a2-542">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="110a2-543">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="110a2-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="110a2-544">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="110a2-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="110a2-545">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="110a2-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="110a2-546">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-546">Requirements</span></span>

|<span data-ttu-id="110a2-547">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-547">Requirement</span></span>| <span data-ttu-id="110a2-548">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-549">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-550">1.1</span><span class="sxs-lookup"><span data-stu-id="110a2-550">1.1</span></span>|
|[<span data-ttu-id="110a2-551">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="110a2-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="110a2-553">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-554">Создание</span><span class="sxs-lookup"><span data-stu-id="110a2-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-555">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-555">Example</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="110a2-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="110a2-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="110a2-557">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="110a2-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="110a2-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="110a2-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="110a2-561">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="110a2-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="110a2-562">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="110a2-562">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="110a2-563">Параметры</span><span class="sxs-lookup"><span data-stu-id="110a2-563">Parameters</span></span>

|<span data-ttu-id="110a2-564">Имя</span><span class="sxs-lookup"><span data-stu-id="110a2-564">Name</span></span>| <span data-ttu-id="110a2-565">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-565">Type</span></span>| <span data-ttu-id="110a2-566">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="110a2-566">Attributes</span></span>| <span data-ttu-id="110a2-567">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="110a2-568">String</span><span class="sxs-lookup"><span data-stu-id="110a2-568">String</span></span>||<span data-ttu-id="110a2-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="110a2-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="110a2-571">String</span><span class="sxs-lookup"><span data-stu-id="110a2-571">String</span></span>||<span data-ttu-id="110a2-572">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-572">The subject of the item to be attached.</span></span> <span data-ttu-id="110a2-573">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="110a2-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="110a2-574">Object</span><span class="sxs-lookup"><span data-stu-id="110a2-574">Object</span></span>| <span data-ttu-id="110a2-575">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-575">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-576">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="110a2-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="110a2-577">Объект</span><span class="sxs-lookup"><span data-stu-id="110a2-577">Object</span></span>| <span data-ttu-id="110a2-578">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-578">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-579">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="110a2-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="110a2-580">функция</span><span class="sxs-lookup"><span data-stu-id="110a2-580">function</span></span>| <span data-ttu-id="110a2-581">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-581">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-582">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="110a2-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="110a2-583">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="110a2-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="110a2-584">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="110a2-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="110a2-585">Ошибки</span><span class="sxs-lookup"><span data-stu-id="110a2-585">Errors</span></span>

| <span data-ttu-id="110a2-586">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="110a2-586">Error code</span></span> | <span data-ttu-id="110a2-587">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="110a2-588">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="110a2-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="110a2-589">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-589">Requirements</span></span>

|<span data-ttu-id="110a2-590">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-590">Requirement</span></span>| <span data-ttu-id="110a2-591">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-592">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-593">1.1</span><span class="sxs-lookup"><span data-stu-id="110a2-593">1.1</span></span>|
|[<span data-ttu-id="110a2-594">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="110a2-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="110a2-596">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-597">Создание</span><span class="sxs-lookup"><span data-stu-id="110a2-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-598">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-598">Example</span></span>

<span data-ttu-id="110a2-599">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="110a2-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="110a2-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="110a2-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="110a2-601">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-602">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="110a2-602">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="110a2-603">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="110a2-603">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="110a2-604">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="110a2-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-605">Возможность включать вложения в вызове `displayReplyAllForm` не поддерживается в наборе требований 1,1.</span><span class="sxs-lookup"><span data-stu-id="110a2-605">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="110a2-606">Добавлена поддержка вложений `displayReplyAllForm` в наборе требований 1,2 и выше.</span><span class="sxs-lookup"><span data-stu-id="110a2-606">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="110a2-607">Параметры</span><span class="sxs-lookup"><span data-stu-id="110a2-607">Parameters</span></span>

|<span data-ttu-id="110a2-608">Имя</span><span class="sxs-lookup"><span data-stu-id="110a2-608">Name</span></span>| <span data-ttu-id="110a2-609">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-609">Type</span></span>| <span data-ttu-id="110a2-610">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-610">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="110a2-611">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="110a2-611">String &#124; Object</span></span>| |<span data-ttu-id="110a2-p138">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="110a2-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="110a2-614">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="110a2-614">**OR**</span></span><br/><span data-ttu-id="110a2-p139">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="110a2-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="110a2-617">String</span><span class="sxs-lookup"><span data-stu-id="110a2-617">String</span></span> | <span data-ttu-id="110a2-618">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-618">&lt;optional&gt;</span></span> | <span data-ttu-id="110a2-p140">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="110a2-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="110a2-621">функция</span><span class="sxs-lookup"><span data-stu-id="110a2-621">function</span></span> | <span data-ttu-id="110a2-622">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-622">&lt;optional&gt;</span></span> | <span data-ttu-id="110a2-623">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="110a2-623">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="110a2-624">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-624">Requirements</span></span>

|<span data-ttu-id="110a2-625">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-625">Requirement</span></span>| <span data-ttu-id="110a2-626">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-627">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-628">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-628">1.0</span></span>|
|[<span data-ttu-id="110a2-629">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-630">ReadItem</span></span>|
|[<span data-ttu-id="110a2-631">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-632">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-632">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="110a2-633">Примеры</span><span class="sxs-lookup"><span data-stu-id="110a2-633">Examples</span></span>

<span data-ttu-id="110a2-634">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="110a2-634">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="110a2-635">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-635">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="110a2-636">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-636">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="110a2-637">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="110a2-637">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="110a2-638">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="110a2-638">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="110a2-639">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-639">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-640">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="110a2-640">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="110a2-641">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="110a2-641">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="110a2-642">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="110a2-642">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-643">Возможность включать вложения в вызове `displayReplyForm` не поддерживается в наборе требований 1,1.</span><span class="sxs-lookup"><span data-stu-id="110a2-643">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="110a2-644">Добавлена поддержка вложений `displayReplyForm` в наборе требований 1,2 и выше.</span><span class="sxs-lookup"><span data-stu-id="110a2-644">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="110a2-645">Параметры</span><span class="sxs-lookup"><span data-stu-id="110a2-645">Parameters</span></span>

|<span data-ttu-id="110a2-646">Имя</span><span class="sxs-lookup"><span data-stu-id="110a2-646">Name</span></span>| <span data-ttu-id="110a2-647">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-647">Type</span></span>| <span data-ttu-id="110a2-648">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-648">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="110a2-649">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="110a2-649">String &#124; Object</span></span>| | <span data-ttu-id="110a2-p142">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="110a2-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="110a2-652">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="110a2-652">**OR**</span></span><br/><span data-ttu-id="110a2-p143">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="110a2-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="110a2-655">String</span><span class="sxs-lookup"><span data-stu-id="110a2-655">String</span></span> | <span data-ttu-id="110a2-656">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-656">&lt;optional&gt;</span></span> | <span data-ttu-id="110a2-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="110a2-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="110a2-659">функция</span><span class="sxs-lookup"><span data-stu-id="110a2-659">function</span></span> | <span data-ttu-id="110a2-660">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-660">&lt;optional&gt;</span></span> | <span data-ttu-id="110a2-661">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="110a2-661">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="110a2-662">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-662">Requirements</span></span>

|<span data-ttu-id="110a2-663">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-663">Requirement</span></span>| <span data-ttu-id="110a2-664">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-665">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-666">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-666">1.0</span></span>|
|[<span data-ttu-id="110a2-667">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-668">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-668">ReadItem</span></span>|
|[<span data-ttu-id="110a2-669">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-670">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-670">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="110a2-671">Примеры</span><span class="sxs-lookup"><span data-stu-id="110a2-671">Examples</span></span>

<span data-ttu-id="110a2-672">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="110a2-672">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="110a2-673">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-673">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="110a2-674">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="110a2-674">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="110a2-675">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="110a2-675">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="110a2-676">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="110a2-676">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="110a2-677">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-677">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-678">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="110a2-678">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-679">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-679">Requirements</span></span>

|<span data-ttu-id="110a2-680">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-680">Requirement</span></span>| <span data-ttu-id="110a2-681">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-681">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-682">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-682">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-683">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-683">1.0</span></span>|
|[<span data-ttu-id="110a2-684">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-684">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-685">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-685">ReadItem</span></span>|
|[<span data-ttu-id="110a2-686">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-686">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-687">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-687">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="110a2-688">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="110a2-688">Returns:</span></span>

<span data-ttu-id="110a2-689">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="110a2-689">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="110a2-690">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-690">Example</span></span>

<span data-ttu-id="110a2-691">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-691">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="110a2-692">getEntitiesByType (entityType) → (Nullable) {Array. < (строка |[ Contact](/javascript/api/outlook/office.contact)|[](/javascript/api/outlook/office.meetingsuggestion)свойства meetingsuggestion|[](/javascript/api/outlook/office.phonenumber)PhoneNumber|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)? View = Outlook-JS-1,1) >}</span><span class="sxs-lookup"><span data-stu-id="110a2-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.1)>}</span></span>

<span data-ttu-id="110a2-693">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-693">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-694">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="110a2-694">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="110a2-695">Параметры</span><span class="sxs-lookup"><span data-stu-id="110a2-695">Parameters</span></span>

|<span data-ttu-id="110a2-696">Имя</span><span class="sxs-lookup"><span data-stu-id="110a2-696">Name</span></span>| <span data-ttu-id="110a2-697">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-697">Type</span></span>| <span data-ttu-id="110a2-698">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-698">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="110a2-699">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="110a2-699">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="110a2-700">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="110a2-700">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="110a2-701">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-701">Requirements</span></span>

|<span data-ttu-id="110a2-702">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-702">Requirement</span></span>| <span data-ttu-id="110a2-703">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-703">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-704">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-704">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-705">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-705">1.0</span></span>|
|[<span data-ttu-id="110a2-706">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-706">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-707">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="110a2-707">Restricted</span></span>|
|[<span data-ttu-id="110a2-708">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-708">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-709">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-709">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="110a2-710">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="110a2-710">Returns:</span></span>

<span data-ttu-id="110a2-711">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="110a2-711">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="110a2-712">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="110a2-712">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="110a2-713">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="110a2-713">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="110a2-714">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="110a2-714">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="110a2-715">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="110a2-715">Value of `entityType`</span></span> | <span data-ttu-id="110a2-716">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="110a2-716">Type of objects in returned array</span></span> | <span data-ttu-id="110a2-717">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-717">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="110a2-718">String</span><span class="sxs-lookup"><span data-stu-id="110a2-718">String</span></span> | <span data-ttu-id="110a2-719">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="110a2-719">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="110a2-720">Contact</span><span class="sxs-lookup"><span data-stu-id="110a2-720">Contact</span></span> | <span data-ttu-id="110a2-721">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="110a2-721">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="110a2-722">String</span><span class="sxs-lookup"><span data-stu-id="110a2-722">String</span></span> | <span data-ttu-id="110a2-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="110a2-723">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="110a2-724">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="110a2-724">MeetingSuggestion</span></span> | <span data-ttu-id="110a2-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="110a2-725">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="110a2-726">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="110a2-726">PhoneNumber</span></span> | <span data-ttu-id="110a2-727">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="110a2-727">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="110a2-728">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="110a2-728">TaskSuggestion</span></span> | <span data-ttu-id="110a2-729">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="110a2-729">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="110a2-730">String</span><span class="sxs-lookup"><span data-stu-id="110a2-730">String</span></span> | <span data-ttu-id="110a2-731">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="110a2-731">**Restricted**</span></span> |

<span data-ttu-id="110a2-732">Тип: Array. < (строка |[ Contact](/javascript/api/outlook/office.contact)|[](/javascript/api/outlook/office.meetingsuggestion)свойства meetingsuggestion|[](/javascript/api/outlook/office.phonenumber)PhoneNumber|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)? View = Outlook-JS-1,1) ></span><span class="sxs-lookup"><span data-stu-id="110a2-732">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.1)></span></span>


##### <a name="example"></a><span data-ttu-id="110a2-733">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-733">Example</span></span>

<span data-ttu-id="110a2-734">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-734">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```javascript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="110a2-735">getFilteredEntitiesByName (имя) → (Nullable) {Array. < (строка |[ Contact](/javascript/api/outlook/office.contact)|[](/javascript/api/outlook/office.meetingsuggestion)свойства meetingsuggestion|[](/javascript/api/outlook/office.phonenumber)PhoneNumber|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)? View = Outlook-JS-1,1) >}</span><span class="sxs-lookup"><span data-stu-id="110a2-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.1)>}</span></span>

<span data-ttu-id="110a2-736">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="110a2-736">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-737">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="110a2-737">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="110a2-738">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="110a2-738">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="110a2-739">Параметры</span><span class="sxs-lookup"><span data-stu-id="110a2-739">Parameters</span></span>

|<span data-ttu-id="110a2-740">Имя</span><span class="sxs-lookup"><span data-stu-id="110a2-740">Name</span></span>| <span data-ttu-id="110a2-741">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-741">Type</span></span>| <span data-ttu-id="110a2-742">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-742">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="110a2-743">String</span><span class="sxs-lookup"><span data-stu-id="110a2-743">String</span></span>|<span data-ttu-id="110a2-744">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="110a2-744">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="110a2-745">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-745">Requirements</span></span>

|<span data-ttu-id="110a2-746">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-746">Requirement</span></span>| <span data-ttu-id="110a2-747">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-748">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-749">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-749">1.0</span></span>|
|[<span data-ttu-id="110a2-750">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-751">ReadItem</span></span>|
|[<span data-ttu-id="110a2-752">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-753">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="110a2-754">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="110a2-754">Returns:</span></span>

<span data-ttu-id="110a2-p146">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="110a2-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="110a2-757">Тип: Array. < (строка |[ Contact](/javascript/api/outlook/office.contact)|[](/javascript/api/outlook/office.meetingsuggestion)свойства meetingsuggestion|[](/javascript/api/outlook/office.phonenumber)PhoneNumber|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)? View = Outlook-JS-1,1) ></span><span class="sxs-lookup"><span data-stu-id="110a2-757">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.1)></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="110a2-758">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="110a2-758">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="110a2-759">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="110a2-759">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-760">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="110a2-760">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="110a2-p147">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="110a2-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="110a2-764">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="110a2-764">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="110a2-765">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="110a2-765">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="110a2-p148">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="110a2-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="110a2-768">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-768">Requirements</span></span>

|<span data-ttu-id="110a2-769">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-769">Requirement</span></span>| <span data-ttu-id="110a2-770">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-771">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-772">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-772">1.0</span></span>|
|[<span data-ttu-id="110a2-773">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-773">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-774">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-774">ReadItem</span></span>|
|[<span data-ttu-id="110a2-775">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-775">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-776">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="110a2-777">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="110a2-777">Returns:</span></span>

<span data-ttu-id="110a2-p149">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="110a2-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="110a2-780">Тип:</span><span class="sxs-lookup"><span data-stu-id="110a2-780">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="110a2-781">Object</span><span class="sxs-lookup"><span data-stu-id="110a2-781">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="110a2-782">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-782">Example</span></span>

<span data-ttu-id="110a2-783">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="110a2-783">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="110a2-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="110a2-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="110a2-785">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="110a2-785">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="110a2-786">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="110a2-786">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="110a2-787">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="110a2-787">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="110a2-p150">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="110a2-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="110a2-790">Параметры</span><span class="sxs-lookup"><span data-stu-id="110a2-790">Parameters</span></span>

|<span data-ttu-id="110a2-791">Имя</span><span class="sxs-lookup"><span data-stu-id="110a2-791">Name</span></span>| <span data-ttu-id="110a2-792">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-792">Type</span></span>| <span data-ttu-id="110a2-793">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-793">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="110a2-794">String</span><span class="sxs-lookup"><span data-stu-id="110a2-794">String</span></span>|<span data-ttu-id="110a2-795">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="110a2-795">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="110a2-796">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-796">Requirements</span></span>

|<span data-ttu-id="110a2-797">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-797">Requirement</span></span>| <span data-ttu-id="110a2-798">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-799">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-800">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-800">1.0</span></span>|
|[<span data-ttu-id="110a2-801">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-801">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-802">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-802">ReadItem</span></span>|
|[<span data-ttu-id="110a2-803">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-803">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-804">Чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="110a2-805">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="110a2-805">Returns:</span></span>

<span data-ttu-id="110a2-806">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="110a2-806">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="110a2-807">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="110a2-807">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="110a2-808">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="110a2-808">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="110a2-809">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-809">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="110a2-810">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="110a2-810">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="110a2-811">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="110a2-811">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="110a2-p151">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="110a2-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="110a2-815">Параметры</span><span class="sxs-lookup"><span data-stu-id="110a2-815">Parameters</span></span>

|<span data-ttu-id="110a2-816">Имя</span><span class="sxs-lookup"><span data-stu-id="110a2-816">Name</span></span>| <span data-ttu-id="110a2-817">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-817">Type</span></span>| <span data-ttu-id="110a2-818">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="110a2-818">Attributes</span></span>| <span data-ttu-id="110a2-819">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-819">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="110a2-820">function</span><span class="sxs-lookup"><span data-stu-id="110a2-820">function</span></span>||<span data-ttu-id="110a2-821">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="110a2-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="110a2-822">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="110a2-822">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="110a2-823">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="110a2-823">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="110a2-824">Объект</span><span class="sxs-lookup"><span data-stu-id="110a2-824">Object</span></span>| <span data-ttu-id="110a2-825">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-825">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-826">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="110a2-826">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="110a2-827">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="110a2-827">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="110a2-828">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-828">Requirements</span></span>

|<span data-ttu-id="110a2-829">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-829">Requirement</span></span>| <span data-ttu-id="110a2-830">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-830">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-831">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="110a2-831">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-832">1.0</span><span class="sxs-lookup"><span data-stu-id="110a2-832">1.0</span></span>|
|[<span data-ttu-id="110a2-833">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-833">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-834">ReadItem</span><span class="sxs-lookup"><span data-stu-id="110a2-834">ReadItem</span></span>|
|[<span data-ttu-id="110a2-835">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-835">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-836">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="110a2-836">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-837">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-837">Example</span></span>

<span data-ttu-id="110a2-p154">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="110a2-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="110a2-841">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="110a2-841">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="110a2-842">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="110a2-842">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="110a2-843">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="110a2-843">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="110a2-844">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="110a2-844">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="110a2-845">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="110a2-845">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="110a2-846">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="110a2-846">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="110a2-847">Параметры</span><span class="sxs-lookup"><span data-stu-id="110a2-847">Parameters</span></span>

|<span data-ttu-id="110a2-848">Имя</span><span class="sxs-lookup"><span data-stu-id="110a2-848">Name</span></span>| <span data-ttu-id="110a2-849">Тип</span><span class="sxs-lookup"><span data-stu-id="110a2-849">Type</span></span>| <span data-ttu-id="110a2-850">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="110a2-850">Attributes</span></span>| <span data-ttu-id="110a2-851">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-851">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="110a2-852">String</span><span class="sxs-lookup"><span data-stu-id="110a2-852">String</span></span>||<span data-ttu-id="110a2-853">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="110a2-853">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="110a2-854">Объект</span><span class="sxs-lookup"><span data-stu-id="110a2-854">Object</span></span>| <span data-ttu-id="110a2-855">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-855">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-856">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="110a2-856">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="110a2-857">Объект</span><span class="sxs-lookup"><span data-stu-id="110a2-857">Object</span></span>| <span data-ttu-id="110a2-858">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-858">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-859">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="110a2-859">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="110a2-860">функция</span><span class="sxs-lookup"><span data-stu-id="110a2-860">function</span></span>| <span data-ttu-id="110a2-861">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="110a2-861">&lt;optional&gt;</span></span>|<span data-ttu-id="110a2-862">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="110a2-862">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="110a2-863">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="110a2-863">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="110a2-864">Ошибки</span><span class="sxs-lookup"><span data-stu-id="110a2-864">Errors</span></span>

| <span data-ttu-id="110a2-865">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="110a2-865">Error code</span></span> | <span data-ttu-id="110a2-866">Описание</span><span class="sxs-lookup"><span data-stu-id="110a2-866">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="110a2-867">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="110a2-867">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="110a2-868">Требования</span><span class="sxs-lookup"><span data-stu-id="110a2-868">Requirements</span></span>

|<span data-ttu-id="110a2-869">Требование</span><span class="sxs-lookup"><span data-stu-id="110a2-869">Requirement</span></span>| <span data-ttu-id="110a2-870">Значение</span><span class="sxs-lookup"><span data-stu-id="110a2-870">Value</span></span>|
|---|---|
|[<span data-ttu-id="110a2-871">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="110a2-871">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="110a2-872">1.1</span><span class="sxs-lookup"><span data-stu-id="110a2-872">1.1</span></span>|
|[<span data-ttu-id="110a2-873">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="110a2-873">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="110a2-874">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="110a2-874">ReadWriteItem</span></span>|
|[<span data-ttu-id="110a2-875">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="110a2-875">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="110a2-876">Создание</span><span class="sxs-lookup"><span data-stu-id="110a2-876">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="110a2-877">Пример</span><span class="sxs-lookup"><span data-stu-id="110a2-877">Example</span></span>

<span data-ttu-id="110a2-878">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="110a2-878">The following code removes an attachment with an identifier of '0'.</span></span>

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```
