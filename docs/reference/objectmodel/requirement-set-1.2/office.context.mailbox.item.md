---
title: Office. Context. Mailbox. Item — набор требований 1,2
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: f0cf0e00a1bbd42b66b0b5e032599c54deb3ac6c
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127438"
---
# <a name="item"></a><span data-ttu-id="04a61-102">item</span><span class="sxs-lookup"><span data-stu-id="04a61-102">item</span></span>

### <span data-ttu-id="04a61-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="04a61-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="04a61-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="04a61-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="04a61-107">Requirements</span></span>

|<span data-ttu-id="04a61-108">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-108">Requirement</span></span>| <span data-ttu-id="04a61-109">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-110">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-111">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-111">1.0</span></span>|
|[<span data-ttu-id="04a61-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="04a61-113">Restricted</span></span>|
|[<span data-ttu-id="04a61-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="04a61-116">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-116">Example</span></span>

<span data-ttu-id="04a61-117">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="04a61-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="04a61-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="04a61-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="04a61-119">вложения: Array. <[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="04a61-119">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="04a61-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-122">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="04a61-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="04a61-123">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="04a61-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-124">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-124">Type</span></span>

*   <span data-ttu-id="04a61-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="04a61-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-126">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-126">Requirements</span></span>

|<span data-ttu-id="04a61-127">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-127">Requirement</span></span>| <span data-ttu-id="04a61-128">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-129">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-130">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-130">1.0</span></span>|
|[<span data-ttu-id="04a61-131">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-132">ReadItem</span></span>|
|[<span data-ttu-id="04a61-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-134">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-135">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-135">Example</span></span>

<span data-ttu-id="04a61-136">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="04a61-137">СК: [получатели](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="04a61-137">bcc: [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="04a61-138">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="04a61-139">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="04a61-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-140">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-140">Type</span></span>

*   [<span data-ttu-id="04a61-141">Получатели</span><span class="sxs-lookup"><span data-stu-id="04a61-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="04a61-142">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-142">Requirements</span></span>

|<span data-ttu-id="04a61-143">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-143">Requirement</span></span>| <span data-ttu-id="04a61-144">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-145">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-146">1.1</span><span class="sxs-lookup"><span data-stu-id="04a61-146">1.1</span></span>|
|[<span data-ttu-id="04a61-147">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-148">ReadItem</span></span>|
|[<span data-ttu-id="04a61-149">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-150">Создание</span><span class="sxs-lookup"><span data-stu-id="04a61-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-151">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="04a61-152">основной текст: [Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="04a61-152">body: [Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="04a61-153">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-154">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-154">Type</span></span>

*   [<span data-ttu-id="04a61-155">Body</span><span class="sxs-lookup"><span data-stu-id="04a61-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="04a61-156">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-156">Requirements</span></span>

|<span data-ttu-id="04a61-157">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-157">Requirement</span></span>| <span data-ttu-id="04a61-158">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-160">1.1</span><span class="sxs-lookup"><span data-stu-id="04a61-160">1.1</span></span>|
|[<span data-ttu-id="04a61-161">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-162">ReadItem</span></span>|
|[<span data-ttu-id="04a61-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-165">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-165">Example</span></span>

<span data-ttu-id="04a61-166">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="04a61-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="04a61-167">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="04a61-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="04a61-168">CC: Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="04a61-168">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="04a61-169">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="04a61-170">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04a61-171">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="04a61-171">Read mode</span></span>

<span data-ttu-id="04a61-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="04a61-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="04a61-174">Режим создания</span><span class="sxs-lookup"><span data-stu-id="04a61-174">Compose mode</span></span>

<span data-ttu-id="04a61-175">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="04a61-176">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-176">Type</span></span>

*   <span data-ttu-id="04a61-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="04a61-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-178">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-178">Requirements</span></span>

|<span data-ttu-id="04a61-179">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-179">Requirement</span></span>| <span data-ttu-id="04a61-180">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-181">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-182">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-182">1.0</span></span>|
|[<span data-ttu-id="04a61-183">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-184">ReadItem</span></span>|
|[<span data-ttu-id="04a61-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-186">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="04a61-187">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="04a61-187">(nullable) conversationId: String</span></span>

<span data-ttu-id="04a61-188">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="04a61-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="04a61-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="04a61-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="04a61-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="04a61-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-193">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-193">Type</span></span>

*   <span data-ttu-id="04a61-194">String</span><span class="sxs-lookup"><span data-stu-id="04a61-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-195">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-195">Requirements</span></span>

|<span data-ttu-id="04a61-196">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-196">Requirement</span></span>| <span data-ttu-id="04a61-197">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-198">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-199">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-199">1.0</span></span>|
|[<span data-ttu-id="04a61-200">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-201">ReadItem</span></span>|
|[<span data-ttu-id="04a61-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-203">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-204">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="04a61-205">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="04a61-205">dateTimeCreated: Date</span></span>

<span data-ttu-id="04a61-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-208">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-208">Type</span></span>

*   <span data-ttu-id="04a61-209">Дата</span><span class="sxs-lookup"><span data-stu-id="04a61-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-210">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-210">Requirements</span></span>

|<span data-ttu-id="04a61-211">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-211">Requirement</span></span>| <span data-ttu-id="04a61-212">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-213">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-214">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-214">1.0</span></span>|
|[<span data-ttu-id="04a61-215">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-216">ReadItem</span></span>|
|[<span data-ttu-id="04a61-217">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-218">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-219">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="04a61-220">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="04a61-220">dateTimeModified: Date</span></span>

<span data-ttu-id="04a61-221">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-221">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="04a61-222">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-223">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="04a61-223">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-224">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-224">Type</span></span>

*   <span data-ttu-id="04a61-225">Дата</span><span class="sxs-lookup"><span data-stu-id="04a61-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-226">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-226">Requirements</span></span>

|<span data-ttu-id="04a61-227">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-227">Requirement</span></span>| <span data-ttu-id="04a61-228">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-229">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-230">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-230">1.0</span></span>|
|[<span data-ttu-id="04a61-231">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-232">ReadItem</span></span>|
|[<span data-ttu-id="04a61-233">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-234">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-235">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="04a61-236">конец: Дата | [Time (время](/javascript/api/outlook_1_2/office.time) )</span><span class="sxs-lookup"><span data-stu-id="04a61-236">end: Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="04a61-237">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="04a61-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="04a61-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04a61-240">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="04a61-240">Read mode</span></span>

<span data-ttu-id="04a61-241">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="04a61-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="04a61-242">Режим создания</span><span class="sxs-lookup"><span data-stu-id="04a61-242">Compose mode</span></span>

<span data-ttu-id="04a61-243">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="04a61-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="04a61-244">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="04a61-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="04a61-245">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="04a61-246">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-246">Type</span></span>

*   <span data-ttu-id="04a61-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="04a61-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-248">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-248">Requirements</span></span>

|<span data-ttu-id="04a61-249">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-249">Requirement</span></span>| <span data-ttu-id="04a61-250">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-251">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-252">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-252">1.0</span></span>|
|[<span data-ttu-id="04a61-253">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-254">ReadItem</span></span>|
|[<span data-ttu-id="04a61-255">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-256">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="04a61-257">от: [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="04a61-257">from: [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="04a61-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="04a61-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="04a61-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-262">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="04a61-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-263">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-263">Type</span></span>

*   [<span data-ttu-id="04a61-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="04a61-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="04a61-265">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-265">Requirements</span></span>

|<span data-ttu-id="04a61-266">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-266">Requirement</span></span>| <span data-ttu-id="04a61-267">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-268">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-269">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-269">1.0</span></span>|
|[<span data-ttu-id="04a61-270">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-271">ReadItem</span></span>|
|[<span data-ttu-id="04a61-272">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-273">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-274">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="04a61-275">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="04a61-275">internetMessageId: String</span></span>

<span data-ttu-id="04a61-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-278">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-278">Type</span></span>

*   <span data-ttu-id="04a61-279">String</span><span class="sxs-lookup"><span data-stu-id="04a61-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-280">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-280">Requirements</span></span>

|<span data-ttu-id="04a61-281">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-281">Requirement</span></span>| <span data-ttu-id="04a61-282">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-283">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-284">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-284">1.0</span></span>|
|[<span data-ttu-id="04a61-285">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-286">ReadItem</span></span>|
|[<span data-ttu-id="04a61-287">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-288">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-289">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="04a61-290">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="04a61-290">itemClass: String</span></span>

<span data-ttu-id="04a61-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="04a61-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="04a61-295">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-295">Type</span></span> | <span data-ttu-id="04a61-296">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-296">Description</span></span> | <span data-ttu-id="04a61-297">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="04a61-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="04a61-298">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="04a61-298">Appointment items</span></span> | <span data-ttu-id="04a61-299">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="04a61-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="04a61-300">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="04a61-300">Message items</span></span> | <span data-ttu-id="04a61-301">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="04a61-302">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="04a61-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-303">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-303">Type</span></span>

*   <span data-ttu-id="04a61-304">String</span><span class="sxs-lookup"><span data-stu-id="04a61-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-305">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-305">Requirements</span></span>

|<span data-ttu-id="04a61-306">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-306">Requirement</span></span>| <span data-ttu-id="04a61-307">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-308">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-309">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-309">1.0</span></span>|
|[<span data-ttu-id="04a61-310">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-311">ReadItem</span></span>|
|[<span data-ttu-id="04a61-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-313">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-314">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="04a61-315">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="04a61-315">(nullable) itemId: String</span></span>

<span data-ttu-id="04a61-316">Получает идентификатор элемента веб-служб Exchange для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-316">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="04a61-317">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-317">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-318">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="04a61-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="04a61-319">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="04a61-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="04a61-320">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="04a61-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="04a61-321">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="04a61-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-322">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-322">Type</span></span>

*   <span data-ttu-id="04a61-323">String</span><span class="sxs-lookup"><span data-stu-id="04a61-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-324">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-324">Requirements</span></span>

|<span data-ttu-id="04a61-325">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-325">Requirement</span></span>| <span data-ttu-id="04a61-326">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-327">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-328">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-328">1.0</span></span>|
|[<span data-ttu-id="04a61-329">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-330">ReadItem</span></span>|
|[<span data-ttu-id="04a61-331">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-332">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-333">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-333">Example</span></span>

<span data-ttu-id="04a61-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="04a61-336">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="04a61-336">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="04a61-337">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="04a61-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="04a61-338">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="04a61-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-339">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-339">Type</span></span>

*   [<span data-ttu-id="04a61-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="04a61-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="04a61-341">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-341">Requirements</span></span>

|<span data-ttu-id="04a61-342">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-342">Requirement</span></span>| <span data-ttu-id="04a61-343">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-344">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-345">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-345">1.0</span></span>|
|[<span data-ttu-id="04a61-346">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-347">ReadItem</span></span>|
|[<span data-ttu-id="04a61-348">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-349">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-350">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="04a61-351">Местоположение: строка | [Location (расположение](/javascript/api/outlook_1_2/office.location) )</span><span class="sxs-lookup"><span data-stu-id="04a61-351">location: String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="04a61-352">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04a61-353">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="04a61-353">Read mode</span></span>

<span data-ttu-id="04a61-354">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="04a61-355">Режим создания</span><span class="sxs-lookup"><span data-stu-id="04a61-355">Compose mode</span></span>

<span data-ttu-id="04a61-356">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="04a61-357">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-357">Type</span></span>

*   <span data-ttu-id="04a61-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="04a61-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-359">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-359">Requirements</span></span>

|<span data-ttu-id="04a61-360">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-360">Requirement</span></span>| <span data-ttu-id="04a61-361">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-363">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-363">1.0</span></span>|
|[<span data-ttu-id="04a61-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-365">ReadItem</span></span>|
|[<span data-ttu-id="04a61-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="04a61-368">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="04a61-368">normalizedSubject: String</span></span>

<span data-ttu-id="04a61-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="04a61-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="04a61-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-373">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-373">Type</span></span>

*   <span data-ttu-id="04a61-374">String</span><span class="sxs-lookup"><span data-stu-id="04a61-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-375">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-375">Requirements</span></span>

|<span data-ttu-id="04a61-376">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-376">Requirement</span></span>| <span data-ttu-id="04a61-377">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-378">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-379">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-379">1.0</span></span>|
|[<span data-ttu-id="04a61-380">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-381">ReadItem</span></span>|
|[<span data-ttu-id="04a61-382">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-383">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-384">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="04a61-385">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="04a61-385">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="04a61-386">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="04a61-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="04a61-387">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04a61-388">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="04a61-388">Read mode</span></span>

<span data-ttu-id="04a61-389">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="04a61-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="04a61-390">Режим создания</span><span class="sxs-lookup"><span data-stu-id="04a61-390">Compose mode</span></span>

<span data-ttu-id="04a61-391">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="04a61-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="04a61-392">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-392">Type</span></span>

*   <span data-ttu-id="04a61-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="04a61-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-394">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-394">Requirements</span></span>

|<span data-ttu-id="04a61-395">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-395">Requirement</span></span>| <span data-ttu-id="04a61-396">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-397">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-398">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-398">1.0</span></span>|
|[<span data-ttu-id="04a61-399">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-400">ReadItem</span></span>|
|[<span data-ttu-id="04a61-401">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-402">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="04a61-403">Организатор: [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="04a61-403">organizer: [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="04a61-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-406">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-406">Type</span></span>

*   [<span data-ttu-id="04a61-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="04a61-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="04a61-408">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-408">Requirements</span></span>

|<span data-ttu-id="04a61-409">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-409">Requirement</span></span>| <span data-ttu-id="04a61-410">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-412">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-412">1.0</span></span>|
|[<span data-ttu-id="04a61-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-414">ReadItem</span></span>|
|[<span data-ttu-id="04a61-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-417">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="04a61-418">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="04a61-418">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="04a61-419">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="04a61-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="04a61-420">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04a61-421">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="04a61-421">Read mode</span></span>

<span data-ttu-id="04a61-422">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="04a61-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="04a61-423">Режим создания</span><span class="sxs-lookup"><span data-stu-id="04a61-423">Compose mode</span></span>

<span data-ttu-id="04a61-424">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="04a61-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="04a61-425">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-425">Type</span></span>

*   <span data-ttu-id="04a61-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="04a61-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-427">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-427">Requirements</span></span>

|<span data-ttu-id="04a61-428">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-428">Requirement</span></span>| <span data-ttu-id="04a61-429">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-430">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-431">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-431">1.0</span></span>|
|[<span data-ttu-id="04a61-432">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-433">ReadItem</span></span>|
|[<span data-ttu-id="04a61-434">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-435">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="04a61-436">Отправитель: [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="04a61-436">sender: [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="04a61-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="04a61-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="04a61-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="04a61-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-441">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="04a61-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="04a61-442">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-442">Type</span></span>

*   [<span data-ttu-id="04a61-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="04a61-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="04a61-444">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-444">Requirements</span></span>

|<span data-ttu-id="04a61-445">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-445">Requirement</span></span>| <span data-ttu-id="04a61-446">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-447">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-448">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-448">1.0</span></span>|
|[<span data-ttu-id="04a61-449">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-450">ReadItem</span></span>|
|[<span data-ttu-id="04a61-451">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-452">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-453">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="04a61-454">Начало: Дата | [Time (время](/javascript/api/outlook_1_2/office.time) )</span><span class="sxs-lookup"><span data-stu-id="04a61-454">start: Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="04a61-455">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="04a61-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="04a61-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04a61-458">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="04a61-458">Read mode</span></span>

<span data-ttu-id="04a61-459">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="04a61-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="04a61-460">Режим создания</span><span class="sxs-lookup"><span data-stu-id="04a61-460">Compose mode</span></span>

<span data-ttu-id="04a61-461">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="04a61-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="04a61-462">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="04a61-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="04a61-463">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="04a61-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="04a61-464">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-464">Type</span></span>

*   <span data-ttu-id="04a61-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="04a61-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-466">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-466">Requirements</span></span>

|<span data-ttu-id="04a61-467">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-467">Requirement</span></span>| <span data-ttu-id="04a61-468">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-469">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-470">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-470">1.0</span></span>|
|[<span data-ttu-id="04a61-471">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-472">ReadItem</span></span>|
|[<span data-ttu-id="04a61-473">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-474">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-474">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="04a61-475">Тема: строка | [Subject (тема](/javascript/api/outlook_1_2/office.subject) )</span><span class="sxs-lookup"><span data-stu-id="04a61-475">subject: String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="04a61-476">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="04a61-477">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="04a61-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04a61-478">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="04a61-478">Read mode</span></span>

<span data-ttu-id="04a61-p130">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="04a61-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="04a61-481">Режим создания</span><span class="sxs-lookup"><span data-stu-id="04a61-481">Compose mode</span></span>

<span data-ttu-id="04a61-482">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="04a61-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="04a61-483">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-483">Type</span></span>

*   <span data-ttu-id="04a61-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="04a61-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-485">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-485">Requirements</span></span>

|<span data-ttu-id="04a61-486">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-486">Requirement</span></span>| <span data-ttu-id="04a61-487">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-488">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-489">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-489">1.0</span></span>|
|[<span data-ttu-id="04a61-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-491">ReadItem</span></span>|
|[<span data-ttu-id="04a61-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-493">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-493">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="04a61-494">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="04a61-494">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="04a61-495">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="04a61-496">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="04a61-497">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="04a61-497">Read mode</span></span>

<span data-ttu-id="04a61-p132">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="04a61-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="04a61-500">Режим создания</span><span class="sxs-lookup"><span data-stu-id="04a61-500">Compose mode</span></span>

<span data-ttu-id="04a61-501">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="04a61-502">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-502">Type</span></span>

*   <span data-ttu-id="04a61-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="04a61-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-504">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-504">Requirements</span></span>

|<span data-ttu-id="04a61-505">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-505">Requirement</span></span>| <span data-ttu-id="04a61-506">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-507">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-508">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-508">1.0</span></span>|
|[<span data-ttu-id="04a61-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-510">ReadItem</span></span>|
|[<span data-ttu-id="04a61-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="04a61-513">Методы</span><span class="sxs-lookup"><span data-stu-id="04a61-513">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="04a61-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="04a61-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="04a61-515">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="04a61-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="04a61-516">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="04a61-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="04a61-517">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="04a61-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-518">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-518">Parameters</span></span>

|<span data-ttu-id="04a61-519">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-519">Name</span></span>| <span data-ttu-id="04a61-520">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-520">Type</span></span>| <span data-ttu-id="04a61-521">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="04a61-521">Attributes</span></span>| <span data-ttu-id="04a61-522">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="04a61-523">String</span><span class="sxs-lookup"><span data-stu-id="04a61-523">String</span></span>||<span data-ttu-id="04a61-p133">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="04a61-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="04a61-526">String</span><span class="sxs-lookup"><span data-stu-id="04a61-526">String</span></span>||<span data-ttu-id="04a61-p134">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="04a61-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="04a61-529">Объект</span><span class="sxs-lookup"><span data-stu-id="04a61-529">Object</span></span>| <span data-ttu-id="04a61-530">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-530">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-531">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="04a61-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04a61-532">Object</span><span class="sxs-lookup"><span data-stu-id="04a61-532">Object</span></span>| <span data-ttu-id="04a61-533">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-533">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-534">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="04a61-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="04a61-535">функция</span><span class="sxs-lookup"><span data-stu-id="04a61-535">function</span></span>| <span data-ttu-id="04a61-536">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-536">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-537">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04a61-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="04a61-538">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="04a61-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="04a61-539">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="04a61-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="04a61-540">Ошибки</span><span class="sxs-lookup"><span data-stu-id="04a61-540">Errors</span></span>

| <span data-ttu-id="04a61-541">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="04a61-541">Error code</span></span> | <span data-ttu-id="04a61-542">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="04a61-543">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="04a61-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="04a61-544">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="04a61-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="04a61-545">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="04a61-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04a61-546">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-546">Requirements</span></span>

|<span data-ttu-id="04a61-547">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-547">Requirement</span></span>| <span data-ttu-id="04a61-548">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-549">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-550">1.1</span><span class="sxs-lookup"><span data-stu-id="04a61-550">1.1</span></span>|
|[<span data-ttu-id="04a61-551">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="04a61-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="04a61-553">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-554">Создание</span><span class="sxs-lookup"><span data-stu-id="04a61-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-555">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-555">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="04a61-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="04a61-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="04a61-557">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="04a61-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="04a61-p135">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="04a61-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="04a61-561">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="04a61-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="04a61-562">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="04a61-562">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-563">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-563">Parameters</span></span>

|<span data-ttu-id="04a61-564">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-564">Name</span></span>| <span data-ttu-id="04a61-565">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-565">Type</span></span>| <span data-ttu-id="04a61-566">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="04a61-566">Attributes</span></span>| <span data-ttu-id="04a61-567">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="04a61-568">String</span><span class="sxs-lookup"><span data-stu-id="04a61-568">String</span></span>||<span data-ttu-id="04a61-p136">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="04a61-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="04a61-571">String</span><span class="sxs-lookup"><span data-stu-id="04a61-571">String</span></span>||<span data-ttu-id="04a61-572">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-572">The subject of the item to be attached.</span></span> <span data-ttu-id="04a61-573">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="04a61-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="04a61-574">Object</span><span class="sxs-lookup"><span data-stu-id="04a61-574">Object</span></span>| <span data-ttu-id="04a61-575">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-575">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-576">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="04a61-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04a61-577">Объект</span><span class="sxs-lookup"><span data-stu-id="04a61-577">Object</span></span>| <span data-ttu-id="04a61-578">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-578">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-579">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="04a61-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="04a61-580">функция</span><span class="sxs-lookup"><span data-stu-id="04a61-580">function</span></span>| <span data-ttu-id="04a61-581">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-581">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-582">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04a61-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="04a61-583">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="04a61-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="04a61-584">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="04a61-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="04a61-585">Ошибки</span><span class="sxs-lookup"><span data-stu-id="04a61-585">Errors</span></span>

| <span data-ttu-id="04a61-586">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="04a61-586">Error code</span></span> | <span data-ttu-id="04a61-587">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="04a61-588">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="04a61-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04a61-589">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-589">Requirements</span></span>

|<span data-ttu-id="04a61-590">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-590">Requirement</span></span>| <span data-ttu-id="04a61-591">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-592">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-593">1.1</span><span class="sxs-lookup"><span data-stu-id="04a61-593">1.1</span></span>|
|[<span data-ttu-id="04a61-594">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="04a61-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="04a61-596">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-597">Создание</span><span class="sxs-lookup"><span data-stu-id="04a61-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-598">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-598">Example</span></span>

<span data-ttu-id="04a61-599">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="04a61-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="04a61-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="04a61-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="04a61-601">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-602">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="04a61-602">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04a61-603">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="04a61-603">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="04a61-604">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="04a61-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="04a61-605">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="04a61-605">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="04a61-606">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="04a61-606">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="04a61-607">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="04a61-607">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-608">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-608">Parameters</span></span>

|<span data-ttu-id="04a61-609">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-609">Name</span></span>| <span data-ttu-id="04a61-610">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-610">Type</span></span>| <span data-ttu-id="04a61-611">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="04a61-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="04a61-612">String &#124; Object</span></span>| |<span data-ttu-id="04a61-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="04a61-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="04a61-615">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="04a61-615">**OR**</span></span><br/><span data-ttu-id="04a61-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="04a61-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="04a61-618">String</span><span class="sxs-lookup"><span data-stu-id="04a61-618">String</span></span> | <span data-ttu-id="04a61-619">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-619">&lt;optional&gt;</span></span> | <span data-ttu-id="04a61-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="04a61-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="04a61-622">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-622">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="04a61-623">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-623">&lt;optional&gt;</span></span> | <span data-ttu-id="04a61-624">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="04a61-624">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="04a61-625">String</span><span class="sxs-lookup"><span data-stu-id="04a61-625">String</span></span> | | <span data-ttu-id="04a61-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="04a61-628">Строка</span><span class="sxs-lookup"><span data-stu-id="04a61-628">String</span></span> | | <span data-ttu-id="04a61-629">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="04a61-629">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="04a61-630">String</span><span class="sxs-lookup"><span data-stu-id="04a61-630">String</span></span> | | <span data-ttu-id="04a61-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="04a61-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="04a61-633">String</span><span class="sxs-lookup"><span data-stu-id="04a61-633">String</span></span> | | <span data-ttu-id="04a61-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="04a61-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="04a61-637">function</span><span class="sxs-lookup"><span data-stu-id="04a61-637">function</span></span> | <span data-ttu-id="04a61-638">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-638">&lt;optional&gt;</span></span> | <span data-ttu-id="04a61-639">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04a61-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04a61-640">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-640">Requirements</span></span>

|<span data-ttu-id="04a61-641">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-641">Requirement</span></span>| <span data-ttu-id="04a61-642">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-643">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-644">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-644">1.0</span></span>|
|[<span data-ttu-id="04a61-645">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-646">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-646">ReadItem</span></span>|
|[<span data-ttu-id="04a61-647">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-648">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-648">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="04a61-649">Примеры</span><span class="sxs-lookup"><span data-stu-id="04a61-649">Examples</span></span>

<span data-ttu-id="04a61-650">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="04a61-650">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="04a61-651">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-651">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="04a61-652">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-652">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="04a61-653">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="04a61-653">Reply with a body and a file attachment.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="04a61-654">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="04a61-654">Reply with a body and an item attachment.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="04a61-655">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="04a61-655">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="04a61-656">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="04a61-656">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="04a61-657">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-657">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-658">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="04a61-658">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04a61-659">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="04a61-659">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="04a61-660">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="04a61-660">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="04a61-661">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="04a61-661">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="04a61-662">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="04a61-662">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="04a61-663">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="04a61-663">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-664">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-664">Parameters</span></span>

|<span data-ttu-id="04a61-665">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-665">Name</span></span>| <span data-ttu-id="04a61-666">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-666">Type</span></span>| <span data-ttu-id="04a61-667">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-667">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="04a61-668">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="04a61-668">String &#124; Object</span></span>| | <span data-ttu-id="04a61-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="04a61-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="04a61-671">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="04a61-671">**OR**</span></span><br/><span data-ttu-id="04a61-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="04a61-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="04a61-674">String</span><span class="sxs-lookup"><span data-stu-id="04a61-674">String</span></span> | <span data-ttu-id="04a61-675">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-675">&lt;optional&gt;</span></span> | <span data-ttu-id="04a61-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="04a61-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="04a61-678">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-678">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="04a61-679">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-679">&lt;optional&gt;</span></span> | <span data-ttu-id="04a61-680">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="04a61-680">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="04a61-681">String</span><span class="sxs-lookup"><span data-stu-id="04a61-681">String</span></span> | | <span data-ttu-id="04a61-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="04a61-684">Строка</span><span class="sxs-lookup"><span data-stu-id="04a61-684">String</span></span> | | <span data-ttu-id="04a61-685">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="04a61-685">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="04a61-686">Строка</span><span class="sxs-lookup"><span data-stu-id="04a61-686">String</span></span> | | <span data-ttu-id="04a61-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="04a61-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="04a61-689">String</span><span class="sxs-lookup"><span data-stu-id="04a61-689">String</span></span> | | <span data-ttu-id="04a61-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="04a61-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="04a61-693">function</span><span class="sxs-lookup"><span data-stu-id="04a61-693">function</span></span> | <span data-ttu-id="04a61-694">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-694">&lt;optional&gt;</span></span> | <span data-ttu-id="04a61-695">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04a61-695">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04a61-696">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-696">Requirements</span></span>

|<span data-ttu-id="04a61-697">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-697">Requirement</span></span>| <span data-ttu-id="04a61-698">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-699">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-700">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-700">1.0</span></span>|
|[<span data-ttu-id="04a61-701">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-702">ReadItem</span></span>|
|[<span data-ttu-id="04a61-703">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-704">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-704">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="04a61-705">Примеры</span><span class="sxs-lookup"><span data-stu-id="04a61-705">Examples</span></span>

<span data-ttu-id="04a61-706">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="04a61-706">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="04a61-707">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-707">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="04a61-708">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-708">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="04a61-709">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="04a61-709">Reply with a body and a file attachment.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="04a61-710">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="04a61-710">Reply with a body and an item attachment.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="04a61-711">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="04a61-711">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="04a61-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="04a61-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="04a61-713">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-713">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-714">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="04a61-714">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-715">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-715">Requirements</span></span>

|<span data-ttu-id="04a61-716">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-716">Requirement</span></span>| <span data-ttu-id="04a61-717">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-717">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-718">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-718">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-719">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-719">1.0</span></span>|
|[<span data-ttu-id="04a61-720">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-720">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-721">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-721">ReadItem</span></span>|
|[<span data-ttu-id="04a61-722">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-722">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-723">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-723">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04a61-724">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="04a61-724">Returns:</span></span>

<span data-ttu-id="04a61-725">Тип: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="04a61-725">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="04a61-726">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-726">Example</span></span>

<span data-ttu-id="04a61-727">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-727">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="04a61-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="04a61-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="04a61-729">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-729">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-730">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="04a61-730">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-731">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-731">Parameters</span></span>

|<span data-ttu-id="04a61-732">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-732">Name</span></span>| <span data-ttu-id="04a61-733">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-733">Type</span></span>| <span data-ttu-id="04a61-734">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-734">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="04a61-735">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="04a61-735">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="04a61-736">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="04a61-736">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04a61-737">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-737">Requirements</span></span>

|<span data-ttu-id="04a61-738">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-738">Requirement</span></span>| <span data-ttu-id="04a61-739">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-740">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-741">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-741">1.0</span></span>|
|[<span data-ttu-id="04a61-742">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-743">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="04a61-743">Restricted</span></span>|
|[<span data-ttu-id="04a61-744">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-745">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-745">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04a61-746">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="04a61-746">Returns:</span></span>

<span data-ttu-id="04a61-747">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="04a61-747">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="04a61-748">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="04a61-748">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="04a61-749">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="04a61-749">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="04a61-750">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="04a61-750">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="04a61-751">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="04a61-751">Value of `entityType`</span></span> | <span data-ttu-id="04a61-752">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="04a61-752">Type of objects in returned array</span></span> | <span data-ttu-id="04a61-753">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-753">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="04a61-754">String</span><span class="sxs-lookup"><span data-stu-id="04a61-754">String</span></span> | <span data-ttu-id="04a61-755">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="04a61-755">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="04a61-756">Contact</span><span class="sxs-lookup"><span data-stu-id="04a61-756">Contact</span></span> | <span data-ttu-id="04a61-757">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="04a61-757">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="04a61-758">String</span><span class="sxs-lookup"><span data-stu-id="04a61-758">String</span></span> | <span data-ttu-id="04a61-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="04a61-759">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="04a61-760">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="04a61-760">MeetingSuggestion</span></span> | <span data-ttu-id="04a61-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="04a61-761">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="04a61-762">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="04a61-762">PhoneNumber</span></span> | <span data-ttu-id="04a61-763">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="04a61-763">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="04a61-764">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="04a61-764">TaskSuggestion</span></span> | <span data-ttu-id="04a61-765">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="04a61-765">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="04a61-766">String</span><span class="sxs-lookup"><span data-stu-id="04a61-766">String</span></span> | <span data-ttu-id="04a61-767">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="04a61-767">**Restricted**</span></span> |

<span data-ttu-id="04a61-768">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="04a61-768">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="04a61-769">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-769">Example</span></span>

<span data-ttu-id="04a61-770">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-770">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="04a61-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="04a61-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="04a61-772">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="04a61-772">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-773">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="04a61-773">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04a61-774">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="04a61-774">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-775">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-775">Parameters</span></span>

|<span data-ttu-id="04a61-776">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-776">Name</span></span>| <span data-ttu-id="04a61-777">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-777">Type</span></span>| <span data-ttu-id="04a61-778">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-778">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="04a61-779">String</span><span class="sxs-lookup"><span data-stu-id="04a61-779">String</span></span>|<span data-ttu-id="04a61-780">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="04a61-780">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04a61-781">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-781">Requirements</span></span>

|<span data-ttu-id="04a61-782">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-782">Requirement</span></span>| <span data-ttu-id="04a61-783">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-783">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-784">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-784">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-785">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-785">1.0</span></span>|
|[<span data-ttu-id="04a61-786">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-786">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-787">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-787">ReadItem</span></span>|
|[<span data-ttu-id="04a61-788">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-788">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-789">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-789">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04a61-790">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="04a61-790">Returns:</span></span>

<span data-ttu-id="04a61-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="04a61-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="04a61-793">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="04a61-793">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="04a61-794">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="04a61-794">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="04a61-795">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="04a61-795">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-796">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="04a61-796">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04a61-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="04a61-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="04a61-800">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="04a61-800">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="04a61-801">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="04a61-801">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="04a61-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="04a61-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="04a61-804">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-804">Requirements</span></span>

|<span data-ttu-id="04a61-805">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-805">Requirement</span></span>| <span data-ttu-id="04a61-806">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-807">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-808">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-808">1.0</span></span>|
|[<span data-ttu-id="04a61-809">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-809">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-810">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-810">ReadItem</span></span>|
|[<span data-ttu-id="04a61-811">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-811">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-812">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-812">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04a61-813">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="04a61-813">Returns:</span></span>

<span data-ttu-id="04a61-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="04a61-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="04a61-816">Тип:</span><span class="sxs-lookup"><span data-stu-id="04a61-816">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="04a61-817">Object</span><span class="sxs-lookup"><span data-stu-id="04a61-817">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="04a61-818">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-818">Example</span></span>

<span data-ttu-id="04a61-819">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="04a61-819">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="04a61-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="04a61-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="04a61-821">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="04a61-821">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="04a61-822">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="04a61-822">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="04a61-823">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="04a61-823">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="04a61-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="04a61-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-826">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-826">Parameters</span></span>

|<span data-ttu-id="04a61-827">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-827">Name</span></span>| <span data-ttu-id="04a61-828">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-828">Type</span></span>| <span data-ttu-id="04a61-829">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-829">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="04a61-830">String</span><span class="sxs-lookup"><span data-stu-id="04a61-830">String</span></span>|<span data-ttu-id="04a61-831">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="04a61-831">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04a61-832">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-832">Requirements</span></span>

|<span data-ttu-id="04a61-833">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-833">Requirement</span></span>| <span data-ttu-id="04a61-834">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-835">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-836">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-836">1.0</span></span>|
|[<span data-ttu-id="04a61-837">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-837">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-838">ReadItem</span></span>|
|[<span data-ttu-id="04a61-839">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-839">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-840">Чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-840">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="04a61-841">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="04a61-841">Returns:</span></span>

<span data-ttu-id="04a61-842">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="04a61-842">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="04a61-843">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="04a61-843">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="04a61-844">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="04a61-844">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="04a61-845">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-845">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="04a61-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="04a61-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="04a61-847">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-847">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="04a61-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="04a61-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-850">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-850">Parameters</span></span>

|<span data-ttu-id="04a61-851">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-851">Name</span></span>| <span data-ttu-id="04a61-852">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-852">Type</span></span>| <span data-ttu-id="04a61-853">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="04a61-853">Attributes</span></span>| <span data-ttu-id="04a61-854">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-854">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="04a61-855">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="04a61-855">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="04a61-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="04a61-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="04a61-859">Объект</span><span class="sxs-lookup"><span data-stu-id="04a61-859">Object</span></span>| <span data-ttu-id="04a61-860">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-860">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-861">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="04a61-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04a61-862">Объект</span><span class="sxs-lookup"><span data-stu-id="04a61-862">Object</span></span>| <span data-ttu-id="04a61-863">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-863">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-864">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="04a61-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="04a61-865">функция</span><span class="sxs-lookup"><span data-stu-id="04a61-865">function</span></span>||<span data-ttu-id="04a61-866">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04a61-866">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="04a61-867">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="04a61-867">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="04a61-868">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="04a61-868">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04a61-869">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-869">Requirements</span></span>

|<span data-ttu-id="04a61-870">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-870">Requirement</span></span>| <span data-ttu-id="04a61-871">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-872">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-873">1.2</span><span class="sxs-lookup"><span data-stu-id="04a61-873">1.2</span></span>|
|[<span data-ttu-id="04a61-874">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-874">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-875">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="04a61-875">ReadWriteItem</span></span>|
|[<span data-ttu-id="04a61-876">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-876">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-877">Создание</span><span class="sxs-lookup"><span data-stu-id="04a61-877">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="04a61-878">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="04a61-878">Returns:</span></span>

<span data-ttu-id="04a61-879">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="04a61-879">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="04a61-880">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="04a61-880">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="04a61-881">String</span><span class="sxs-lookup"><span data-stu-id="04a61-881">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="04a61-882">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-882">Example</span></span>

```javascript
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="04a61-883">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="04a61-883">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="04a61-884">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="04a61-884">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="04a61-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="04a61-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-888">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-888">Parameters</span></span>

|<span data-ttu-id="04a61-889">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-889">Name</span></span>| <span data-ttu-id="04a61-890">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-890">Type</span></span>| <span data-ttu-id="04a61-891">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="04a61-891">Attributes</span></span>| <span data-ttu-id="04a61-892">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-892">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="04a61-893">function</span><span class="sxs-lookup"><span data-stu-id="04a61-893">function</span></span>||<span data-ttu-id="04a61-894">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04a61-894">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="04a61-895">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="04a61-895">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="04a61-896">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="04a61-896">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="04a61-897">Объект</span><span class="sxs-lookup"><span data-stu-id="04a61-897">Object</span></span>| <span data-ttu-id="04a61-898">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-898">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-899">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="04a61-899">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="04a61-900">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="04a61-900">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04a61-901">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-901">Requirements</span></span>

|<span data-ttu-id="04a61-902">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-902">Requirement</span></span>| <span data-ttu-id="04a61-903">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-904">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-905">1.0</span><span class="sxs-lookup"><span data-stu-id="04a61-905">1.0</span></span>|
|[<span data-ttu-id="04a61-906">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-906">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-907">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04a61-907">ReadItem</span></span>|
|[<span data-ttu-id="04a61-908">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-908">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-909">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04a61-909">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-910">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-910">Example</span></span>

<span data-ttu-id="04a61-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="04a61-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="04a61-914">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="04a61-914">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="04a61-915">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="04a61-915">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="04a61-916">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="04a61-916">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="04a61-917">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="04a61-917">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="04a61-918">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="04a61-918">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="04a61-919">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="04a61-919">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-920">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-920">Parameters</span></span>

|<span data-ttu-id="04a61-921">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-921">Name</span></span>| <span data-ttu-id="04a61-922">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-922">Type</span></span>| <span data-ttu-id="04a61-923">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="04a61-923">Attributes</span></span>| <span data-ttu-id="04a61-924">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-924">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="04a61-925">String</span><span class="sxs-lookup"><span data-stu-id="04a61-925">String</span></span>||<span data-ttu-id="04a61-926">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="04a61-926">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="04a61-927">Объект</span><span class="sxs-lookup"><span data-stu-id="04a61-927">Object</span></span>| <span data-ttu-id="04a61-928">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-928">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-929">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="04a61-929">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04a61-930">Объект</span><span class="sxs-lookup"><span data-stu-id="04a61-930">Object</span></span>| <span data-ttu-id="04a61-931">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-931">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-932">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="04a61-932">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="04a61-933">функция</span><span class="sxs-lookup"><span data-stu-id="04a61-933">function</span></span>| <span data-ttu-id="04a61-934">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-934">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-935">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04a61-935">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="04a61-936">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="04a61-936">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="04a61-937">Ошибки</span><span class="sxs-lookup"><span data-stu-id="04a61-937">Errors</span></span>

| <span data-ttu-id="04a61-938">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="04a61-938">Error code</span></span> | <span data-ttu-id="04a61-939">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-939">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="04a61-940">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="04a61-940">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04a61-941">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-941">Requirements</span></span>

|<span data-ttu-id="04a61-942">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-942">Requirement</span></span>| <span data-ttu-id="04a61-943">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-944">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="04a61-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-945">1.1</span><span class="sxs-lookup"><span data-stu-id="04a61-945">1.1</span></span>|
|[<span data-ttu-id="04a61-946">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-947">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="04a61-947">ReadWriteItem</span></span>|
|[<span data-ttu-id="04a61-948">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-949">Создание</span><span class="sxs-lookup"><span data-stu-id="04a61-949">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-950">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-950">Example</span></span>

<span data-ttu-id="04a61-951">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="04a61-951">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="04a61-952">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="04a61-952">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="04a61-953">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="04a61-953">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="04a61-p166">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="04a61-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="04a61-957">Параметры</span><span class="sxs-lookup"><span data-stu-id="04a61-957">Parameters</span></span>

|<span data-ttu-id="04a61-958">Имя</span><span class="sxs-lookup"><span data-stu-id="04a61-958">Name</span></span>| <span data-ttu-id="04a61-959">Тип</span><span class="sxs-lookup"><span data-stu-id="04a61-959">Type</span></span>| <span data-ttu-id="04a61-960">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="04a61-960">Attributes</span></span>| <span data-ttu-id="04a61-961">Описание</span><span class="sxs-lookup"><span data-stu-id="04a61-961">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="04a61-962">String</span><span class="sxs-lookup"><span data-stu-id="04a61-962">String</span></span>||<span data-ttu-id="04a61-p167">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="04a61-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="04a61-966">Object</span><span class="sxs-lookup"><span data-stu-id="04a61-966">Object</span></span>| <span data-ttu-id="04a61-967">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-967">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-968">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="04a61-968">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="04a61-969">Объект</span><span class="sxs-lookup"><span data-stu-id="04a61-969">Object</span></span>| <span data-ttu-id="04a61-970">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-970">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-971">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="04a61-971">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="04a61-972">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="04a61-972">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="04a61-973">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="04a61-973">&lt;optional&gt;</span></span>|<span data-ttu-id="04a61-974">Если `text`текущий стиль применяется в Outlook для веб-клиентов и клиентов для настольных ПК.</span><span class="sxs-lookup"><span data-stu-id="04a61-974">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="04a61-975">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="04a61-975">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="04a61-976">Если `html` и поле поддерживает HTML (тема не используется), текущий стиль применяется в Outlook в Интернете, а в настольных клиентах Outlook применяется стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="04a61-976">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="04a61-977">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="04a61-977">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="04a61-978">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="04a61-978">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="04a61-979">функция</span><span class="sxs-lookup"><span data-stu-id="04a61-979">function</span></span>||<span data-ttu-id="04a61-980">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="04a61-980">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04a61-981">Требования</span><span class="sxs-lookup"><span data-stu-id="04a61-981">Requirements</span></span>

|<span data-ttu-id="04a61-982">Требование</span><span class="sxs-lookup"><span data-stu-id="04a61-982">Requirement</span></span>| <span data-ttu-id="04a61-983">Значение</span><span class="sxs-lookup"><span data-stu-id="04a61-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="04a61-984">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="04a61-984">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04a61-985">1.2</span><span class="sxs-lookup"><span data-stu-id="04a61-985">1.2</span></span>|
|[<span data-ttu-id="04a61-986">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="04a61-986">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04a61-987">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="04a61-987">ReadWriteItem</span></span>|
|[<span data-ttu-id="04a61-988">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04a61-988">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04a61-989">Создание</span><span class="sxs-lookup"><span data-stu-id="04a61-989">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="04a61-990">Пример</span><span class="sxs-lookup"><span data-stu-id="04a61-990">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
