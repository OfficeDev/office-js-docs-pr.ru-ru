---
title: Office. Context. Mailbox. Item — набор требований 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8e411ac1ce58dd59ad3bfc6590a310289bbe686d
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870515"
---
# <a name="item"></a><span data-ttu-id="9297f-102">item</span><span class="sxs-lookup"><span data-stu-id="9297f-102">item</span></span>

### <span data-ttu-id="9297f-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="9297f-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="9297f-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="9297f-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="9297f-107">Requirements</span></span>

|<span data-ttu-id="9297f-108">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-108">Requirement</span></span>| <span data-ttu-id="9297f-109">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-110">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-111">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-111">1.0</span></span>|
|[<span data-ttu-id="9297f-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9297f-113">Restricted</span></span>|
|[<span data-ttu-id="9297f-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="9297f-116">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-116">Example</span></span>

<span data-ttu-id="9297f-117">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9297f-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="9297f-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="9297f-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="9297f-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9297f-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="9297f-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-122">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="9297f-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="9297f-123">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="9297f-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-124">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-124">Type</span></span>

*   <span data-ttu-id="9297f-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9297f-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-126">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-126">Requirements</span></span>

|<span data-ttu-id="9297f-127">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-127">Requirement</span></span>| <span data-ttu-id="9297f-128">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-129">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-130">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-130">1.0</span></span>|
|[<span data-ttu-id="9297f-131">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-132">ReadItem</span></span>|
|[<span data-ttu-id="9297f-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-134">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-135">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-135">Example</span></span>

<span data-ttu-id="9297f-136">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9297f-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9297f-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9297f-138">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="9297f-139">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9297f-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-140">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-140">Type</span></span>

*   [<span data-ttu-id="9297f-141">Получатели</span><span class="sxs-lookup"><span data-stu-id="9297f-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="9297f-142">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-142">Requirements</span></span>

|<span data-ttu-id="9297f-143">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-143">Requirement</span></span>| <span data-ttu-id="9297f-144">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-145">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-146">1.1</span><span class="sxs-lookup"><span data-stu-id="9297f-146">1.1</span></span>|
|[<span data-ttu-id="9297f-147">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-148">ReadItem</span></span>|
|[<span data-ttu-id="9297f-149">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-150">Создание</span><span class="sxs-lookup"><span data-stu-id="9297f-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-151">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="9297f-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="9297f-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="9297f-153">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-154">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-154">Type</span></span>

*   [<span data-ttu-id="9297f-155">Body</span><span class="sxs-lookup"><span data-stu-id="9297f-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="9297f-156">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-156">Requirements</span></span>

|<span data-ttu-id="9297f-157">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-157">Requirement</span></span>| <span data-ttu-id="9297f-158">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-160">1.1</span><span class="sxs-lookup"><span data-stu-id="9297f-160">1.1</span></span>|
|[<span data-ttu-id="9297f-161">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-162">ReadItem</span></span>|
|[<span data-ttu-id="9297f-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-165">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-165">Example</span></span>

<span data-ttu-id="9297f-166">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="9297f-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="9297f-167">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9297f-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9297f-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9297f-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9297f-169">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="9297f-170">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9297f-171">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9297f-171">Read mode</span></span>

<span data-ttu-id="9297f-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="9297f-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="9297f-174">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9297f-174">Compose mode</span></span>

<span data-ttu-id="9297f-175">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9297f-176">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-176">Type</span></span>

*   <span data-ttu-id="9297f-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9297f-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-178">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-178">Requirements</span></span>

|<span data-ttu-id="9297f-179">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-179">Requirement</span></span>| <span data-ttu-id="9297f-180">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-181">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-182">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-182">1.0</span></span>|
|[<span data-ttu-id="9297f-183">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-184">ReadItem</span></span>|
|[<span data-ttu-id="9297f-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="9297f-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="9297f-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="9297f-188">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="9297f-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="9297f-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="9297f-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="9297f-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="9297f-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-193">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-193">Type</span></span>

*   <span data-ttu-id="9297f-194">String</span><span class="sxs-lookup"><span data-stu-id="9297f-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-195">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-195">Requirements</span></span>

|<span data-ttu-id="9297f-196">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-196">Requirement</span></span>| <span data-ttu-id="9297f-197">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-198">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-199">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-199">1.0</span></span>|
|[<span data-ttu-id="9297f-200">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-201">ReadItem</span></span>|
|[<span data-ttu-id="9297f-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-203">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-204">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="9297f-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="9297f-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="9297f-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-208">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-208">Type</span></span>

*   <span data-ttu-id="9297f-209">Дата</span><span class="sxs-lookup"><span data-stu-id="9297f-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-210">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-210">Requirements</span></span>

|<span data-ttu-id="9297f-211">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-211">Requirement</span></span>| <span data-ttu-id="9297f-212">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-213">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-214">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-214">1.0</span></span>|
|[<span data-ttu-id="9297f-215">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-216">ReadItem</span></span>|
|[<span data-ttu-id="9297f-217">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-218">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-219">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="9297f-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="9297f-220">dateTimeModified :Date</span></span>

<span data-ttu-id="9297f-p111">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-223">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9297f-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-224">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-224">Type</span></span>

*   <span data-ttu-id="9297f-225">Дата</span><span class="sxs-lookup"><span data-stu-id="9297f-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-226">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-226">Requirements</span></span>

|<span data-ttu-id="9297f-227">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-227">Requirement</span></span>| <span data-ttu-id="9297f-228">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-229">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-230">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-230">1.0</span></span>|
|[<span data-ttu-id="9297f-231">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-232">ReadItem</span></span>|
|[<span data-ttu-id="9297f-233">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-234">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-235">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="9297f-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="9297f-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="9297f-237">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="9297f-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="9297f-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9297f-240">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9297f-240">Read mode</span></span>

<span data-ttu-id="9297f-241">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="9297f-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="9297f-242">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9297f-242">Compose mode</span></span>

<span data-ttu-id="9297f-243">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="9297f-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="9297f-244">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="9297f-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="9297f-245">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9297f-246">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-246">Type</span></span>

*   <span data-ttu-id="9297f-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="9297f-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-248">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-248">Requirements</span></span>

|<span data-ttu-id="9297f-249">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-249">Requirement</span></span>| <span data-ttu-id="9297f-250">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-251">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-252">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-252">1.0</span></span>|
|[<span data-ttu-id="9297f-253">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-254">ReadItem</span></span>|
|[<span data-ttu-id="9297f-255">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-256">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="9297f-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9297f-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="9297f-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="9297f-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="9297f-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-262">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9297f-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-263">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-263">Type</span></span>

*   [<span data-ttu-id="9297f-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9297f-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9297f-265">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-265">Requirements</span></span>

|<span data-ttu-id="9297f-266">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-266">Requirement</span></span>| <span data-ttu-id="9297f-267">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-268">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-269">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-269">1.0</span></span>|
|[<span data-ttu-id="9297f-270">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-271">ReadItem</span></span>|
|[<span data-ttu-id="9297f-272">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-273">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-274">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="9297f-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="9297f-275">internetMessageId :String</span></span>

<span data-ttu-id="9297f-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-278">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-278">Type</span></span>

*   <span data-ttu-id="9297f-279">String</span><span class="sxs-lookup"><span data-stu-id="9297f-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-280">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-280">Requirements</span></span>

|<span data-ttu-id="9297f-281">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-281">Requirement</span></span>| <span data-ttu-id="9297f-282">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-283">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-284">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-284">1.0</span></span>|
|[<span data-ttu-id="9297f-285">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-286">ReadItem</span></span>|
|[<span data-ttu-id="9297f-287">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-288">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-289">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="9297f-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="9297f-290">itemClass :String</span></span>

<span data-ttu-id="9297f-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="9297f-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="9297f-295">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-295">Type</span></span> | <span data-ttu-id="9297f-296">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-296">Description</span></span> | <span data-ttu-id="9297f-297">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="9297f-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="9297f-298">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="9297f-298">Appointment items</span></span> | <span data-ttu-id="9297f-299">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="9297f-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="9297f-300">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="9297f-300">Message items</span></span> | <span data-ttu-id="9297f-301">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="9297f-302">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="9297f-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-303">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-303">Type</span></span>

*   <span data-ttu-id="9297f-304">String</span><span class="sxs-lookup"><span data-stu-id="9297f-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-305">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-305">Requirements</span></span>

|<span data-ttu-id="9297f-306">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-306">Requirement</span></span>| <span data-ttu-id="9297f-307">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-308">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-309">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-309">1.0</span></span>|
|[<span data-ttu-id="9297f-310">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-311">ReadItem</span></span>|
|[<span data-ttu-id="9297f-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-313">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-314">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="9297f-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="9297f-315">(nullable) itemId :String</span></span>

<span data-ttu-id="9297f-p118">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-318">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="9297f-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9297f-319">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="9297f-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="9297f-320">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="9297f-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="9297f-321">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="9297f-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-322">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-322">Type</span></span>

*   <span data-ttu-id="9297f-323">String</span><span class="sxs-lookup"><span data-stu-id="9297f-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-324">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-324">Requirements</span></span>

|<span data-ttu-id="9297f-325">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-325">Requirement</span></span>| <span data-ttu-id="9297f-326">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-327">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-328">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-328">1.0</span></span>|
|[<span data-ttu-id="9297f-329">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-330">ReadItem</span></span>|
|[<span data-ttu-id="9297f-331">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-332">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-333">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-333">Example</span></span>

<span data-ttu-id="9297f-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="9297f-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="9297f-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="9297f-337">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="9297f-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="9297f-338">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="9297f-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-339">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-339">Type</span></span>

*   [<span data-ttu-id="9297f-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="9297f-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="9297f-341">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-341">Requirements</span></span>

|<span data-ttu-id="9297f-342">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-342">Requirement</span></span>| <span data-ttu-id="9297f-343">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-344">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-345">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-345">1.0</span></span>|
|[<span data-ttu-id="9297f-346">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-347">ReadItem</span></span>|
|[<span data-ttu-id="9297f-348">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-349">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-350">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="9297f-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="9297f-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="9297f-352">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9297f-353">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9297f-353">Read mode</span></span>

<span data-ttu-id="9297f-354">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="9297f-355">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9297f-355">Compose mode</span></span>

<span data-ttu-id="9297f-356">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9297f-357">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-357">Type</span></span>

*   <span data-ttu-id="9297f-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="9297f-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-359">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-359">Requirements</span></span>

|<span data-ttu-id="9297f-360">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-360">Requirement</span></span>| <span data-ttu-id="9297f-361">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-363">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-363">1.0</span></span>|
|[<span data-ttu-id="9297f-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-365">ReadItem</span></span>|
|[<span data-ttu-id="9297f-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="9297f-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="9297f-368">normalizedSubject :String</span></span>

<span data-ttu-id="9297f-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="9297f-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="9297f-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-373">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-373">Type</span></span>

*   <span data-ttu-id="9297f-374">String</span><span class="sxs-lookup"><span data-stu-id="9297f-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-375">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-375">Requirements</span></span>

|<span data-ttu-id="9297f-376">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-376">Requirement</span></span>| <span data-ttu-id="9297f-377">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-378">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-379">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-379">1.0</span></span>|
|[<span data-ttu-id="9297f-380">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-381">ReadItem</span></span>|
|[<span data-ttu-id="9297f-382">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-383">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-384">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9297f-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9297f-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9297f-386">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="9297f-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="9297f-387">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9297f-388">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9297f-388">Read mode</span></span>

<span data-ttu-id="9297f-389">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="9297f-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9297f-390">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9297f-390">Compose mode</span></span>

<span data-ttu-id="9297f-391">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="9297f-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9297f-392">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-392">Type</span></span>

*   <span data-ttu-id="9297f-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9297f-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-394">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-394">Requirements</span></span>

|<span data-ttu-id="9297f-395">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-395">Requirement</span></span>| <span data-ttu-id="9297f-396">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-397">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-398">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-398">1.0</span></span>|
|[<span data-ttu-id="9297f-399">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-400">ReadItem</span></span>|
|[<span data-ttu-id="9297f-401">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-402">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="9297f-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9297f-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="9297f-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-406">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-406">Type</span></span>

*   [<span data-ttu-id="9297f-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9297f-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9297f-408">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-408">Requirements</span></span>

|<span data-ttu-id="9297f-409">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-409">Requirement</span></span>| <span data-ttu-id="9297f-410">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-412">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-412">1.0</span></span>|
|[<span data-ttu-id="9297f-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-414">ReadItem</span></span>|
|[<span data-ttu-id="9297f-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-417">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9297f-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9297f-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9297f-419">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="9297f-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="9297f-420">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9297f-421">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9297f-421">Read mode</span></span>

<span data-ttu-id="9297f-422">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="9297f-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9297f-423">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9297f-423">Compose mode</span></span>

<span data-ttu-id="9297f-424">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="9297f-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="9297f-425">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-425">Type</span></span>

*   <span data-ttu-id="9297f-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9297f-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-427">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-427">Requirements</span></span>

|<span data-ttu-id="9297f-428">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-428">Requirement</span></span>| <span data-ttu-id="9297f-429">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-430">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-431">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-431">1.0</span></span>|
|[<span data-ttu-id="9297f-432">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-433">ReadItem</span></span>|
|[<span data-ttu-id="9297f-434">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-435">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="9297f-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9297f-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="9297f-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9297f-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="9297f-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="9297f-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-441">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9297f-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9297f-442">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-442">Type</span></span>

*   [<span data-ttu-id="9297f-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9297f-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9297f-444">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-444">Requirements</span></span>

|<span data-ttu-id="9297f-445">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-445">Requirement</span></span>| <span data-ttu-id="9297f-446">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-447">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-448">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-448">1.0</span></span>|
|[<span data-ttu-id="9297f-449">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-450">ReadItem</span></span>|
|[<span data-ttu-id="9297f-451">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-452">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-453">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="9297f-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="9297f-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="9297f-455">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="9297f-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="9297f-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9297f-458">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9297f-458">Read mode</span></span>

<span data-ttu-id="9297f-459">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="9297f-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="9297f-460">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9297f-460">Compose mode</span></span>

<span data-ttu-id="9297f-461">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="9297f-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="9297f-462">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="9297f-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="9297f-463">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9297f-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9297f-464">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-464">Type</span></span>

*   <span data-ttu-id="9297f-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="9297f-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-466">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-466">Requirements</span></span>

|<span data-ttu-id="9297f-467">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-467">Requirement</span></span>| <span data-ttu-id="9297f-468">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-469">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-470">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-470">1.0</span></span>|
|[<span data-ttu-id="9297f-471">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-472">ReadItem</span></span>|
|[<span data-ttu-id="9297f-473">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-474">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="9297f-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9297f-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="9297f-476">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="9297f-477">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="9297f-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9297f-478">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9297f-478">Read mode</span></span>

<span data-ttu-id="9297f-p130">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="9297f-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="9297f-481">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9297f-481">Compose mode</span></span>

<span data-ttu-id="9297f-482">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="9297f-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="9297f-483">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-483">Type</span></span>

*   <span data-ttu-id="9297f-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9297f-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-485">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-485">Requirements</span></span>

|<span data-ttu-id="9297f-486">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-486">Requirement</span></span>| <span data-ttu-id="9297f-487">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-488">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-489">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-489">1.0</span></span>|
|[<span data-ttu-id="9297f-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-491">ReadItem</span></span>|
|[<span data-ttu-id="9297f-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-493">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9297f-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9297f-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9297f-495">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="9297f-496">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9297f-497">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9297f-497">Read mode</span></span>

<span data-ttu-id="9297f-p132">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="9297f-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="9297f-500">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9297f-500">Compose mode</span></span>

<span data-ttu-id="9297f-501">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9297f-502">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-502">Type</span></span>

*   <span data-ttu-id="9297f-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9297f-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-504">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-504">Requirements</span></span>

|<span data-ttu-id="9297f-505">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-505">Requirement</span></span>| <span data-ttu-id="9297f-506">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-507">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-508">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-508">1.0</span></span>|
|[<span data-ttu-id="9297f-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-510">ReadItem</span></span>|
|[<span data-ttu-id="9297f-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="9297f-513">Методы</span><span class="sxs-lookup"><span data-stu-id="9297f-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="9297f-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9297f-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9297f-515">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="9297f-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9297f-516">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="9297f-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="9297f-517">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="9297f-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-518">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-518">Parameters</span></span>

|<span data-ttu-id="9297f-519">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-519">Name</span></span>| <span data-ttu-id="9297f-520">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-520">Type</span></span>| <span data-ttu-id="9297f-521">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9297f-521">Attributes</span></span>| <span data-ttu-id="9297f-522">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="9297f-523">String</span><span class="sxs-lookup"><span data-stu-id="9297f-523">String</span></span>||<span data-ttu-id="9297f-p133">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="9297f-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9297f-526">String</span><span class="sxs-lookup"><span data-stu-id="9297f-526">String</span></span>||<span data-ttu-id="9297f-p134">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9297f-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9297f-529">Объект</span><span class="sxs-lookup"><span data-stu-id="9297f-529">Object</span></span>| <span data-ttu-id="9297f-530">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-530">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-531">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9297f-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9297f-532">Object</span><span class="sxs-lookup"><span data-stu-id="9297f-532">Object</span></span>| <span data-ttu-id="9297f-533">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-533">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-534">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9297f-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9297f-535">функция</span><span class="sxs-lookup"><span data-stu-id="9297f-535">function</span></span>| <span data-ttu-id="9297f-536">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-536">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-537">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9297f-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9297f-538">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9297f-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9297f-539">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="9297f-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9297f-540">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9297f-540">Errors</span></span>

| <span data-ttu-id="9297f-541">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9297f-541">Error code</span></span> | <span data-ttu-id="9297f-542">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="9297f-543">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="9297f-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="9297f-544">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="9297f-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9297f-545">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="9297f-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9297f-546">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-546">Requirements</span></span>

|<span data-ttu-id="9297f-547">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-547">Requirement</span></span>| <span data-ttu-id="9297f-548">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-549">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-550">1.1</span><span class="sxs-lookup"><span data-stu-id="9297f-550">1.1</span></span>|
|[<span data-ttu-id="9297f-551">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9297f-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="9297f-553">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-554">Создание</span><span class="sxs-lookup"><span data-stu-id="9297f-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-555">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-555">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="9297f-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9297f-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9297f-557">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="9297f-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="9297f-p135">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9297f-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="9297f-561">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="9297f-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="9297f-562">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="9297f-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-563">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-563">Parameters</span></span>

|<span data-ttu-id="9297f-564">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-564">Name</span></span>| <span data-ttu-id="9297f-565">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-565">Type</span></span>| <span data-ttu-id="9297f-566">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9297f-566">Attributes</span></span>| <span data-ttu-id="9297f-567">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="9297f-568">String</span><span class="sxs-lookup"><span data-stu-id="9297f-568">String</span></span>||<span data-ttu-id="9297f-p136">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="9297f-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9297f-571">String</span><span class="sxs-lookup"><span data-stu-id="9297f-571">String</span></span>||<span data-ttu-id="9297f-572">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-572">The subject of the item to be attached.</span></span> <span data-ttu-id="9297f-573">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9297f-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9297f-574">Object</span><span class="sxs-lookup"><span data-stu-id="9297f-574">Object</span></span>| <span data-ttu-id="9297f-575">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-575">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-576">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9297f-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9297f-577">Объект</span><span class="sxs-lookup"><span data-stu-id="9297f-577">Object</span></span>| <span data-ttu-id="9297f-578">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-578">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-579">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9297f-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9297f-580">функция</span><span class="sxs-lookup"><span data-stu-id="9297f-580">function</span></span>| <span data-ttu-id="9297f-581">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-581">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-582">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9297f-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9297f-583">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9297f-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9297f-584">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="9297f-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9297f-585">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9297f-585">Errors</span></span>

| <span data-ttu-id="9297f-586">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9297f-586">Error code</span></span> | <span data-ttu-id="9297f-587">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9297f-588">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="9297f-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9297f-589">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-589">Requirements</span></span>

|<span data-ttu-id="9297f-590">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-590">Requirement</span></span>| <span data-ttu-id="9297f-591">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-592">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-593">1.1</span><span class="sxs-lookup"><span data-stu-id="9297f-593">1.1</span></span>|
|[<span data-ttu-id="9297f-594">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9297f-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="9297f-596">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-597">Создание</span><span class="sxs-lookup"><span data-stu-id="9297f-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-598">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-598">Example</span></span>

<span data-ttu-id="9297f-599">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="9297f-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="9297f-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9297f-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="9297f-601">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-602">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9297f-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9297f-603">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="9297f-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9297f-604">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="9297f-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="9297f-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="9297f-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-608">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-608">Parameters</span></span>

|<span data-ttu-id="9297f-609">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-609">Name</span></span>| <span data-ttu-id="9297f-610">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-610">Type</span></span>| <span data-ttu-id="9297f-611">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9297f-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9297f-612">String &#124; Object</span></span>| |<span data-ttu-id="9297f-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9297f-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9297f-615">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="9297f-615">**OR**</span></span><br/><span data-ttu-id="9297f-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="9297f-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9297f-618">String</span><span class="sxs-lookup"><span data-stu-id="9297f-618">String</span></span> | <span data-ttu-id="9297f-619">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-619">&lt;optional&gt;</span></span> | <span data-ttu-id="9297f-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9297f-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="9297f-622">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-622">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9297f-623">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-623">&lt;optional&gt;</span></span> | <span data-ttu-id="9297f-624">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="9297f-624">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="9297f-625">String</span><span class="sxs-lookup"><span data-stu-id="9297f-625">String</span></span> | | <span data-ttu-id="9297f-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="9297f-628">Строка</span><span class="sxs-lookup"><span data-stu-id="9297f-628">String</span></span> | | <span data-ttu-id="9297f-629">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9297f-629">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="9297f-630">String</span><span class="sxs-lookup"><span data-stu-id="9297f-630">String</span></span> | | <span data-ttu-id="9297f-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="9297f-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="9297f-633">String</span><span class="sxs-lookup"><span data-stu-id="9297f-633">String</span></span> | | <span data-ttu-id="9297f-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="9297f-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="9297f-637">function</span><span class="sxs-lookup"><span data-stu-id="9297f-637">function</span></span> | <span data-ttu-id="9297f-638">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-638">&lt;optional&gt;</span></span> | <span data-ttu-id="9297f-639">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9297f-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9297f-640">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-640">Requirements</span></span>

|<span data-ttu-id="9297f-641">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-641">Requirement</span></span>| <span data-ttu-id="9297f-642">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-643">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-644">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-644">1.0</span></span>|
|[<span data-ttu-id="9297f-645">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-646">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-646">ReadItem</span></span>|
|[<span data-ttu-id="9297f-647">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-648">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-648">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9297f-649">Примеры</span><span class="sxs-lookup"><span data-stu-id="9297f-649">Examples</span></span>

<span data-ttu-id="9297f-650">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="9297f-650">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="9297f-651">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-651">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="9297f-652">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-652">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9297f-653">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="9297f-653">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9297f-654">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="9297f-654">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9297f-655">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="9297f-655">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="9297f-656">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9297f-656">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="9297f-657">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-657">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-658">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9297f-658">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9297f-659">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="9297f-659">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9297f-660">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="9297f-660">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="9297f-p145">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="9297f-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-664">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-664">Parameters</span></span>

|<span data-ttu-id="9297f-665">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-665">Name</span></span>| <span data-ttu-id="9297f-666">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-666">Type</span></span>| <span data-ttu-id="9297f-667">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-667">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9297f-668">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9297f-668">String &#124; Object</span></span>| | <span data-ttu-id="9297f-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9297f-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9297f-671">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="9297f-671">**OR**</span></span><br/><span data-ttu-id="9297f-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="9297f-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9297f-674">String</span><span class="sxs-lookup"><span data-stu-id="9297f-674">String</span></span> | <span data-ttu-id="9297f-675">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-675">&lt;optional&gt;</span></span> | <span data-ttu-id="9297f-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9297f-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="9297f-678">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-678">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9297f-679">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-679">&lt;optional&gt;</span></span> | <span data-ttu-id="9297f-680">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="9297f-680">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="9297f-681">String</span><span class="sxs-lookup"><span data-stu-id="9297f-681">String</span></span> | | <span data-ttu-id="9297f-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="9297f-684">Строка</span><span class="sxs-lookup"><span data-stu-id="9297f-684">String</span></span> | | <span data-ttu-id="9297f-685">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9297f-685">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="9297f-686">Строка</span><span class="sxs-lookup"><span data-stu-id="9297f-686">String</span></span> | | <span data-ttu-id="9297f-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="9297f-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="9297f-689">String</span><span class="sxs-lookup"><span data-stu-id="9297f-689">String</span></span> | | <span data-ttu-id="9297f-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="9297f-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="9297f-693">function</span><span class="sxs-lookup"><span data-stu-id="9297f-693">function</span></span> | <span data-ttu-id="9297f-694">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-694">&lt;optional&gt;</span></span> | <span data-ttu-id="9297f-695">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9297f-695">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9297f-696">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-696">Requirements</span></span>

|<span data-ttu-id="9297f-697">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-697">Requirement</span></span>| <span data-ttu-id="9297f-698">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-699">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-700">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-700">1.0</span></span>|
|[<span data-ttu-id="9297f-701">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-702">ReadItem</span></span>|
|[<span data-ttu-id="9297f-703">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-704">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-704">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9297f-705">Примеры</span><span class="sxs-lookup"><span data-stu-id="9297f-705">Examples</span></span>

<span data-ttu-id="9297f-706">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="9297f-706">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="9297f-707">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-707">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="9297f-708">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-708">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9297f-709">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="9297f-709">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9297f-710">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="9297f-710">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9297f-711">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="9297f-711">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="9297f-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="9297f-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="9297f-713">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-713">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-714">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9297f-714">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-715">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-715">Requirements</span></span>

|<span data-ttu-id="9297f-716">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-716">Requirement</span></span>| <span data-ttu-id="9297f-717">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-717">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-718">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-718">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-719">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-719">1.0</span></span>|
|[<span data-ttu-id="9297f-720">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-720">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-721">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-721">ReadItem</span></span>|
|[<span data-ttu-id="9297f-722">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-722">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-723">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-723">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9297f-724">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9297f-724">Returns:</span></span>

<span data-ttu-id="9297f-725">Тип: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="9297f-725">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="9297f-726">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-726">Example</span></span>

<span data-ttu-id="9297f-727">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-727">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="9297f-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9297f-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9297f-729">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-729">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-730">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9297f-730">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-731">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-731">Parameters</span></span>

|<span data-ttu-id="9297f-732">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-732">Name</span></span>| <span data-ttu-id="9297f-733">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-733">Type</span></span>| <span data-ttu-id="9297f-734">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-734">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="9297f-735">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="9297f-735">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="9297f-736">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="9297f-736">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9297f-737">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-737">Requirements</span></span>

|<span data-ttu-id="9297f-738">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-738">Requirement</span></span>| <span data-ttu-id="9297f-739">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-740">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-741">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-741">1.0</span></span>|
|[<span data-ttu-id="9297f-742">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-743">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9297f-743">Restricted</span></span>|
|[<span data-ttu-id="9297f-744">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-745">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-745">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9297f-746">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9297f-746">Returns:</span></span>

<span data-ttu-id="9297f-747">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="9297f-747">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="9297f-748">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="9297f-748">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="9297f-749">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="9297f-749">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="9297f-750">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="9297f-750">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="9297f-751">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="9297f-751">Value of `entityType`</span></span> | <span data-ttu-id="9297f-752">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="9297f-752">Type of objects in returned array</span></span> | <span data-ttu-id="9297f-753">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-753">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="9297f-754">String</span><span class="sxs-lookup"><span data-stu-id="9297f-754">String</span></span> | <span data-ttu-id="9297f-755">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9297f-755">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="9297f-756">Contact</span><span class="sxs-lookup"><span data-stu-id="9297f-756">Contact</span></span> | <span data-ttu-id="9297f-757">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9297f-757">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="9297f-758">String</span><span class="sxs-lookup"><span data-stu-id="9297f-758">String</span></span> | <span data-ttu-id="9297f-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9297f-759">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="9297f-760">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9297f-760">MeetingSuggestion</span></span> | <span data-ttu-id="9297f-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9297f-761">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="9297f-762">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9297f-762">PhoneNumber</span></span> | <span data-ttu-id="9297f-763">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9297f-763">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="9297f-764">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9297f-764">TaskSuggestion</span></span> | <span data-ttu-id="9297f-765">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9297f-765">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="9297f-766">String</span><span class="sxs-lookup"><span data-stu-id="9297f-766">String</span></span> | <span data-ttu-id="9297f-767">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9297f-767">**Restricted**</span></span> |

<span data-ttu-id="9297f-768">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9297f-768">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="9297f-769">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-769">Example</span></span>

<span data-ttu-id="9297f-770">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-770">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="9297f-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9297f-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9297f-772">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9297f-772">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-773">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9297f-773">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9297f-774">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="9297f-774">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-775">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-775">Parameters</span></span>

|<span data-ttu-id="9297f-776">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-776">Name</span></span>| <span data-ttu-id="9297f-777">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-777">Type</span></span>| <span data-ttu-id="9297f-778">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-778">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9297f-779">String</span><span class="sxs-lookup"><span data-stu-id="9297f-779">String</span></span>|<span data-ttu-id="9297f-780">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="9297f-780">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9297f-781">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-781">Requirements</span></span>

|<span data-ttu-id="9297f-782">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-782">Requirement</span></span>| <span data-ttu-id="9297f-783">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-783">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-784">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-784">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-785">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-785">1.0</span></span>|
|[<span data-ttu-id="9297f-786">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-786">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-787">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-787">ReadItem</span></span>|
|[<span data-ttu-id="9297f-788">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-788">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-789">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-789">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9297f-790">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9297f-790">Returns:</span></span>

<span data-ttu-id="9297f-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="9297f-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="9297f-793">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9297f-793">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="9297f-794">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9297f-794">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="9297f-795">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9297f-795">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-796">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9297f-796">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9297f-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="9297f-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9297f-800">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="9297f-800">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9297f-801">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="9297f-801">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="9297f-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="9297f-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9297f-804">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-804">Requirements</span></span>

|<span data-ttu-id="9297f-805">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-805">Requirement</span></span>| <span data-ttu-id="9297f-806">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-807">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-808">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-808">1.0</span></span>|
|[<span data-ttu-id="9297f-809">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-809">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-810">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-810">ReadItem</span></span>|
|[<span data-ttu-id="9297f-811">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-811">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-812">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-812">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9297f-813">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9297f-813">Returns:</span></span>

<span data-ttu-id="9297f-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="9297f-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="9297f-816">Тип:</span><span class="sxs-lookup"><span data-stu-id="9297f-816">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9297f-817">Object</span><span class="sxs-lookup"><span data-stu-id="9297f-817">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9297f-818">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-818">Example</span></span>

<span data-ttu-id="9297f-819">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="9297f-819">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="9297f-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="9297f-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="9297f-821">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9297f-821">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9297f-822">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9297f-822">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9297f-823">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="9297f-823">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="9297f-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="9297f-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-826">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-826">Parameters</span></span>

|<span data-ttu-id="9297f-827">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-827">Name</span></span>| <span data-ttu-id="9297f-828">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-828">Type</span></span>| <span data-ttu-id="9297f-829">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-829">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9297f-830">String</span><span class="sxs-lookup"><span data-stu-id="9297f-830">String</span></span>|<span data-ttu-id="9297f-831">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="9297f-831">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9297f-832">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-832">Requirements</span></span>

|<span data-ttu-id="9297f-833">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-833">Requirement</span></span>| <span data-ttu-id="9297f-834">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-835">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-836">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-836">1.0</span></span>|
|[<span data-ttu-id="9297f-837">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-837">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-838">ReadItem</span></span>|
|[<span data-ttu-id="9297f-839">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-839">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-840">Чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-840">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9297f-841">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9297f-841">Returns:</span></span>

<span data-ttu-id="9297f-842">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9297f-842">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="9297f-843">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="9297f-843">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9297f-844">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="9297f-844">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9297f-845">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-845">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="9297f-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="9297f-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="9297f-847">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-847">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="9297f-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="9297f-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-850">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-850">Parameters</span></span>

|<span data-ttu-id="9297f-851">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-851">Name</span></span>| <span data-ttu-id="9297f-852">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-852">Type</span></span>| <span data-ttu-id="9297f-853">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9297f-853">Attributes</span></span>| <span data-ttu-id="9297f-854">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-854">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="9297f-855">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9297f-855">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="9297f-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="9297f-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="9297f-859">Объект</span><span class="sxs-lookup"><span data-stu-id="9297f-859">Object</span></span>| <span data-ttu-id="9297f-860">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-860">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-861">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9297f-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9297f-862">Объект</span><span class="sxs-lookup"><span data-stu-id="9297f-862">Object</span></span>| <span data-ttu-id="9297f-863">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-863">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-864">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9297f-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9297f-865">функция</span><span class="sxs-lookup"><span data-stu-id="9297f-865">function</span></span>||<span data-ttu-id="9297f-866">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9297f-866">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9297f-867">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="9297f-867">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="9297f-868">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="9297f-868">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9297f-869">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-869">Requirements</span></span>

|<span data-ttu-id="9297f-870">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-870">Requirement</span></span>| <span data-ttu-id="9297f-871">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-872">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-873">1.2</span><span class="sxs-lookup"><span data-stu-id="9297f-873">1.2</span></span>|
|[<span data-ttu-id="9297f-874">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-874">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-875">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9297f-875">ReadWriteItem</span></span>|
|[<span data-ttu-id="9297f-876">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-876">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-877">Создание</span><span class="sxs-lookup"><span data-stu-id="9297f-877">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="9297f-878">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9297f-878">Returns:</span></span>

<span data-ttu-id="9297f-879">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="9297f-879">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="9297f-880">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="9297f-880">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9297f-881">String</span><span class="sxs-lookup"><span data-stu-id="9297f-881">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9297f-882">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-882">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="9297f-883">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9297f-883">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="9297f-884">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="9297f-884">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="9297f-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="9297f-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-888">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-888">Parameters</span></span>

|<span data-ttu-id="9297f-889">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-889">Name</span></span>| <span data-ttu-id="9297f-890">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-890">Type</span></span>| <span data-ttu-id="9297f-891">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9297f-891">Attributes</span></span>| <span data-ttu-id="9297f-892">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-892">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9297f-893">function</span><span class="sxs-lookup"><span data-stu-id="9297f-893">function</span></span>||<span data-ttu-id="9297f-894">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9297f-894">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9297f-895">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9297f-895">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9297f-896">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="9297f-896">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="9297f-897">Объект</span><span class="sxs-lookup"><span data-stu-id="9297f-897">Object</span></span>| <span data-ttu-id="9297f-898">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-898">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-899">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9297f-899">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="9297f-900">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9297f-900">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9297f-901">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-901">Requirements</span></span>

|<span data-ttu-id="9297f-902">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-902">Requirement</span></span>| <span data-ttu-id="9297f-903">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-904">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-905">1.0</span><span class="sxs-lookup"><span data-stu-id="9297f-905">1.0</span></span>|
|[<span data-ttu-id="9297f-906">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-906">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-907">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9297f-907">ReadItem</span></span>|
|[<span data-ttu-id="9297f-908">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-908">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-909">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9297f-909">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-910">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-910">Example</span></span>

<span data-ttu-id="9297f-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="9297f-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="9297f-914">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9297f-914">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="9297f-915">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="9297f-915">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="9297f-p165">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="9297f-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-920">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-920">Parameters</span></span>

|<span data-ttu-id="9297f-921">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-921">Name</span></span>| <span data-ttu-id="9297f-922">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-922">Type</span></span>| <span data-ttu-id="9297f-923">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9297f-923">Attributes</span></span>| <span data-ttu-id="9297f-924">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-924">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="9297f-925">String</span><span class="sxs-lookup"><span data-stu-id="9297f-925">String</span></span>||<span data-ttu-id="9297f-926">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="9297f-926">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="9297f-927">Объект</span><span class="sxs-lookup"><span data-stu-id="9297f-927">Object</span></span>| <span data-ttu-id="9297f-928">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-928">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-929">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9297f-929">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9297f-930">Объект</span><span class="sxs-lookup"><span data-stu-id="9297f-930">Object</span></span>| <span data-ttu-id="9297f-931">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-931">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-932">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9297f-932">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9297f-933">функция</span><span class="sxs-lookup"><span data-stu-id="9297f-933">function</span></span>| <span data-ttu-id="9297f-934">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-934">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-935">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9297f-935">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9297f-936">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="9297f-936">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9297f-937">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9297f-937">Errors</span></span>

| <span data-ttu-id="9297f-938">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9297f-938">Error code</span></span> | <span data-ttu-id="9297f-939">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-939">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="9297f-940">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="9297f-940">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9297f-941">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-941">Requirements</span></span>

|<span data-ttu-id="9297f-942">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-942">Requirement</span></span>| <span data-ttu-id="9297f-943">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-944">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9297f-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-945">1.1</span><span class="sxs-lookup"><span data-stu-id="9297f-945">1.1</span></span>|
|[<span data-ttu-id="9297f-946">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-947">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9297f-947">ReadWriteItem</span></span>|
|[<span data-ttu-id="9297f-948">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-949">Создание</span><span class="sxs-lookup"><span data-stu-id="9297f-949">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-950">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-950">Example</span></span>

<span data-ttu-id="9297f-951">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="9297f-951">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="9297f-952">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="9297f-952">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="9297f-953">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="9297f-953">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="9297f-p166">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="9297f-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9297f-957">Параметры</span><span class="sxs-lookup"><span data-stu-id="9297f-957">Parameters</span></span>

|<span data-ttu-id="9297f-958">Имя</span><span class="sxs-lookup"><span data-stu-id="9297f-958">Name</span></span>| <span data-ttu-id="9297f-959">Тип</span><span class="sxs-lookup"><span data-stu-id="9297f-959">Type</span></span>| <span data-ttu-id="9297f-960">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9297f-960">Attributes</span></span>| <span data-ttu-id="9297f-961">Описание</span><span class="sxs-lookup"><span data-stu-id="9297f-961">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="9297f-962">String</span><span class="sxs-lookup"><span data-stu-id="9297f-962">String</span></span>||<span data-ttu-id="9297f-p167">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="9297f-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="9297f-966">Object</span><span class="sxs-lookup"><span data-stu-id="9297f-966">Object</span></span>| <span data-ttu-id="9297f-967">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-967">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-968">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9297f-968">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9297f-969">Объект</span><span class="sxs-lookup"><span data-stu-id="9297f-969">Object</span></span>| <span data-ttu-id="9297f-970">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-970">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-971">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="9297f-971">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="9297f-972">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9297f-972">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="9297f-973">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9297f-973">&lt;optional&gt;</span></span>|<span data-ttu-id="9297f-p168">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="9297f-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="9297f-p169">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="9297f-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="9297f-978">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="9297f-978">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="9297f-979">функция</span><span class="sxs-lookup"><span data-stu-id="9297f-979">function</span></span>||<span data-ttu-id="9297f-980">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9297f-980">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9297f-981">Требования</span><span class="sxs-lookup"><span data-stu-id="9297f-981">Requirements</span></span>

|<span data-ttu-id="9297f-982">Требование</span><span class="sxs-lookup"><span data-stu-id="9297f-982">Requirement</span></span>| <span data-ttu-id="9297f-983">Значение</span><span class="sxs-lookup"><span data-stu-id="9297f-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="9297f-984">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9297f-984">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9297f-985">1.2</span><span class="sxs-lookup"><span data-stu-id="9297f-985">1.2</span></span>|
|[<span data-ttu-id="9297f-986">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9297f-986">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9297f-987">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9297f-987">ReadWriteItem</span></span>|
|[<span data-ttu-id="9297f-988">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9297f-988">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9297f-989">Создание</span><span class="sxs-lookup"><span data-stu-id="9297f-989">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9297f-990">Пример</span><span class="sxs-lookup"><span data-stu-id="9297f-990">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
