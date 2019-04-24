---
title: Office. Context. Mailbox. Item — набор требований 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8e411ac1ce58dd59ad3bfc6590a310289bbe686d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450271"
---
# <a name="item"></a><span data-ttu-id="5ce1c-102">item</span><span class="sxs-lookup"><span data-stu-id="5ce1c-102">item</span></span>

### <span data-ttu-id="5ce1c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="5ce1c-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="5ce1c-107">Requirements</span></span>

|<span data-ttu-id="5ce1c-108">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-108">Requirement</span></span>| <span data-ttu-id="5ce1c-109">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-110">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-111">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-111">1.0</span></span>|
|[<span data-ttu-id="5ce1c-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5ce1c-113">Restricted</span></span>|
|[<span data-ttu-id="5ce1c-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="5ce1c-116">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-116">Example</span></span>

<span data-ttu-id="5ce1c-117">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="5ce1c-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="5ce1c-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="5ce1c-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5ce1c-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="5ce1c-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-122">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="5ce1c-123">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-124">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-124">Type</span></span>

*   <span data-ttu-id="5ce1c-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5ce1c-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-126">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-126">Requirements</span></span>

|<span data-ttu-id="5ce1c-127">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-127">Requirement</span></span>| <span data-ttu-id="5ce1c-128">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-129">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-130">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-130">1.0</span></span>|
|[<span data-ttu-id="5ce1c-131">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-132">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-134">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-135">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-135">Example</span></span>

<span data-ttu-id="5ce1c-136">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5ce1c-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5ce1c-138">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="5ce1c-139">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-140">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-140">Type</span></span>

*   [<span data-ttu-id="5ce1c-141">Получатели</span><span class="sxs-lookup"><span data-stu-id="5ce1c-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="5ce1c-142">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-142">Requirements</span></span>

|<span data-ttu-id="5ce1c-143">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-143">Requirement</span></span>| <span data-ttu-id="5ce1c-144">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-145">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-146">1.1</span><span class="sxs-lookup"><span data-stu-id="5ce1c-146">1.1</span></span>|
|[<span data-ttu-id="5ce1c-147">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-148">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-149">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-150">Создание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-151">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="5ce1c-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="5ce1c-153">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-154">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-154">Type</span></span>

*   [<span data-ttu-id="5ce1c-155">Body</span><span class="sxs-lookup"><span data-stu-id="5ce1c-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="5ce1c-156">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-156">Requirements</span></span>

|<span data-ttu-id="5ce1c-157">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-157">Requirement</span></span>| <span data-ttu-id="5ce1c-158">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-160">1.1</span><span class="sxs-lookup"><span data-stu-id="5ce1c-160">1.1</span></span>|
|[<span data-ttu-id="5ce1c-161">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-162">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-165">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-165">Example</span></span>

<span data-ttu-id="5ce1c-166">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="5ce1c-167">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5ce1c-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5ce1c-169">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="5ce1c-170">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5ce1c-171">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5ce1c-171">Read mode</span></span>

<span data-ttu-id="5ce1c-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="5ce1c-174">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5ce1c-174">Compose mode</span></span>

<span data-ttu-id="5ce1c-175">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5ce1c-176">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-176">Type</span></span>

*   <span data-ttu-id="5ce1c-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-178">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-178">Requirements</span></span>

|<span data-ttu-id="5ce1c-179">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-179">Requirement</span></span>| <span data-ttu-id="5ce1c-180">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-181">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-182">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-182">1.0</span></span>|
|[<span data-ttu-id="5ce1c-183">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-184">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="5ce1c-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="5ce1c-188">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="5ce1c-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="5ce1c-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-193">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-193">Type</span></span>

*   <span data-ttu-id="5ce1c-194">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-195">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-195">Requirements</span></span>

|<span data-ttu-id="5ce1c-196">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-196">Requirement</span></span>| <span data-ttu-id="5ce1c-197">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-198">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-199">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-199">1.0</span></span>|
|[<span data-ttu-id="5ce1c-200">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-201">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-203">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-204">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="5ce1c-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="5ce1c-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="5ce1c-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-208">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-208">Type</span></span>

*   <span data-ttu-id="5ce1c-209">Дата</span><span class="sxs-lookup"><span data-stu-id="5ce1c-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-210">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-210">Requirements</span></span>

|<span data-ttu-id="5ce1c-211">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-211">Requirement</span></span>| <span data-ttu-id="5ce1c-212">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-213">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-214">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-214">1.0</span></span>|
|[<span data-ttu-id="5ce1c-215">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-216">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-217">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-218">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-219">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="5ce1c-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="5ce1c-220">dateTimeModified :Date</span></span>

<span data-ttu-id="5ce1c-p111">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-223">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-224">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-224">Type</span></span>

*   <span data-ttu-id="5ce1c-225">Дата</span><span class="sxs-lookup"><span data-stu-id="5ce1c-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-226">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-226">Requirements</span></span>

|<span data-ttu-id="5ce1c-227">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-227">Requirement</span></span>| <span data-ttu-id="5ce1c-228">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-229">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-230">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-230">1.0</span></span>|
|[<span data-ttu-id="5ce1c-231">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-232">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-233">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-234">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-235">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="5ce1c-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="5ce1c-237">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="5ce1c-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5ce1c-240">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5ce1c-240">Read mode</span></span>

<span data-ttu-id="5ce1c-241">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="5ce1c-242">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5ce1c-242">Compose mode</span></span>

<span data-ttu-id="5ce1c-243">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="5ce1c-244">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="5ce1c-245">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="5ce1c-246">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-246">Type</span></span>

*   <span data-ttu-id="5ce1c-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-248">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-248">Requirements</span></span>

|<span data-ttu-id="5ce1c-249">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-249">Requirement</span></span>| <span data-ttu-id="5ce1c-250">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-251">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-252">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-252">1.0</span></span>|
|[<span data-ttu-id="5ce1c-253">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-254">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-255">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-256">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="5ce1c-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="5ce1c-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="5ce1c-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-262">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-263">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-263">Type</span></span>

*   [<span data-ttu-id="5ce1c-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5ce1c-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5ce1c-265">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-265">Requirements</span></span>

|<span data-ttu-id="5ce1c-266">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-266">Requirement</span></span>| <span data-ttu-id="5ce1c-267">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-268">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-269">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-269">1.0</span></span>|
|[<span data-ttu-id="5ce1c-270">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-271">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-272">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-273">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-274">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="5ce1c-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-275">internetMessageId :String</span></span>

<span data-ttu-id="5ce1c-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-278">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-278">Type</span></span>

*   <span data-ttu-id="5ce1c-279">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-280">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-280">Requirements</span></span>

|<span data-ttu-id="5ce1c-281">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-281">Requirement</span></span>| <span data-ttu-id="5ce1c-282">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-283">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-284">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-284">1.0</span></span>|
|[<span data-ttu-id="5ce1c-285">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-286">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-287">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-288">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-289">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="5ce1c-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-290">itemClass :String</span></span>

<span data-ttu-id="5ce1c-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="5ce1c-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="5ce1c-295">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-295">Type</span></span> | <span data-ttu-id="5ce1c-296">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-296">Description</span></span> | <span data-ttu-id="5ce1c-297">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="5ce1c-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="5ce1c-298">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="5ce1c-298">Appointment items</span></span> | <span data-ttu-id="5ce1c-299">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="5ce1c-300">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="5ce1c-300">Message items</span></span> | <span data-ttu-id="5ce1c-301">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="5ce1c-302">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-303">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-303">Type</span></span>

*   <span data-ttu-id="5ce1c-304">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-305">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-305">Requirements</span></span>

|<span data-ttu-id="5ce1c-306">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-306">Requirement</span></span>| <span data-ttu-id="5ce1c-307">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-308">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-309">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-309">1.0</span></span>|
|[<span data-ttu-id="5ce1c-310">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-311">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-313">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-314">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="5ce1c-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-315">(nullable) itemId :String</span></span>

<span data-ttu-id="5ce1c-p118">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-318">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="5ce1c-319">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="5ce1c-320">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="5ce1c-321">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-322">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-322">Type</span></span>

*   <span data-ttu-id="5ce1c-323">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-324">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-324">Requirements</span></span>

|<span data-ttu-id="5ce1c-325">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-325">Requirement</span></span>| <span data-ttu-id="5ce1c-326">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-327">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-328">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-328">1.0</span></span>|
|[<span data-ttu-id="5ce1c-329">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-330">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-331">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-332">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-333">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-333">Example</span></span>

<span data-ttu-id="5ce1c-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="5ce1c-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="5ce1c-337">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="5ce1c-338">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-339">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-339">Type</span></span>

*   [<span data-ttu-id="5ce1c-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="5ce1c-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="5ce1c-341">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-341">Requirements</span></span>

|<span data-ttu-id="5ce1c-342">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-342">Requirement</span></span>| <span data-ttu-id="5ce1c-343">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-344">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-345">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-345">1.0</span></span>|
|[<span data-ttu-id="5ce1c-346">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-347">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-348">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-349">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-350">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="5ce1c-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="5ce1c-352">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5ce1c-353">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5ce1c-353">Read mode</span></span>

<span data-ttu-id="5ce1c-354">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="5ce1c-355">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5ce1c-355">Compose mode</span></span>

<span data-ttu-id="5ce1c-356">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5ce1c-357">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-357">Type</span></span>

*   <span data-ttu-id="5ce1c-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-359">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-359">Requirements</span></span>

|<span data-ttu-id="5ce1c-360">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-360">Requirement</span></span>| <span data-ttu-id="5ce1c-361">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-363">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-363">1.0</span></span>|
|[<span data-ttu-id="5ce1c-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-365">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="5ce1c-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-368">normalizedSubject :String</span></span>

<span data-ttu-id="5ce1c-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="5ce1c-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-373">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-373">Type</span></span>

*   <span data-ttu-id="5ce1c-374">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-375">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-375">Requirements</span></span>

|<span data-ttu-id="5ce1c-376">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-376">Requirement</span></span>| <span data-ttu-id="5ce1c-377">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-378">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-379">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-379">1.0</span></span>|
|[<span data-ttu-id="5ce1c-380">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-381">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-382">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-383">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-384">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5ce1c-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5ce1c-386">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="5ce1c-387">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5ce1c-388">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5ce1c-388">Read mode</span></span>

<span data-ttu-id="5ce1c-389">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="5ce1c-390">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5ce1c-390">Compose mode</span></span>

<span data-ttu-id="5ce1c-391">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5ce1c-392">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-392">Type</span></span>

*   <span data-ttu-id="5ce1c-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-394">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-394">Requirements</span></span>

|<span data-ttu-id="5ce1c-395">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-395">Requirement</span></span>| <span data-ttu-id="5ce1c-396">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-397">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-398">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-398">1.0</span></span>|
|[<span data-ttu-id="5ce1c-399">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-400">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-401">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-402">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="5ce1c-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="5ce1c-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-406">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-406">Type</span></span>

*   [<span data-ttu-id="5ce1c-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5ce1c-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5ce1c-408">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-408">Requirements</span></span>

|<span data-ttu-id="5ce1c-409">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-409">Requirement</span></span>| <span data-ttu-id="5ce1c-410">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-412">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-412">1.0</span></span>|
|[<span data-ttu-id="5ce1c-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-414">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-417">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5ce1c-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5ce1c-419">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="5ce1c-420">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5ce1c-421">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5ce1c-421">Read mode</span></span>

<span data-ttu-id="5ce1c-422">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="5ce1c-423">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5ce1c-423">Compose mode</span></span>

<span data-ttu-id="5ce1c-424">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="5ce1c-425">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-425">Type</span></span>

*   <span data-ttu-id="5ce1c-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-427">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-427">Requirements</span></span>

|<span data-ttu-id="5ce1c-428">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-428">Requirement</span></span>| <span data-ttu-id="5ce1c-429">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-430">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-431">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-431">1.0</span></span>|
|[<span data-ttu-id="5ce1c-432">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-433">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-434">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-435">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="5ce1c-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="5ce1c-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="5ce1c-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-441">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5ce1c-442">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-442">Type</span></span>

*   [<span data-ttu-id="5ce1c-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5ce1c-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5ce1c-444">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-444">Requirements</span></span>

|<span data-ttu-id="5ce1c-445">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-445">Requirement</span></span>| <span data-ttu-id="5ce1c-446">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-447">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-448">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-448">1.0</span></span>|
|[<span data-ttu-id="5ce1c-449">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-450">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-451">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-452">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-453">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="5ce1c-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="5ce1c-455">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="5ce1c-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5ce1c-458">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5ce1c-458">Read mode</span></span>

<span data-ttu-id="5ce1c-459">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="5ce1c-460">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5ce1c-460">Compose mode</span></span>

<span data-ttu-id="5ce1c-461">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="5ce1c-462">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="5ce1c-463">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="5ce1c-464">Type</span><span class="sxs-lookup"><span data-stu-id="5ce1c-464">Type</span></span>

*   <span data-ttu-id="5ce1c-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-466">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-466">Requirements</span></span>

|<span data-ttu-id="5ce1c-467">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-467">Requirement</span></span>| <span data-ttu-id="5ce1c-468">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-469">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-470">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-470">1.0</span></span>|
|[<span data-ttu-id="5ce1c-471">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-472">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-473">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-474">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="5ce1c-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="5ce1c-476">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="5ce1c-477">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5ce1c-478">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5ce1c-478">Read mode</span></span>

<span data-ttu-id="5ce1c-p130">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="5ce1c-481">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5ce1c-481">Compose mode</span></span>

<span data-ttu-id="5ce1c-482">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="5ce1c-483">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-483">Type</span></span>

*   <span data-ttu-id="5ce1c-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-485">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-485">Requirements</span></span>

|<span data-ttu-id="5ce1c-486">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-486">Requirement</span></span>| <span data-ttu-id="5ce1c-487">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-488">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-489">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-489">1.0</span></span>|
|[<span data-ttu-id="5ce1c-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-491">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-493">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="5ce1c-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="5ce1c-495">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="5ce1c-496">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5ce1c-497">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5ce1c-497">Read mode</span></span>

<span data-ttu-id="5ce1c-p132">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="5ce1c-500">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5ce1c-500">Compose mode</span></span>

<span data-ttu-id="5ce1c-501">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5ce1c-502">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-502">Type</span></span>

*   <span data-ttu-id="5ce1c-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-504">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-504">Requirements</span></span>

|<span data-ttu-id="5ce1c-505">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-505">Requirement</span></span>| <span data-ttu-id="5ce1c-506">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-507">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-508">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-508">1.0</span></span>|
|[<span data-ttu-id="5ce1c-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-510">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="5ce1c-513">Методы</span><span class="sxs-lookup"><span data-stu-id="5ce1c-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="5ce1c-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5ce1c-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5ce1c-515">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="5ce1c-516">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="5ce1c-517">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-518">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-518">Parameters</span></span>

|<span data-ttu-id="5ce1c-519">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-519">Name</span></span>| <span data-ttu-id="5ce1c-520">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-520">Type</span></span>| <span data-ttu-id="5ce1c-521">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5ce1c-521">Attributes</span></span>| <span data-ttu-id="5ce1c-522">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="5ce1c-523">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-523">String</span></span>||<span data-ttu-id="5ce1c-p133">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5ce1c-526">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-526">String</span></span>||<span data-ttu-id="5ce1c-p134">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5ce1c-529">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-529">Object</span></span>| <span data-ttu-id="5ce1c-530">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-530">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-531">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5ce1c-532">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-532">Object</span></span>| <span data-ttu-id="5ce1c-533">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-533">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-534">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5ce1c-535">функция</span><span class="sxs-lookup"><span data-stu-id="5ce1c-535">function</span></span>| <span data-ttu-id="5ce1c-536">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-536">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-537">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5ce1c-538">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5ce1c-539">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5ce1c-540">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5ce1c-540">Errors</span></span>

| <span data-ttu-id="5ce1c-541">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5ce1c-541">Error code</span></span> | <span data-ttu-id="5ce1c-542">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="5ce1c-543">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="5ce1c-544">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5ce1c-545">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5ce1c-546">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-546">Requirements</span></span>

|<span data-ttu-id="5ce1c-547">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-547">Requirement</span></span>| <span data-ttu-id="5ce1c-548">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-549">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-550">1.1</span><span class="sxs-lookup"><span data-stu-id="5ce1c-550">1.1</span></span>|
|[<span data-ttu-id="5ce1c-551">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="5ce1c-553">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-554">Создание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-555">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-555">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="5ce1c-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5ce1c-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5ce1c-557">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="5ce1c-p135">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="5ce1c-561">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="5ce1c-562">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-563">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-563">Parameters</span></span>

|<span data-ttu-id="5ce1c-564">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-564">Name</span></span>| <span data-ttu-id="5ce1c-565">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-565">Type</span></span>| <span data-ttu-id="5ce1c-566">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5ce1c-566">Attributes</span></span>| <span data-ttu-id="5ce1c-567">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="5ce1c-568">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-568">String</span></span>||<span data-ttu-id="5ce1c-p136">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5ce1c-571">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-571">String</span></span>||<span data-ttu-id="5ce1c-572">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-572">The subject of the item to be attached.</span></span> <span data-ttu-id="5ce1c-573">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5ce1c-574">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-574">Object</span></span>| <span data-ttu-id="5ce1c-575">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-575">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-576">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5ce1c-577">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-577">Object</span></span>| <span data-ttu-id="5ce1c-578">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-578">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-579">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5ce1c-580">функция</span><span class="sxs-lookup"><span data-stu-id="5ce1c-580">function</span></span>| <span data-ttu-id="5ce1c-581">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-581">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-582">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5ce1c-583">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5ce1c-584">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5ce1c-585">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5ce1c-585">Errors</span></span>

| <span data-ttu-id="5ce1c-586">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5ce1c-586">Error code</span></span> | <span data-ttu-id="5ce1c-587">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5ce1c-588">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5ce1c-589">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-589">Requirements</span></span>

|<span data-ttu-id="5ce1c-590">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-590">Requirement</span></span>| <span data-ttu-id="5ce1c-591">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-592">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-593">1.1</span><span class="sxs-lookup"><span data-stu-id="5ce1c-593">1.1</span></span>|
|[<span data-ttu-id="5ce1c-594">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="5ce1c-596">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-597">Создание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-598">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-598">Example</span></span>

<span data-ttu-id="5ce1c-599">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="5ce1c-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="5ce1c-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="5ce1c-601">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-602">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5ce1c-603">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5ce1c-604">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="5ce1c-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-608">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-608">Parameters</span></span>

|<span data-ttu-id="5ce1c-609">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-609">Name</span></span>| <span data-ttu-id="5ce1c-610">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-610">Type</span></span>| <span data-ttu-id="5ce1c-611">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="5ce1c-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-612">String &#124; Object</span></span>| |<span data-ttu-id="5ce1c-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5ce1c-615">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5ce1c-615">**OR**</span></span><br/><span data-ttu-id="5ce1c-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5ce1c-618">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-618">String</span></span> | <span data-ttu-id="5ce1c-619">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-619">&lt;optional&gt;</span></span> | <span data-ttu-id="5ce1c-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5ce1c-622">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-622">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5ce1c-623">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-623">&lt;optional&gt;</span></span> | <span data-ttu-id="5ce1c-624">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-624">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5ce1c-625">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-625">String</span></span> | | <span data-ttu-id="5ce1c-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5ce1c-628">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-628">String</span></span> | | <span data-ttu-id="5ce1c-629">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-629">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5ce1c-630">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-630">String</span></span> | | <span data-ttu-id="5ce1c-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5ce1c-633">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-633">String</span></span> | | <span data-ttu-id="5ce1c-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5ce1c-637">function</span><span class="sxs-lookup"><span data-stu-id="5ce1c-637">function</span></span> | <span data-ttu-id="5ce1c-638">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-638">&lt;optional&gt;</span></span> | <span data-ttu-id="5ce1c-639">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5ce1c-640">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-640">Requirements</span></span>

|<span data-ttu-id="5ce1c-641">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-641">Requirement</span></span>| <span data-ttu-id="5ce1c-642">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-643">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-644">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-644">1.0</span></span>|
|[<span data-ttu-id="5ce1c-645">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-646">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-646">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-647">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-648">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-648">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5ce1c-649">Примеры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-649">Examples</span></span>

<span data-ttu-id="5ce1c-650">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-650">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="5ce1c-651">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-651">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="5ce1c-652">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-652">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5ce1c-653">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-653">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5ce1c-654">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-654">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5ce1c-655">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-655">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="5ce1c-656">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="5ce1c-656">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="5ce1c-657">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-657">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-658">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-658">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5ce1c-659">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-659">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5ce1c-660">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-660">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="5ce1c-p145">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-664">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-664">Parameters</span></span>

|<span data-ttu-id="5ce1c-665">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-665">Name</span></span>| <span data-ttu-id="5ce1c-666">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-666">Type</span></span>| <span data-ttu-id="5ce1c-667">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-667">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="5ce1c-668">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-668">String &#124; Object</span></span>| | <span data-ttu-id="5ce1c-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5ce1c-671">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5ce1c-671">**OR**</span></span><br/><span data-ttu-id="5ce1c-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5ce1c-674">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-674">String</span></span> | <span data-ttu-id="5ce1c-675">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-675">&lt;optional&gt;</span></span> | <span data-ttu-id="5ce1c-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5ce1c-678">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-678">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5ce1c-679">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-679">&lt;optional&gt;</span></span> | <span data-ttu-id="5ce1c-680">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-680">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5ce1c-681">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-681">String</span></span> | | <span data-ttu-id="5ce1c-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5ce1c-684">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-684">String</span></span> | | <span data-ttu-id="5ce1c-685">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-685">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5ce1c-686">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-686">String</span></span> | | <span data-ttu-id="5ce1c-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5ce1c-689">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-689">String</span></span> | | <span data-ttu-id="5ce1c-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5ce1c-693">function</span><span class="sxs-lookup"><span data-stu-id="5ce1c-693">function</span></span> | <span data-ttu-id="5ce1c-694">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-694">&lt;optional&gt;</span></span> | <span data-ttu-id="5ce1c-695">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-695">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5ce1c-696">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-696">Requirements</span></span>

|<span data-ttu-id="5ce1c-697">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-697">Requirement</span></span>| <span data-ttu-id="5ce1c-698">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-699">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-700">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-700">1.0</span></span>|
|[<span data-ttu-id="5ce1c-701">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-702">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-703">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-704">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-704">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5ce1c-705">Примеры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-705">Examples</span></span>

<span data-ttu-id="5ce1c-706">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-706">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="5ce1c-707">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-707">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="5ce1c-708">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-708">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5ce1c-709">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-709">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5ce1c-710">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-710">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5ce1c-711">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-711">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="5ce1c-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="5ce1c-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="5ce1c-713">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-713">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-714">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-714">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-715">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-715">Requirements</span></span>

|<span data-ttu-id="5ce1c-716">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-716">Requirement</span></span>| <span data-ttu-id="5ce1c-717">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-717">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-718">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-718">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-719">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-719">1.0</span></span>|
|[<span data-ttu-id="5ce1c-720">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-720">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-721">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-721">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-722">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-722">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-723">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-723">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5ce1c-724">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5ce1c-724">Returns:</span></span>

<span data-ttu-id="5ce1c-725">Тип: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-725">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="5ce1c-726">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-726">Example</span></span>

<span data-ttu-id="5ce1c-727">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-727">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="5ce1c-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5ce1c-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5ce1c-729">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-729">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-730">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-730">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-731">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-731">Parameters</span></span>

|<span data-ttu-id="5ce1c-732">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-732">Name</span></span>| <span data-ttu-id="5ce1c-733">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-733">Type</span></span>| <span data-ttu-id="5ce1c-734">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-734">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="5ce1c-735">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="5ce1c-735">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="5ce1c-736">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-736">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5ce1c-737">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-737">Requirements</span></span>

|<span data-ttu-id="5ce1c-738">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-738">Requirement</span></span>| <span data-ttu-id="5ce1c-739">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-740">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-741">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-741">1.0</span></span>|
|[<span data-ttu-id="5ce1c-742">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-743">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5ce1c-743">Restricted</span></span>|
|[<span data-ttu-id="5ce1c-744">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-745">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-745">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5ce1c-746">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5ce1c-746">Returns:</span></span>

<span data-ttu-id="5ce1c-747">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-747">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="5ce1c-748">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-748">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="5ce1c-749">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-749">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="5ce1c-750">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-750">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="5ce1c-751">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="5ce1c-751">Value of `entityType`</span></span> | <span data-ttu-id="5ce1c-752">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="5ce1c-752">Type of objects in returned array</span></span> | <span data-ttu-id="5ce1c-753">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-753">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="5ce1c-754">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-754">String</span></span> | <span data-ttu-id="5ce1c-755">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5ce1c-755">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="5ce1c-756">Contact</span><span class="sxs-lookup"><span data-stu-id="5ce1c-756">Contact</span></span> | <span data-ttu-id="5ce1c-757">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5ce1c-757">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="5ce1c-758">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-758">String</span></span> | <span data-ttu-id="5ce1c-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5ce1c-759">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="5ce1c-760">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="5ce1c-760">MeetingSuggestion</span></span> | <span data-ttu-id="5ce1c-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5ce1c-761">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="5ce1c-762">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="5ce1c-762">PhoneNumber</span></span> | <span data-ttu-id="5ce1c-763">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5ce1c-763">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="5ce1c-764">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="5ce1c-764">TaskSuggestion</span></span> | <span data-ttu-id="5ce1c-765">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5ce1c-765">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="5ce1c-766">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-766">String</span></span> | <span data-ttu-id="5ce1c-767">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5ce1c-767">**Restricted**</span></span> |

<span data-ttu-id="5ce1c-768">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5ce1c-768">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="5ce1c-769">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-769">Example</span></span>

<span data-ttu-id="5ce1c-770">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-770">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="5ce1c-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5ce1c-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5ce1c-772">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-772">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-773">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-773">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5ce1c-774">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-774">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-775">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-775">Parameters</span></span>

|<span data-ttu-id="5ce1c-776">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-776">Name</span></span>| <span data-ttu-id="5ce1c-777">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-777">Type</span></span>| <span data-ttu-id="5ce1c-778">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-778">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5ce1c-779">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-779">String</span></span>|<span data-ttu-id="5ce1c-780">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-780">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5ce1c-781">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-781">Requirements</span></span>

|<span data-ttu-id="5ce1c-782">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-782">Requirement</span></span>| <span data-ttu-id="5ce1c-783">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-783">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-784">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-784">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-785">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-785">1.0</span></span>|
|[<span data-ttu-id="5ce1c-786">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-786">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-787">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-787">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-788">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-788">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-789">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-789">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5ce1c-790">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5ce1c-790">Returns:</span></span>

<span data-ttu-id="5ce1c-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="5ce1c-793">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5ce1c-793">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="5ce1c-794">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="5ce1c-794">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="5ce1c-795">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-795">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-796">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-796">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5ce1c-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="5ce1c-800">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-800">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="5ce1c-801">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-801">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="5ce1c-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5ce1c-804">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-804">Requirements</span></span>

|<span data-ttu-id="5ce1c-805">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-805">Requirement</span></span>| <span data-ttu-id="5ce1c-806">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-807">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-808">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-808">1.0</span></span>|
|[<span data-ttu-id="5ce1c-809">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-809">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-810">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-810">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-811">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-811">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-812">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-812">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5ce1c-813">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5ce1c-813">Returns:</span></span>

<span data-ttu-id="5ce1c-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="5ce1c-816">Тип:</span><span class="sxs-lookup"><span data-stu-id="5ce1c-816">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5ce1c-817">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-817">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5ce1c-818">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-818">Example</span></span>

<span data-ttu-id="5ce1c-819">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-819">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="5ce1c-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="5ce1c-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="5ce1c-821">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-821">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5ce1c-822">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-822">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5ce1c-823">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-823">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="5ce1c-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-826">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-826">Parameters</span></span>

|<span data-ttu-id="5ce1c-827">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-827">Name</span></span>| <span data-ttu-id="5ce1c-828">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-828">Type</span></span>| <span data-ttu-id="5ce1c-829">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-829">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5ce1c-830">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-830">String</span></span>|<span data-ttu-id="5ce1c-831">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-831">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5ce1c-832">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-832">Requirements</span></span>

|<span data-ttu-id="5ce1c-833">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-833">Requirement</span></span>| <span data-ttu-id="5ce1c-834">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-835">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-836">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-836">1.0</span></span>|
|[<span data-ttu-id="5ce1c-837">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-837">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-838">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-839">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-839">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-840">Чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-840">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5ce1c-841">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5ce1c-841">Returns:</span></span>

<span data-ttu-id="5ce1c-842">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-842">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="5ce1c-843">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5ce1c-843">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5ce1c-844">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="5ce1c-844">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5ce1c-845">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-845">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="5ce1c-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="5ce1c-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="5ce1c-847">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-847">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="5ce1c-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-850">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-850">Parameters</span></span>

|<span data-ttu-id="5ce1c-851">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-851">Name</span></span>| <span data-ttu-id="5ce1c-852">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-852">Type</span></span>| <span data-ttu-id="5ce1c-853">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5ce1c-853">Attributes</span></span>| <span data-ttu-id="5ce1c-854">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-854">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="5ce1c-855">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5ce1c-855">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="5ce1c-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="5ce1c-859">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-859">Object</span></span>| <span data-ttu-id="5ce1c-860">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-860">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-861">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5ce1c-862">Объект</span><span class="sxs-lookup"><span data-stu-id="5ce1c-862">Object</span></span>| <span data-ttu-id="5ce1c-863">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-863">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-864">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5ce1c-865">функция</span><span class="sxs-lookup"><span data-stu-id="5ce1c-865">function</span></span>||<span data-ttu-id="5ce1c-866">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-866">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5ce1c-867">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-867">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="5ce1c-868">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-868">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5ce1c-869">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-869">Requirements</span></span>

|<span data-ttu-id="5ce1c-870">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-870">Requirement</span></span>| <span data-ttu-id="5ce1c-871">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-872">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-873">1.2</span><span class="sxs-lookup"><span data-stu-id="5ce1c-873">1.2</span></span>|
|[<span data-ttu-id="5ce1c-874">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-874">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-875">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-875">ReadWriteItem</span></span>|
|[<span data-ttu-id="5ce1c-876">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-876">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-877">Создание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-877">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="5ce1c-878">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5ce1c-878">Returns:</span></span>

<span data-ttu-id="5ce1c-879">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-879">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="5ce1c-880">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5ce1c-880">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5ce1c-881">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-881">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5ce1c-882">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-882">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="5ce1c-883">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5ce1c-883">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="5ce1c-884">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-884">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="5ce1c-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-888">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-888">Parameters</span></span>

|<span data-ttu-id="5ce1c-889">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-889">Name</span></span>| <span data-ttu-id="5ce1c-890">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-890">Type</span></span>| <span data-ttu-id="5ce1c-891">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5ce1c-891">Attributes</span></span>| <span data-ttu-id="5ce1c-892">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-892">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5ce1c-893">function</span><span class="sxs-lookup"><span data-stu-id="5ce1c-893">function</span></span>||<span data-ttu-id="5ce1c-894">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-894">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5ce1c-895">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-895">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="5ce1c-896">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-896">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="5ce1c-897">Объект</span><span class="sxs-lookup"><span data-stu-id="5ce1c-897">Object</span></span>| <span data-ttu-id="5ce1c-898">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-898">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-899">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-899">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="5ce1c-900">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-900">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5ce1c-901">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-901">Requirements</span></span>

|<span data-ttu-id="5ce1c-902">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-902">Requirement</span></span>| <span data-ttu-id="5ce1c-903">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-904">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-905">1.0</span><span class="sxs-lookup"><span data-stu-id="5ce1c-905">1.0</span></span>|
|[<span data-ttu-id="5ce1c-906">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-906">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-907">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-907">ReadItem</span></span>|
|[<span data-ttu-id="5ce1c-908">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-908">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-909">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-909">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-910">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-910">Example</span></span>

<span data-ttu-id="5ce1c-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="5ce1c-914">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5ce1c-914">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="5ce1c-915">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-915">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="5ce1c-p165">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-920">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-920">Parameters</span></span>

|<span data-ttu-id="5ce1c-921">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-921">Name</span></span>| <span data-ttu-id="5ce1c-922">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-922">Type</span></span>| <span data-ttu-id="5ce1c-923">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5ce1c-923">Attributes</span></span>| <span data-ttu-id="5ce1c-924">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-924">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="5ce1c-925">Строка</span><span class="sxs-lookup"><span data-stu-id="5ce1c-925">String</span></span>||<span data-ttu-id="5ce1c-926">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-926">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="5ce1c-927">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-927">Object</span></span>| <span data-ttu-id="5ce1c-928">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-928">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-929">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-929">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5ce1c-930">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-930">Object</span></span>| <span data-ttu-id="5ce1c-931">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-931">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-932">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-932">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5ce1c-933">функция</span><span class="sxs-lookup"><span data-stu-id="5ce1c-933">function</span></span>| <span data-ttu-id="5ce1c-934">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-934">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-935">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-935">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5ce1c-936">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-936">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5ce1c-937">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5ce1c-937">Errors</span></span>

| <span data-ttu-id="5ce1c-938">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5ce1c-938">Error code</span></span> | <span data-ttu-id="5ce1c-939">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-939">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="5ce1c-940">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-940">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5ce1c-941">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-941">Requirements</span></span>

|<span data-ttu-id="5ce1c-942">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-942">Requirement</span></span>| <span data-ttu-id="5ce1c-943">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-944">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5ce1c-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-945">1.1</span><span class="sxs-lookup"><span data-stu-id="5ce1c-945">1.1</span></span>|
|[<span data-ttu-id="5ce1c-946">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-947">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-947">ReadWriteItem</span></span>|
|[<span data-ttu-id="5ce1c-948">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-949">Создание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-949">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-950">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-950">Example</span></span>

<span data-ttu-id="5ce1c-951">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="5ce1c-951">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="5ce1c-952">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="5ce1c-952">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="5ce1c-953">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-953">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="5ce1c-p166">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5ce1c-957">Параметры</span><span class="sxs-lookup"><span data-stu-id="5ce1c-957">Parameters</span></span>

|<span data-ttu-id="5ce1c-958">Имя</span><span class="sxs-lookup"><span data-stu-id="5ce1c-958">Name</span></span>| <span data-ttu-id="5ce1c-959">Тип</span><span class="sxs-lookup"><span data-stu-id="5ce1c-959">Type</span></span>| <span data-ttu-id="5ce1c-960">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5ce1c-960">Attributes</span></span>| <span data-ttu-id="5ce1c-961">Описание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-961">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="5ce1c-962">String</span><span class="sxs-lookup"><span data-stu-id="5ce1c-962">String</span></span>||<span data-ttu-id="5ce1c-p167">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="5ce1c-966">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-966">Object</span></span>| <span data-ttu-id="5ce1c-967">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-967">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-968">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-968">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5ce1c-969">Object</span><span class="sxs-lookup"><span data-stu-id="5ce1c-969">Object</span></span>| <span data-ttu-id="5ce1c-970">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-970">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-971">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-971">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="5ce1c-972">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5ce1c-972">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="5ce1c-973">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5ce1c-973">&lt;optional&gt;</span></span>|<span data-ttu-id="5ce1c-p168">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="5ce1c-p169">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="5ce1c-978">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="5ce1c-978">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="5ce1c-979">функция</span><span class="sxs-lookup"><span data-stu-id="5ce1c-979">function</span></span>||<span data-ttu-id="5ce1c-980">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5ce1c-980">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5ce1c-981">Требования</span><span class="sxs-lookup"><span data-stu-id="5ce1c-981">Requirements</span></span>

|<span data-ttu-id="5ce1c-982">Требование</span><span class="sxs-lookup"><span data-stu-id="5ce1c-982">Requirement</span></span>| <span data-ttu-id="5ce1c-983">Значение</span><span class="sxs-lookup"><span data-stu-id="5ce1c-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="5ce1c-984">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5ce1c-984">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5ce1c-985">1.2</span><span class="sxs-lookup"><span data-stu-id="5ce1c-985">1.2</span></span>|
|[<span data-ttu-id="5ce1c-986">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5ce1c-986">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5ce1c-987">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5ce1c-987">ReadWriteItem</span></span>|
|[<span data-ttu-id="5ce1c-988">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5ce1c-988">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5ce1c-989">Создание</span><span class="sxs-lookup"><span data-stu-id="5ce1c-989">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5ce1c-990">Пример</span><span class="sxs-lookup"><span data-stu-id="5ce1c-990">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
