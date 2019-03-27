---
title: Office. Context. Mailbox. Item — набор требований 1,4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: cd8e71e39940fcf0de50982ef1cdb6825abb7221
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870508"
---
# <a name="item"></a><span data-ttu-id="195ea-102">item</span><span class="sxs-lookup"><span data-stu-id="195ea-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="195ea-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="195ea-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="195ea-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="195ea-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="195ea-106">Requirements</span></span>

|<span data-ttu-id="195ea-107">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-107">Requirement</span></span>| <span data-ttu-id="195ea-108">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-110">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-110">1.0</span></span>|
|[<span data-ttu-id="195ea-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="195ea-112">Restricted</span></span>|
|[<span data-ttu-id="195ea-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-114">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="195ea-115">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-115">Example</span></span>

<span data-ttu-id="195ea-116">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="195ea-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="195ea-117">Элементы</span><span class="sxs-lookup"><span data-stu-id="195ea-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="195ea-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="195ea-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="195ea-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-121">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="195ea-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="195ea-122">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="195ea-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-123">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-123">Type</span></span>

*   <span data-ttu-id="195ea-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="195ea-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-125">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-125">Requirements</span></span>

|<span data-ttu-id="195ea-126">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-126">Requirement</span></span>| <span data-ttu-id="195ea-127">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-128">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-129">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-129">1.0</span></span>|
|[<span data-ttu-id="195ea-130">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-130">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-131">ReadItem</span></span>|
|[<span data-ttu-id="195ea-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-132">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-133">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-134">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-134">Example</span></span>

<span data-ttu-id="195ea-135">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="195ea-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="195ea-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="195ea-137">Получает объект, который предоставляет методы для получения или обновления строки "СК" (Скрытая копия) сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-137">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="195ea-138">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="195ea-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-139">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-139">Type</span></span>

*   [<span data-ttu-id="195ea-140">Получатели</span><span class="sxs-lookup"><span data-stu-id="195ea-140">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="195ea-141">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-141">Requirements</span></span>

|<span data-ttu-id="195ea-142">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-142">Requirement</span></span>| <span data-ttu-id="195ea-143">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-145">1.1</span><span class="sxs-lookup"><span data-stu-id="195ea-145">1.1</span></span>|
|[<span data-ttu-id="195ea-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-147">ReadItem</span></span>|
|[<span data-ttu-id="195ea-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-149">Создание</span><span class="sxs-lookup"><span data-stu-id="195ea-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-150">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-150">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="195ea-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="195ea-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="195ea-152">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-153">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-153">Type</span></span>

*   [<span data-ttu-id="195ea-154">Body</span><span class="sxs-lookup"><span data-stu-id="195ea-154">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="195ea-155">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-155">Requirements</span></span>

|<span data-ttu-id="195ea-156">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-156">Requirement</span></span>| <span data-ttu-id="195ea-157">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-159">1.1</span><span class="sxs-lookup"><span data-stu-id="195ea-159">1.1</span></span>|
|[<span data-ttu-id="195ea-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-161">ReadItem</span></span>|
|[<span data-ttu-id="195ea-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-164">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-164">Example</span></span>

<span data-ttu-id="195ea-165">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="195ea-165">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="195ea-166">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="195ea-166">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="195ea-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="195ea-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="195ea-168">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-168">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="195ea-169">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-169">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="195ea-170">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="195ea-170">Read mode</span></span>

<span data-ttu-id="195ea-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="195ea-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="195ea-173">Режим создания</span><span class="sxs-lookup"><span data-stu-id="195ea-173">Compose mode</span></span>

<span data-ttu-id="195ea-174">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-174">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="195ea-175">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-175">Type</span></span>

*   <span data-ttu-id="195ea-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="195ea-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-177">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-177">Requirements</span></span>

|<span data-ttu-id="195ea-178">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-178">Requirement</span></span>| <span data-ttu-id="195ea-179">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-180">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-181">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-181">1.0</span></span>|
|[<span data-ttu-id="195ea-182">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-183">ReadItem</span></span>|
|[<span data-ttu-id="195ea-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-185">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="195ea-186">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="195ea-186">(nullable) conversationId :String</span></span>

<span data-ttu-id="195ea-187">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="195ea-187">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="195ea-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="195ea-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="195ea-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="195ea-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-192">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-192">Type</span></span>

*   <span data-ttu-id="195ea-193">String</span><span class="sxs-lookup"><span data-stu-id="195ea-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-194">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-194">Requirements</span></span>

|<span data-ttu-id="195ea-195">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-195">Requirement</span></span>| <span data-ttu-id="195ea-196">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-197">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-198">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-198">1.0</span></span>|
|[<span data-ttu-id="195ea-199">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-199">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-200">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-200">ReadItem</span></span>|
|[<span data-ttu-id="195ea-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-201">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-203">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-203">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="195ea-204">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="195ea-204">dateTimeCreated :Date</span></span>

<span data-ttu-id="195ea-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-207">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-207">Type</span></span>

*   <span data-ttu-id="195ea-208">Дата</span><span class="sxs-lookup"><span data-stu-id="195ea-208">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-209">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-209">Requirements</span></span>

|<span data-ttu-id="195ea-210">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-210">Requirement</span></span>| <span data-ttu-id="195ea-211">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-212">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-213">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-213">1.0</span></span>|
|[<span data-ttu-id="195ea-214">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-215">ReadItem</span></span>|
|[<span data-ttu-id="195ea-216">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-217">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-218">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-218">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="195ea-219">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="195ea-219">dateTimeModified :Date</span></span>

<span data-ttu-id="195ea-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-222">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="195ea-222">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-223">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-223">Type</span></span>

*   <span data-ttu-id="195ea-224">Дата</span><span class="sxs-lookup"><span data-stu-id="195ea-224">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-225">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-225">Requirements</span></span>

|<span data-ttu-id="195ea-226">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-226">Requirement</span></span>| <span data-ttu-id="195ea-227">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-228">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-229">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-229">1.0</span></span>|
|[<span data-ttu-id="195ea-230">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-231">ReadItem</span></span>|
|[<span data-ttu-id="195ea-232">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-233">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-233">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-234">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-234">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="195ea-235">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="195ea-235">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="195ea-236">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-236">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="195ea-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="195ea-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="195ea-239">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="195ea-239">Read mode</span></span>

<span data-ttu-id="195ea-240">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="195ea-240">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="195ea-241">Режим создания</span><span class="sxs-lookup"><span data-stu-id="195ea-241">Compose mode</span></span>

<span data-ttu-id="195ea-242">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="195ea-242">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="195ea-243">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="195ea-243">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="195ea-244">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-244">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="195ea-245">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-245">Type</span></span>

*   <span data-ttu-id="195ea-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="195ea-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-247">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-247">Requirements</span></span>

|<span data-ttu-id="195ea-248">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-248">Requirement</span></span>| <span data-ttu-id="195ea-249">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-250">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-251">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-251">1.0</span></span>|
|[<span data-ttu-id="195ea-252">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-253">ReadItem</span></span>|
|[<span data-ttu-id="195ea-254">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-255">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-255">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="195ea-256">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="195ea-256">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="195ea-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="195ea-p113">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="195ea-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-261">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="195ea-261">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-262">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-262">Type</span></span>

*   [<span data-ttu-id="195ea-263">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="195ea-263">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="195ea-264">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-264">Requirements</span></span>

|<span data-ttu-id="195ea-265">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-265">Requirement</span></span>| <span data-ttu-id="195ea-266">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-267">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-268">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-268">1.0</span></span>|
|[<span data-ttu-id="195ea-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-270">ReadItem</span></span>|
|[<span data-ttu-id="195ea-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-272">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-272">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-273">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-273">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="195ea-274">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="195ea-274">internetMessageId :String</span></span>

<span data-ttu-id="195ea-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-277">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-277">Type</span></span>

*   <span data-ttu-id="195ea-278">String</span><span class="sxs-lookup"><span data-stu-id="195ea-278">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-279">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-279">Requirements</span></span>

|<span data-ttu-id="195ea-280">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-280">Requirement</span></span>| <span data-ttu-id="195ea-281">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-282">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-283">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-283">1.0</span></span>|
|[<span data-ttu-id="195ea-284">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-285">ReadItem</span></span>|
|[<span data-ttu-id="195ea-286">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-287">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-288">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-288">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="195ea-289">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="195ea-289">itemClass :String</span></span>

<span data-ttu-id="195ea-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="195ea-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="195ea-294">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-294">Type</span></span> | <span data-ttu-id="195ea-295">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-295">Description</span></span> | <span data-ttu-id="195ea-296">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="195ea-296">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="195ea-297">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="195ea-297">Appointment items</span></span> | <span data-ttu-id="195ea-298">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="195ea-298">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="195ea-299">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="195ea-299">Message items</span></span> | <span data-ttu-id="195ea-300">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-300">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="195ea-301">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="195ea-301">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-302">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-302">Type</span></span>

*   <span data-ttu-id="195ea-303">String</span><span class="sxs-lookup"><span data-stu-id="195ea-303">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-304">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-304">Requirements</span></span>

|<span data-ttu-id="195ea-305">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-305">Requirement</span></span>| <span data-ttu-id="195ea-306">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-307">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-308">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-308">1.0</span></span>|
|[<span data-ttu-id="195ea-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-310">ReadItem</span></span>|
|[<span data-ttu-id="195ea-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-313">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-313">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="195ea-314">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="195ea-314">(nullable) itemId :String</span></span>

<span data-ttu-id="195ea-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-317">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="195ea-317">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="195ea-318">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="195ea-318">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="195ea-319">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="195ea-319">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="195ea-320">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="195ea-320">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="195ea-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-323">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-323">Type</span></span>

*   <span data-ttu-id="195ea-324">String</span><span class="sxs-lookup"><span data-stu-id="195ea-324">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-325">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-325">Requirements</span></span>

|<span data-ttu-id="195ea-326">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-326">Requirement</span></span>| <span data-ttu-id="195ea-327">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-328">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-329">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-329">1.0</span></span>|
|[<span data-ttu-id="195ea-330">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-331">ReadItem</span></span>|
|[<span data-ttu-id="195ea-332">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-333">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-334">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-334">Example</span></span>

<span data-ttu-id="195ea-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="195ea-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="195ea-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="195ea-338">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="195ea-338">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="195ea-339">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="195ea-339">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-340">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-340">Type</span></span>

*   [<span data-ttu-id="195ea-341">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="195ea-341">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="195ea-342">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-342">Requirements</span></span>

|<span data-ttu-id="195ea-343">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-343">Requirement</span></span>| <span data-ttu-id="195ea-344">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-344">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-345">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-346">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-346">1.0</span></span>|
|[<span data-ttu-id="195ea-347">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-348">ReadItem</span></span>|
|[<span data-ttu-id="195ea-349">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-350">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-350">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-351">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-351">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="195ea-352">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="195ea-352">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="195ea-353">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-353">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="195ea-354">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="195ea-354">Read mode</span></span>

<span data-ttu-id="195ea-355">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-355">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="195ea-356">Режим создания</span><span class="sxs-lookup"><span data-stu-id="195ea-356">Compose mode</span></span>

<span data-ttu-id="195ea-357">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-357">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="195ea-358">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-358">Type</span></span>

*   <span data-ttu-id="195ea-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="195ea-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-360">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-360">Requirements</span></span>

|<span data-ttu-id="195ea-361">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-361">Requirement</span></span>| <span data-ttu-id="195ea-362">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-363">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-364">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-364">1.0</span></span>|
|[<span data-ttu-id="195ea-365">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-366">ReadItem</span></span>|
|[<span data-ttu-id="195ea-367">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-368">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-368">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="195ea-369">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="195ea-369">normalizedSubject :String</span></span>

<span data-ttu-id="195ea-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="195ea-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="195ea-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-374">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-374">Type</span></span>

*   <span data-ttu-id="195ea-375">String</span><span class="sxs-lookup"><span data-stu-id="195ea-375">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-376">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-376">Requirements</span></span>

|<span data-ttu-id="195ea-377">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-377">Requirement</span></span>| <span data-ttu-id="195ea-378">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-378">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-379">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-380">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-380">1.0</span></span>|
|[<span data-ttu-id="195ea-381">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-382">ReadItem</span></span>|
|[<span data-ttu-id="195ea-383">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-384">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-384">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-385">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-385">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="195ea-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="195ea-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="195ea-387">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-387">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-388">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-388">Type</span></span>

*   [<span data-ttu-id="195ea-389">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="195ea-389">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="195ea-390">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-390">Requirements</span></span>

|<span data-ttu-id="195ea-391">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-391">Requirement</span></span>| <span data-ttu-id="195ea-392">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-392">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-393">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-394">1.3</span><span class="sxs-lookup"><span data-stu-id="195ea-394">1.3</span></span>|
|[<span data-ttu-id="195ea-395">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-395">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-396">ReadItem</span></span>|
|[<span data-ttu-id="195ea-397">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-398">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-398">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-399">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-399">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="195ea-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="195ea-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="195ea-401">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="195ea-401">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="195ea-402">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-402">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="195ea-403">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="195ea-403">Read mode</span></span>

<span data-ttu-id="195ea-404">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="195ea-404">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="195ea-405">Режим создания</span><span class="sxs-lookup"><span data-stu-id="195ea-405">Compose mode</span></span>

<span data-ttu-id="195ea-406">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="195ea-406">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="195ea-407">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-407">Type</span></span>

*   <span data-ttu-id="195ea-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="195ea-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-409">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-409">Requirements</span></span>

|<span data-ttu-id="195ea-410">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-410">Requirement</span></span>| <span data-ttu-id="195ea-411">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-412">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-413">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-413">1.0</span></span>|
|[<span data-ttu-id="195ea-414">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-415">ReadItem</span></span>|
|[<span data-ttu-id="195ea-416">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-417">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-417">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="195ea-418">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="195ea-418">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="195ea-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-421">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-421">Type</span></span>

*   [<span data-ttu-id="195ea-422">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="195ea-422">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="195ea-423">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-423">Requirements</span></span>

|<span data-ttu-id="195ea-424">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-424">Requirement</span></span>| <span data-ttu-id="195ea-425">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-426">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-427">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-427">1.0</span></span>|
|[<span data-ttu-id="195ea-428">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-429">ReadItem</span></span>|
|[<span data-ttu-id="195ea-430">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-431">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-432">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-432">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="195ea-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="195ea-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="195ea-434">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="195ea-434">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="195ea-435">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-435">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="195ea-436">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="195ea-436">Read mode</span></span>

<span data-ttu-id="195ea-437">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="195ea-437">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="195ea-438">Режим создания</span><span class="sxs-lookup"><span data-stu-id="195ea-438">Compose mode</span></span>

<span data-ttu-id="195ea-439">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="195ea-439">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="195ea-440">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-440">Type</span></span>

*   <span data-ttu-id="195ea-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="195ea-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-442">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-442">Requirements</span></span>

|<span data-ttu-id="195ea-443">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-443">Requirement</span></span>| <span data-ttu-id="195ea-444">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-445">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-446">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-446">1.0</span></span>|
|[<span data-ttu-id="195ea-447">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-447">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-448">ReadItem</span></span>|
|[<span data-ttu-id="195ea-449">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-449">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-450">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-450">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="195ea-451">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="195ea-451">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="195ea-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="195ea-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="195ea-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="195ea-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-456">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="195ea-456">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="195ea-457">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-457">Type</span></span>

*   [<span data-ttu-id="195ea-458">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="195ea-458">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="195ea-459">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-459">Requirements</span></span>

|<span data-ttu-id="195ea-460">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-460">Requirement</span></span>| <span data-ttu-id="195ea-461">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-462">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-463">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-463">1.0</span></span>|
|[<span data-ttu-id="195ea-464">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-465">ReadItem</span></span>|
|[<span data-ttu-id="195ea-466">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-467">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-468">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-468">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="195ea-469">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="195ea-469">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="195ea-470">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-470">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="195ea-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="195ea-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="195ea-473">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="195ea-473">Read mode</span></span>

<span data-ttu-id="195ea-474">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="195ea-474">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="195ea-475">Режим создания</span><span class="sxs-lookup"><span data-stu-id="195ea-475">Compose mode</span></span>

<span data-ttu-id="195ea-476">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="195ea-476">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="195ea-477">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="195ea-477">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="195ea-478">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="195ea-478">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="195ea-479">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-479">Type</span></span>

*   <span data-ttu-id="195ea-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="195ea-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-481">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-481">Requirements</span></span>

|<span data-ttu-id="195ea-482">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-482">Requirement</span></span>| <span data-ttu-id="195ea-483">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-484">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-485">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-485">1.0</span></span>|
|[<span data-ttu-id="195ea-486">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-487">ReadItem</span></span>|
|[<span data-ttu-id="195ea-488">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-489">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-489">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="195ea-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="195ea-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="195ea-491">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="195ea-492">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="195ea-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="195ea-493">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="195ea-493">Read mode</span></span>

<span data-ttu-id="195ea-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="195ea-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="195ea-496">Режим создания</span><span class="sxs-lookup"><span data-stu-id="195ea-496">Compose mode</span></span>

<span data-ttu-id="195ea-497">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="195ea-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="195ea-498">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-498">Type</span></span>

*   <span data-ttu-id="195ea-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="195ea-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-500">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-500">Requirements</span></span>

|<span data-ttu-id="195ea-501">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-501">Requirement</span></span>| <span data-ttu-id="195ea-502">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-503">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-504">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-504">1.0</span></span>|
|[<span data-ttu-id="195ea-505">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-506">ReadItem</span></span>|
|[<span data-ttu-id="195ea-507">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-508">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-508">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="195ea-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="195ea-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="195ea-510">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="195ea-511">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="195ea-512">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="195ea-512">Read mode</span></span>

<span data-ttu-id="195ea-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="195ea-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="195ea-515">Режим создания</span><span class="sxs-lookup"><span data-stu-id="195ea-515">Compose mode</span></span>

<span data-ttu-id="195ea-516">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="195ea-517">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-517">Type</span></span>

*   <span data-ttu-id="195ea-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="195ea-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-519">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-519">Requirements</span></span>

|<span data-ttu-id="195ea-520">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-520">Requirement</span></span>| <span data-ttu-id="195ea-521">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-522">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-523">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-523">1.0</span></span>|
|[<span data-ttu-id="195ea-524">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-525">ReadItem</span></span>|
|[<span data-ttu-id="195ea-526">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-527">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-527">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="195ea-528">Методы</span><span class="sxs-lookup"><span data-stu-id="195ea-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="195ea-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="195ea-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="195ea-530">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="195ea-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="195ea-531">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="195ea-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="195ea-532">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="195ea-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-533">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-533">Parameters</span></span>

|<span data-ttu-id="195ea-534">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-534">Name</span></span>| <span data-ttu-id="195ea-535">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-535">Type</span></span>| <span data-ttu-id="195ea-536">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="195ea-536">Attributes</span></span>| <span data-ttu-id="195ea-537">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="195ea-538">String</span><span class="sxs-lookup"><span data-stu-id="195ea-538">String</span></span>||<span data-ttu-id="195ea-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="195ea-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="195ea-541">String</span><span class="sxs-lookup"><span data-stu-id="195ea-541">String</span></span>||<span data-ttu-id="195ea-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="195ea-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="195ea-544">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-544">Object</span></span>| <span data-ttu-id="195ea-545">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-545">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-546">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="195ea-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="195ea-547">Object</span><span class="sxs-lookup"><span data-stu-id="195ea-547">Object</span></span>| <span data-ttu-id="195ea-548">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-548">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-549">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="195ea-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="195ea-550">функция</span><span class="sxs-lookup"><span data-stu-id="195ea-550">function</span></span>| <span data-ttu-id="195ea-551">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-551">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-552">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="195ea-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="195ea-553">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="195ea-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="195ea-554">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="195ea-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="195ea-555">Ошибки</span><span class="sxs-lookup"><span data-stu-id="195ea-555">Errors</span></span>

| <span data-ttu-id="195ea-556">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="195ea-556">Error code</span></span> | <span data-ttu-id="195ea-557">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="195ea-558">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="195ea-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="195ea-559">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="195ea-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="195ea-560">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="195ea-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="195ea-561">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-561">Requirements</span></span>

|<span data-ttu-id="195ea-562">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-562">Requirement</span></span>| <span data-ttu-id="195ea-563">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-564">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-565">1.1</span><span class="sxs-lookup"><span data-stu-id="195ea-565">1.1</span></span>|
|[<span data-ttu-id="195ea-566">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="195ea-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="195ea-568">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-569">Создание</span><span class="sxs-lookup"><span data-stu-id="195ea-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-570">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="195ea-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="195ea-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="195ea-572">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="195ea-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="195ea-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="195ea-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="195ea-576">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="195ea-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="195ea-577">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="195ea-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-578">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-578">Parameters</span></span>

|<span data-ttu-id="195ea-579">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-579">Name</span></span>| <span data-ttu-id="195ea-580">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-580">Type</span></span>| <span data-ttu-id="195ea-581">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="195ea-581">Attributes</span></span>| <span data-ttu-id="195ea-582">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="195ea-583">String</span><span class="sxs-lookup"><span data-stu-id="195ea-583">String</span></span>||<span data-ttu-id="195ea-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="195ea-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="195ea-586">String</span><span class="sxs-lookup"><span data-stu-id="195ea-586">String</span></span>||<span data-ttu-id="195ea-587">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-587">The subject of the item to be attached.</span></span> <span data-ttu-id="195ea-588">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="195ea-588">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="195ea-589">Object</span><span class="sxs-lookup"><span data-stu-id="195ea-589">Object</span></span>| <span data-ttu-id="195ea-590">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-590">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-591">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="195ea-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="195ea-592">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-592">Object</span></span>| <span data-ttu-id="195ea-593">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-593">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-594">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="195ea-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="195ea-595">функция</span><span class="sxs-lookup"><span data-stu-id="195ea-595">function</span></span>| <span data-ttu-id="195ea-596">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-596">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-597">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="195ea-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="195ea-598">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="195ea-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="195ea-599">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="195ea-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="195ea-600">Ошибки</span><span class="sxs-lookup"><span data-stu-id="195ea-600">Errors</span></span>

| <span data-ttu-id="195ea-601">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="195ea-601">Error code</span></span> | <span data-ttu-id="195ea-602">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="195ea-603">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="195ea-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="195ea-604">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-604">Requirements</span></span>

|<span data-ttu-id="195ea-605">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-605">Requirement</span></span>| <span data-ttu-id="195ea-606">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-607">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-608">1.1</span><span class="sxs-lookup"><span data-stu-id="195ea-608">1.1</span></span>|
|[<span data-ttu-id="195ea-609">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="195ea-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="195ea-611">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-612">Создание</span><span class="sxs-lookup"><span data-stu-id="195ea-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-613">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-613">Example</span></span>

<span data-ttu-id="195ea-614">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="195ea-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="195ea-615">close()</span><span class="sxs-lookup"><span data-stu-id="195ea-615">close()</span></span>

<span data-ttu-id="195ea-616">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="195ea-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="195ea-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="195ea-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-619">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="195ea-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="195ea-620">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="195ea-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-621">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-621">Requirements</span></span>

|<span data-ttu-id="195ea-622">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-622">Requirement</span></span>| <span data-ttu-id="195ea-623">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-624">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-625">1.3</span><span class="sxs-lookup"><span data-stu-id="195ea-625">1.3</span></span>|
|[<span data-ttu-id="195ea-626">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-626">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-627">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="195ea-627">Restricted</span></span>|
|[<span data-ttu-id="195ea-628">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-628">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-629">Создание</span><span class="sxs-lookup"><span data-stu-id="195ea-629">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="195ea-630">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="195ea-630">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="195ea-631">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-632">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="195ea-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="195ea-633">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="195ea-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="195ea-634">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="195ea-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="195ea-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="195ea-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-638">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-638">Parameters</span></span>

|<span data-ttu-id="195ea-639">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-639">Name</span></span>| <span data-ttu-id="195ea-640">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-640">Type</span></span>| <span data-ttu-id="195ea-641">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="195ea-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="195ea-642">String &#124; Object</span></span>| |<span data-ttu-id="195ea-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="195ea-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="195ea-645">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="195ea-645">**OR**</span></span><br/><span data-ttu-id="195ea-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="195ea-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="195ea-648">String</span><span class="sxs-lookup"><span data-stu-id="195ea-648">String</span></span> | <span data-ttu-id="195ea-649">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-649">&lt;optional&gt;</span></span> | <span data-ttu-id="195ea-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="195ea-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="195ea-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="195ea-653">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-653">&lt;optional&gt;</span></span> | <span data-ttu-id="195ea-654">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="195ea-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="195ea-655">String</span><span class="sxs-lookup"><span data-stu-id="195ea-655">String</span></span> | | <span data-ttu-id="195ea-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="195ea-658">Строка</span><span class="sxs-lookup"><span data-stu-id="195ea-658">String</span></span> | | <span data-ttu-id="195ea-659">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="195ea-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="195ea-660">String</span><span class="sxs-lookup"><span data-stu-id="195ea-660">String</span></span> | | <span data-ttu-id="195ea-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="195ea-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="195ea-663">String</span><span class="sxs-lookup"><span data-stu-id="195ea-663">String</span></span> | | <span data-ttu-id="195ea-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="195ea-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="195ea-667">function</span><span class="sxs-lookup"><span data-stu-id="195ea-667">function</span></span> | <span data-ttu-id="195ea-668">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-668">&lt;optional&gt;</span></span> | <span data-ttu-id="195ea-669">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="195ea-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="195ea-670">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-670">Requirements</span></span>

|<span data-ttu-id="195ea-671">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-671">Requirement</span></span>| <span data-ttu-id="195ea-672">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-673">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-674">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-674">1.0</span></span>|
|[<span data-ttu-id="195ea-675">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-676">ReadItem</span></span>|
|[<span data-ttu-id="195ea-677">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-678">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="195ea-679">Примеры</span><span class="sxs-lookup"><span data-stu-id="195ea-679">Examples</span></span>

<span data-ttu-id="195ea-680">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="195ea-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="195ea-681">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-681">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="195ea-682">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-682">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="195ea-683">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="195ea-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="195ea-684">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="195ea-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="195ea-685">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="195ea-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="195ea-686">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="195ea-686">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="195ea-687">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-688">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="195ea-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="195ea-689">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="195ea-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="195ea-690">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="195ea-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="195ea-p145">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="195ea-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-694">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-694">Parameters</span></span>

|<span data-ttu-id="195ea-695">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-695">Name</span></span>| <span data-ttu-id="195ea-696">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-696">Type</span></span>| <span data-ttu-id="195ea-697">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="195ea-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="195ea-698">String &#124; Object</span></span>| | <span data-ttu-id="195ea-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="195ea-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="195ea-701">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="195ea-701">**OR**</span></span><br/><span data-ttu-id="195ea-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="195ea-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="195ea-704">String</span><span class="sxs-lookup"><span data-stu-id="195ea-704">String</span></span> | <span data-ttu-id="195ea-705">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-705">&lt;optional&gt;</span></span> | <span data-ttu-id="195ea-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="195ea-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="195ea-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="195ea-709">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-709">&lt;optional&gt;</span></span> | <span data-ttu-id="195ea-710">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="195ea-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="195ea-711">String</span><span class="sxs-lookup"><span data-stu-id="195ea-711">String</span></span> | | <span data-ttu-id="195ea-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="195ea-714">Строка</span><span class="sxs-lookup"><span data-stu-id="195ea-714">String</span></span> | | <span data-ttu-id="195ea-715">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="195ea-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="195ea-716">Строка</span><span class="sxs-lookup"><span data-stu-id="195ea-716">String</span></span> | | <span data-ttu-id="195ea-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="195ea-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="195ea-719">String</span><span class="sxs-lookup"><span data-stu-id="195ea-719">String</span></span> | | <span data-ttu-id="195ea-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="195ea-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="195ea-723">function</span><span class="sxs-lookup"><span data-stu-id="195ea-723">function</span></span> | <span data-ttu-id="195ea-724">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-724">&lt;optional&gt;</span></span> | <span data-ttu-id="195ea-725">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="195ea-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="195ea-726">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-726">Requirements</span></span>

|<span data-ttu-id="195ea-727">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-727">Requirement</span></span>| <span data-ttu-id="195ea-728">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-729">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-730">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-730">1.0</span></span>|
|[<span data-ttu-id="195ea-731">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-732">ReadItem</span></span>|
|[<span data-ttu-id="195ea-733">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-734">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="195ea-735">Примеры</span><span class="sxs-lookup"><span data-stu-id="195ea-735">Examples</span></span>

<span data-ttu-id="195ea-736">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="195ea-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="195ea-737">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-737">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="195ea-738">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-738">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="195ea-739">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="195ea-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="195ea-740">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="195ea-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="195ea-741">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="195ea-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="195ea-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="195ea-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="195ea-743">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-744">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="195ea-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-745">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-745">Requirements</span></span>

|<span data-ttu-id="195ea-746">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-746">Requirement</span></span>| <span data-ttu-id="195ea-747">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-748">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-749">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-749">1.0</span></span>|
|[<span data-ttu-id="195ea-750">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-751">ReadItem</span></span>|
|[<span data-ttu-id="195ea-752">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-753">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="195ea-754">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="195ea-754">Returns:</span></span>

<span data-ttu-id="195ea-755">Тип: [Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="195ea-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="195ea-756">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-756">Example</span></span>

<span data-ttu-id="195ea-757">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-757">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="195ea-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="195ea-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="195ea-759">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-760">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="195ea-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-761">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-761">Parameters</span></span>

|<span data-ttu-id="195ea-762">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-762">Name</span></span>| <span data-ttu-id="195ea-763">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-763">Type</span></span>| <span data-ttu-id="195ea-764">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="195ea-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="195ea-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="195ea-766">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="195ea-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="195ea-767">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-767">Requirements</span></span>

|<span data-ttu-id="195ea-768">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-768">Requirement</span></span>| <span data-ttu-id="195ea-769">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-770">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-771">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-771">1.0</span></span>|
|[<span data-ttu-id="195ea-772">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-772">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-773">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="195ea-773">Restricted</span></span>|
|[<span data-ttu-id="195ea-774">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-774">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-775">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="195ea-776">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="195ea-776">Returns:</span></span>

<span data-ttu-id="195ea-777">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="195ea-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="195ea-778">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="195ea-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="195ea-779">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="195ea-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="195ea-780">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="195ea-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="195ea-781">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="195ea-781">Value of `entityType`</span></span> | <span data-ttu-id="195ea-782">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="195ea-782">Type of objects in returned array</span></span> | <span data-ttu-id="195ea-783">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="195ea-784">String</span><span class="sxs-lookup"><span data-stu-id="195ea-784">String</span></span> | <span data-ttu-id="195ea-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="195ea-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="195ea-786">Contact</span><span class="sxs-lookup"><span data-stu-id="195ea-786">Contact</span></span> | <span data-ttu-id="195ea-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="195ea-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="195ea-788">String</span><span class="sxs-lookup"><span data-stu-id="195ea-788">String</span></span> | <span data-ttu-id="195ea-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="195ea-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="195ea-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="195ea-790">MeetingSuggestion</span></span> | <span data-ttu-id="195ea-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="195ea-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="195ea-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="195ea-792">PhoneNumber</span></span> | <span data-ttu-id="195ea-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="195ea-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="195ea-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="195ea-794">TaskSuggestion</span></span> | <span data-ttu-id="195ea-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="195ea-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="195ea-796">String</span><span class="sxs-lookup"><span data-stu-id="195ea-796">String</span></span> | <span data-ttu-id="195ea-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="195ea-797">**Restricted**</span></span> |

<span data-ttu-id="195ea-798">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="195ea-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="195ea-799">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-799">Example</span></span>

<span data-ttu-id="195ea-800">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="195ea-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="195ea-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="195ea-802">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="195ea-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-803">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="195ea-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="195ea-804">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="195ea-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-805">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-805">Parameters</span></span>

|<span data-ttu-id="195ea-806">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-806">Name</span></span>| <span data-ttu-id="195ea-807">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-807">Type</span></span>| <span data-ttu-id="195ea-808">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="195ea-809">String</span><span class="sxs-lookup"><span data-stu-id="195ea-809">String</span></span>|<span data-ttu-id="195ea-810">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="195ea-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="195ea-811">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-811">Requirements</span></span>

|<span data-ttu-id="195ea-812">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-812">Requirement</span></span>| <span data-ttu-id="195ea-813">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-814">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-815">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-815">1.0</span></span>|
|[<span data-ttu-id="195ea-816">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-816">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-817">ReadItem</span></span>|
|[<span data-ttu-id="195ea-818">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-818">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-819">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="195ea-820">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="195ea-820">Returns:</span></span>

<span data-ttu-id="195ea-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="195ea-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="195ea-823">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="195ea-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="195ea-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="195ea-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="195ea-825">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="195ea-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-826">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="195ea-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="195ea-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="195ea-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="195ea-830">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="195ea-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="195ea-831">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="195ea-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="195ea-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="195ea-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="195ea-835">Requirements</span><span class="sxs-lookup"><span data-stu-id="195ea-835">Requirements</span></span>

|<span data-ttu-id="195ea-836">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-836">Requirement</span></span>| <span data-ttu-id="195ea-837">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-838">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-839">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-839">1.0</span></span>|
|[<span data-ttu-id="195ea-840">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-840">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-841">ReadItem</span></span>|
|[<span data-ttu-id="195ea-842">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-842">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-843">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="195ea-844">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="195ea-844">Returns:</span></span>

<span data-ttu-id="195ea-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="195ea-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="195ea-847">Тип:</span><span class="sxs-lookup"><span data-stu-id="195ea-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="195ea-848">Object</span><span class="sxs-lookup"><span data-stu-id="195ea-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="195ea-849">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-849">Example</span></span>

<span data-ttu-id="195ea-850">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="195ea-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="195ea-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="195ea-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="195ea-852">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="195ea-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-853">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="195ea-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="195ea-854">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="195ea-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="195ea-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="195ea-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-857">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-857">Parameters</span></span>

|<span data-ttu-id="195ea-858">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-858">Name</span></span>| <span data-ttu-id="195ea-859">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-859">Type</span></span>| <span data-ttu-id="195ea-860">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="195ea-861">String</span><span class="sxs-lookup"><span data-stu-id="195ea-861">String</span></span>|<span data-ttu-id="195ea-862">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="195ea-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="195ea-863">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-863">Requirements</span></span>

|<span data-ttu-id="195ea-864">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-864">Requirement</span></span>| <span data-ttu-id="195ea-865">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-866">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-867">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-867">1.0</span></span>|
|[<span data-ttu-id="195ea-868">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-869">ReadItem</span></span>|
|[<span data-ttu-id="195ea-870">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-871">Чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="195ea-872">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="195ea-872">Returns:</span></span>

<span data-ttu-id="195ea-873">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="195ea-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="195ea-874">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="195ea-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="195ea-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="195ea-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="195ea-876">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-876">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="195ea-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="195ea-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="195ea-878">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="195ea-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="195ea-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-881">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-881">Parameters</span></span>

|<span data-ttu-id="195ea-882">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-882">Name</span></span>| <span data-ttu-id="195ea-883">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-883">Type</span></span>| <span data-ttu-id="195ea-884">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="195ea-884">Attributes</span></span>| <span data-ttu-id="195ea-885">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="195ea-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="195ea-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="195ea-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="195ea-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="195ea-890">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-890">Object</span></span>| <span data-ttu-id="195ea-891">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-891">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-892">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="195ea-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="195ea-893">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-893">Object</span></span>| <span data-ttu-id="195ea-894">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-894">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-895">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="195ea-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="195ea-896">функция</span><span class="sxs-lookup"><span data-stu-id="195ea-896">function</span></span>||<span data-ttu-id="195ea-897">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="195ea-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="195ea-898">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="195ea-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="195ea-899">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="195ea-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="195ea-900">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-900">Requirements</span></span>

|<span data-ttu-id="195ea-901">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-901">Requirement</span></span>| <span data-ttu-id="195ea-902">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-903">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-904">1.2</span><span class="sxs-lookup"><span data-stu-id="195ea-904">1.2</span></span>|
|[<span data-ttu-id="195ea-905">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="195ea-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="195ea-907">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-908">Создание</span><span class="sxs-lookup"><span data-stu-id="195ea-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="195ea-909">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="195ea-909">Returns:</span></span>

<span data-ttu-id="195ea-910">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="195ea-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="195ea-911">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="195ea-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="195ea-912">String</span><span class="sxs-lookup"><span data-stu-id="195ea-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="195ea-913">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="195ea-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="195ea-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="195ea-915">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="195ea-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="195ea-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-919">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-919">Parameters</span></span>

|<span data-ttu-id="195ea-920">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-920">Name</span></span>| <span data-ttu-id="195ea-921">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-921">Type</span></span>| <span data-ttu-id="195ea-922">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="195ea-922">Attributes</span></span>| <span data-ttu-id="195ea-923">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="195ea-924">function</span><span class="sxs-lookup"><span data-stu-id="195ea-924">function</span></span>||<span data-ttu-id="195ea-925">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="195ea-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="195ea-926">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="195ea-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="195ea-927">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="195ea-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="195ea-928">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-928">Object</span></span>| <span data-ttu-id="195ea-929">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-929">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-930">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="195ea-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="195ea-931">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="195ea-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="195ea-932">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-932">Requirements</span></span>

|<span data-ttu-id="195ea-933">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-933">Requirement</span></span>| <span data-ttu-id="195ea-934">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-935">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-936">1.0</span><span class="sxs-lookup"><span data-stu-id="195ea-936">1.0</span></span>|
|[<span data-ttu-id="195ea-937">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="195ea-938">ReadItem</span></span>|
|[<span data-ttu-id="195ea-939">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-940">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="195ea-940">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-941">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-941">Example</span></span>

<span data-ttu-id="195ea-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="195ea-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="195ea-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="195ea-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="195ea-946">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="195ea-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="195ea-p165">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="195ea-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-951">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-951">Parameters</span></span>

|<span data-ttu-id="195ea-952">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-952">Name</span></span>| <span data-ttu-id="195ea-953">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-953">Type</span></span>| <span data-ttu-id="195ea-954">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="195ea-954">Attributes</span></span>| <span data-ttu-id="195ea-955">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="195ea-956">String</span><span class="sxs-lookup"><span data-stu-id="195ea-956">String</span></span>||<span data-ttu-id="195ea-957">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="195ea-957">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="195ea-958">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-958">Object</span></span>| <span data-ttu-id="195ea-959">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-959">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-960">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="195ea-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="195ea-961">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-961">Object</span></span>| <span data-ttu-id="195ea-962">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-962">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-963">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="195ea-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="195ea-964">функция</span><span class="sxs-lookup"><span data-stu-id="195ea-964">function</span></span>| <span data-ttu-id="195ea-965">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-965">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-966">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="195ea-966">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="195ea-967">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="195ea-967">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="195ea-968">Ошибки</span><span class="sxs-lookup"><span data-stu-id="195ea-968">Errors</span></span>

| <span data-ttu-id="195ea-969">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="195ea-969">Error code</span></span> | <span data-ttu-id="195ea-970">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-970">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="195ea-971">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="195ea-971">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="195ea-972">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-972">Requirements</span></span>

|<span data-ttu-id="195ea-973">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-973">Requirement</span></span>| <span data-ttu-id="195ea-974">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-974">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-975">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="195ea-975">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-976">1.1</span><span class="sxs-lookup"><span data-stu-id="195ea-976">1.1</span></span>|
|[<span data-ttu-id="195ea-977">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-977">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-978">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="195ea-978">ReadWriteItem</span></span>|
|[<span data-ttu-id="195ea-979">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-979">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-980">Создание</span><span class="sxs-lookup"><span data-stu-id="195ea-980">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-981">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-981">Example</span></span>

<span data-ttu-id="195ea-982">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="195ea-982">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="195ea-983">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="195ea-983">saveAsync([options], callback)</span></span>

<span data-ttu-id="195ea-984">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="195ea-984">Asynchronously saves an item.</span></span>

<span data-ttu-id="195ea-p166">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="195ea-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-988">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="195ea-988">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="195ea-989">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="195ea-989">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="195ea-p168">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="195ea-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="195ea-993">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="195ea-993">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="195ea-994">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="195ea-994">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="195ea-995">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="195ea-995">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="195ea-996">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="195ea-996">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-997">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-997">Parameters</span></span>

|<span data-ttu-id="195ea-998">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-998">Name</span></span>| <span data-ttu-id="195ea-999">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-999">Type</span></span>| <span data-ttu-id="195ea-1000">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="195ea-1000">Attributes</span></span>| <span data-ttu-id="195ea-1001">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-1001">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="195ea-1002">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-1002">Object</span></span>| <span data-ttu-id="195ea-1003">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-1003">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-1004">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="195ea-1004">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="195ea-1005">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-1005">Object</span></span>| <span data-ttu-id="195ea-1006">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-1007">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="195ea-1007">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="195ea-1008">функция</span><span class="sxs-lookup"><span data-stu-id="195ea-1008">function</span></span>||<span data-ttu-id="195ea-1009">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="195ea-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="195ea-1010">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="195ea-1010">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="195ea-1011">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-1011">Requirements</span></span>

|<span data-ttu-id="195ea-1012">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-1012">Requirement</span></span>| <span data-ttu-id="195ea-1013">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-1014">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-1015">1.3</span><span class="sxs-lookup"><span data-stu-id="195ea-1015">1.3</span></span>|
|[<span data-ttu-id="195ea-1016">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="195ea-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="195ea-1018">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-1019">Создание</span><span class="sxs-lookup"><span data-stu-id="195ea-1019">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="195ea-1020">Примеры</span><span class="sxs-lookup"><span data-stu-id="195ea-1020">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="195ea-p170">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="195ea-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="195ea-1023">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="195ea-1023">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="195ea-1024">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="195ea-1024">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="195ea-p171">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="195ea-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="195ea-1028">Параметры</span><span class="sxs-lookup"><span data-stu-id="195ea-1028">Parameters</span></span>

|<span data-ttu-id="195ea-1029">Имя</span><span class="sxs-lookup"><span data-stu-id="195ea-1029">Name</span></span>| <span data-ttu-id="195ea-1030">Тип</span><span class="sxs-lookup"><span data-stu-id="195ea-1030">Type</span></span>| <span data-ttu-id="195ea-1031">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="195ea-1031">Attributes</span></span>| <span data-ttu-id="195ea-1032">Описание</span><span class="sxs-lookup"><span data-stu-id="195ea-1032">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="195ea-1033">String</span><span class="sxs-lookup"><span data-stu-id="195ea-1033">String</span></span>||<span data-ttu-id="195ea-p172">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="195ea-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="195ea-1037">Object</span><span class="sxs-lookup"><span data-stu-id="195ea-1037">Object</span></span>| <span data-ttu-id="195ea-1038">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-1039">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="195ea-1039">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="195ea-1040">Объект</span><span class="sxs-lookup"><span data-stu-id="195ea-1040">Object</span></span>| <span data-ttu-id="195ea-1041">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-1042">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="195ea-1042">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="195ea-1043">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="195ea-1043">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="195ea-1044">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="195ea-1044">&lt;optional&gt;</span></span>|<span data-ttu-id="195ea-p173">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="195ea-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="195ea-p174">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="195ea-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="195ea-1049">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="195ea-1049">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="195ea-1050">функция</span><span class="sxs-lookup"><span data-stu-id="195ea-1050">function</span></span>||<span data-ttu-id="195ea-1051">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="195ea-1051">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="195ea-1052">Требования</span><span class="sxs-lookup"><span data-stu-id="195ea-1052">Requirements</span></span>

|<span data-ttu-id="195ea-1053">Требование</span><span class="sxs-lookup"><span data-stu-id="195ea-1053">Requirement</span></span>| <span data-ttu-id="195ea-1054">Значение</span><span class="sxs-lookup"><span data-stu-id="195ea-1054">Value</span></span>|
|---|---|
|[<span data-ttu-id="195ea-1055">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="195ea-1055">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="195ea-1056">1.2</span><span class="sxs-lookup"><span data-stu-id="195ea-1056">1.2</span></span>|
|[<span data-ttu-id="195ea-1057">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="195ea-1057">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="195ea-1058">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="195ea-1058">ReadWriteItem</span></span>|
|[<span data-ttu-id="195ea-1059">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="195ea-1059">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="195ea-1060">Создание</span><span class="sxs-lookup"><span data-stu-id="195ea-1060">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="195ea-1061">Пример</span><span class="sxs-lookup"><span data-stu-id="195ea-1061">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
