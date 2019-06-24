---
title: Office. Context. Mailbox. Item — набор требований 1,4
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: c3b554b223a72363fec583db34c1dae74a302690
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127382"
---
# <a name="item"></a><span data-ttu-id="f8e9e-102">item</span><span class="sxs-lookup"><span data-stu-id="f8e9e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f8e9e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f8e9e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f8e9e-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="f8e9e-106">Requirements</span></span>

|<span data-ttu-id="f8e9e-107">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-107">Requirement</span></span>| <span data-ttu-id="f8e9e-108">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-110">1.0</span></span>|
|[<span data-ttu-id="f8e9e-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="f8e9e-112">Restricted</span></span>|
|[<span data-ttu-id="f8e9e-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-114">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="f8e9e-115">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-115">Example</span></span>

<span data-ttu-id="f8e9e-116">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="f8e9e-117">Элементы</span><span class="sxs-lookup"><span data-stu-id="f8e9e-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="f8e9e-118">вложения: Array. <[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f8e9e-118">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="f8e9e-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-121">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f8e9e-122">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-123">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-123">Type</span></span>

*   <span data-ttu-id="f8e9e-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f8e9e-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-125">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-125">Requirements</span></span>

|<span data-ttu-id="f8e9e-126">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-126">Requirement</span></span>| <span data-ttu-id="f8e9e-127">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-128">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-129">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-129">1.0</span></span>|
|[<span data-ttu-id="f8e9e-130">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-130">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-131">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-132">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-133">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-134">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-134">Example</span></span>

<span data-ttu-id="f8e9e-135">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f8e9e-136">СК: [получатели](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-136">bcc: [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f8e9e-137">Получает объект, который предоставляет методы для получения или обновления строки "СК" (Скрытая копия) сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-137">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f8e9e-138">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-139">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-139">Type</span></span>

*   [<span data-ttu-id="f8e9e-140">Получатели</span><span class="sxs-lookup"><span data-stu-id="f8e9e-140">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f8e9e-141">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-141">Requirements</span></span>

|<span data-ttu-id="f8e9e-142">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-142">Requirement</span></span>| <span data-ttu-id="f8e9e-143">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-145">1.1</span><span class="sxs-lookup"><span data-stu-id="f8e9e-145">1.1</span></span>|
|[<span data-ttu-id="f8e9e-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-147">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-149">Создание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-150">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-150">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="f8e9e-151">основной текст: [Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-151">body: [Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="f8e9e-152">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-153">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-153">Type</span></span>

*   [<span data-ttu-id="f8e9e-154">Body</span><span class="sxs-lookup"><span data-stu-id="f8e9e-154">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="f8e9e-155">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-155">Requirements</span></span>

|<span data-ttu-id="f8e9e-156">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-156">Requirement</span></span>| <span data-ttu-id="f8e9e-157">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-159">1.1</span><span class="sxs-lookup"><span data-stu-id="f8e9e-159">1.1</span></span>|
|[<span data-ttu-id="f8e9e-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-161">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-164">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-164">Example</span></span>

<span data-ttu-id="f8e9e-165">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-165">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="f8e9e-166">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-166">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f8e9e-167">CC: Array. <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-167">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f8e9e-168">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-168">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f8e9e-169">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-169">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8e9e-170">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="f8e9e-170">Read mode</span></span>

<span data-ttu-id="f8e9e-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="f8e9e-173">Режим создания</span><span class="sxs-lookup"><span data-stu-id="f8e9e-173">Compose mode</span></span>

<span data-ttu-id="f8e9e-174">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-174">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f8e9e-175">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-175">Type</span></span>

*   <span data-ttu-id="f8e9e-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-177">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-177">Requirements</span></span>

|<span data-ttu-id="f8e9e-178">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-178">Requirement</span></span>| <span data-ttu-id="f8e9e-179">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-180">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-181">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-181">1.0</span></span>|
|[<span data-ttu-id="f8e9e-182">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-183">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-185">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="f8e9e-186">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="f8e9e-186">(nullable) conversationId: String</span></span>

<span data-ttu-id="f8e9e-187">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-187">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f8e9e-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f8e9e-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-192">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-192">Type</span></span>

*   <span data-ttu-id="f8e9e-193">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-194">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-194">Requirements</span></span>

|<span data-ttu-id="f8e9e-195">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-195">Requirement</span></span>| <span data-ttu-id="f8e9e-196">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-197">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-198">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-198">1.0</span></span>|
|[<span data-ttu-id="f8e9e-199">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-199">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-200">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-200">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-201">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-203">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-203">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="f8e9e-204">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="f8e9e-204">dateTimeCreated: Date</span></span>

<span data-ttu-id="f8e9e-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-207">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-207">Type</span></span>

*   <span data-ttu-id="f8e9e-208">Дата</span><span class="sxs-lookup"><span data-stu-id="f8e9e-208">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-209">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-209">Requirements</span></span>

|<span data-ttu-id="f8e9e-210">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-210">Requirement</span></span>| <span data-ttu-id="f8e9e-211">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-212">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-213">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-213">1.0</span></span>|
|[<span data-ttu-id="f8e9e-214">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-215">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-216">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-217">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-218">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-218">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="f8e9e-219">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="f8e9e-219">dateTimeModified: Date</span></span>

<span data-ttu-id="f8e9e-220">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-220">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="f8e9e-221">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-221">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-222">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-222">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-223">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-223">Type</span></span>

*   <span data-ttu-id="f8e9e-224">Дата</span><span class="sxs-lookup"><span data-stu-id="f8e9e-224">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-225">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-225">Requirements</span></span>

|<span data-ttu-id="f8e9e-226">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-226">Requirement</span></span>| <span data-ttu-id="f8e9e-227">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-228">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-229">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-229">1.0</span></span>|
|[<span data-ttu-id="f8e9e-230">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-231">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-232">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-233">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-233">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-234">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-234">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="f8e9e-235">конец: Дата | [Time (время](/javascript/api/outlook_1_4/office.time) )</span><span class="sxs-lookup"><span data-stu-id="f8e9e-235">end: Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="f8e9e-236">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-236">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f8e9e-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8e9e-239">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="f8e9e-239">Read mode</span></span>

<span data-ttu-id="f8e9e-240">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-240">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="f8e9e-241">Режим создания</span><span class="sxs-lookup"><span data-stu-id="f8e9e-241">Compose mode</span></span>

<span data-ttu-id="f8e9e-242">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-242">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f8e9e-243">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-243">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f8e9e-244">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-244">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f8e9e-245">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-245">Type</span></span>

*   <span data-ttu-id="f8e9e-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-247">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-247">Requirements</span></span>

|<span data-ttu-id="f8e9e-248">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-248">Requirement</span></span>| <span data-ttu-id="f8e9e-249">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-250">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-251">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-251">1.0</span></span>|
|[<span data-ttu-id="f8e9e-252">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-253">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-254">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-255">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-255">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="f8e9e-256">от: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-256">from: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="f8e9e-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="f8e9e-p113">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-261">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-261">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-262">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-262">Type</span></span>

*   [<span data-ttu-id="f8e9e-263">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f8e9e-263">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f8e9e-264">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-264">Requirements</span></span>

|<span data-ttu-id="f8e9e-265">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-265">Requirement</span></span>| <span data-ttu-id="f8e9e-266">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-267">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-268">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-268">1.0</span></span>|
|[<span data-ttu-id="f8e9e-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-270">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-272">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-272">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-273">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-273">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="f8e9e-274">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="f8e9e-274">internetMessageId: String</span></span>

<span data-ttu-id="f8e9e-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-277">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-277">Type</span></span>

*   <span data-ttu-id="f8e9e-278">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-278">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-279">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-279">Requirements</span></span>

|<span data-ttu-id="f8e9e-280">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-280">Requirement</span></span>| <span data-ttu-id="f8e9e-281">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-282">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-283">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-283">1.0</span></span>|
|[<span data-ttu-id="f8e9e-284">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-285">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-286">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-287">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-288">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-288">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="f8e9e-289">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="f8e9e-289">itemClass: String</span></span>

<span data-ttu-id="f8e9e-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f8e9e-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="f8e9e-294">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-294">Type</span></span> | <span data-ttu-id="f8e9e-295">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-295">Description</span></span> | <span data-ttu-id="f8e9e-296">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="f8e9e-296">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="f8e9e-297">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="f8e9e-297">Appointment items</span></span> | <span data-ttu-id="f8e9e-298">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-298">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="f8e9e-299">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="f8e9e-299">Message items</span></span> | <span data-ttu-id="f8e9e-300">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-300">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="f8e9e-301">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-301">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-302">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-302">Type</span></span>

*   <span data-ttu-id="f8e9e-303">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-303">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-304">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-304">Requirements</span></span>

|<span data-ttu-id="f8e9e-305">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-305">Requirement</span></span>| <span data-ttu-id="f8e9e-306">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-307">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-308">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-308">1.0</span></span>|
|[<span data-ttu-id="f8e9e-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-310">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-313">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-313">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f8e9e-314">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="f8e9e-314">(nullable) itemId: String</span></span>

<span data-ttu-id="f8e9e-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-317">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-317">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f8e9e-318">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-318">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f8e9e-319">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-319">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f8e9e-320">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-320">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f8e9e-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-323">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-323">Type</span></span>

*   <span data-ttu-id="f8e9e-324">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-324">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-325">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-325">Requirements</span></span>

|<span data-ttu-id="f8e9e-326">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-326">Requirement</span></span>| <span data-ttu-id="f8e9e-327">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-328">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-329">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-329">1.0</span></span>|
|[<span data-ttu-id="f8e9e-330">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-331">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-332">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-333">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-334">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-334">Example</span></span>

<span data-ttu-id="f8e9e-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="f8e9e-337">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-337">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f8e9e-338">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-338">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f8e9e-339">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-339">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-340">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-340">Type</span></span>

*   [<span data-ttu-id="f8e9e-341">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f8e9e-341">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f8e9e-342">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-342">Requirements</span></span>

|<span data-ttu-id="f8e9e-343">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-343">Requirement</span></span>| <span data-ttu-id="f8e9e-344">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-344">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-345">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-346">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-346">1.0</span></span>|
|[<span data-ttu-id="f8e9e-347">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-348">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-349">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-350">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-350">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-351">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-351">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="f8e9e-352">Местоположение: строка | [Location (расположение](/javascript/api/outlook_1_4/office.location) )</span><span class="sxs-lookup"><span data-stu-id="f8e9e-352">location: String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="f8e9e-353">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-353">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8e9e-354">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="f8e9e-354">Read mode</span></span>

<span data-ttu-id="f8e9e-355">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-355">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="f8e9e-356">Режим создания</span><span class="sxs-lookup"><span data-stu-id="f8e9e-356">Compose mode</span></span>

<span data-ttu-id="f8e9e-357">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-357">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f8e9e-358">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-358">Type</span></span>

*   <span data-ttu-id="f8e9e-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-360">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-360">Requirements</span></span>

|<span data-ttu-id="f8e9e-361">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-361">Requirement</span></span>| <span data-ttu-id="f8e9e-362">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-363">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-364">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-364">1.0</span></span>|
|[<span data-ttu-id="f8e9e-365">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-366">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-367">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-368">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-368">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f8e9e-369">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="f8e9e-369">normalizedSubject: String</span></span>

<span data-ttu-id="f8e9e-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f8e9e-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-374">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-374">Type</span></span>

*   <span data-ttu-id="f8e9e-375">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-375">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-376">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-376">Requirements</span></span>

|<span data-ttu-id="f8e9e-377">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-377">Requirement</span></span>| <span data-ttu-id="f8e9e-378">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-378">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-379">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-380">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-380">1.0</span></span>|
|[<span data-ttu-id="f8e9e-381">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-382">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-383">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-384">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-384">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-385">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-385">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="f8e9e-386">notificationMessages: [notificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-386">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="f8e9e-387">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-387">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-388">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-388">Type</span></span>

*   [<span data-ttu-id="f8e9e-389">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f8e9e-389">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="f8e9e-390">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-390">Requirements</span></span>

|<span data-ttu-id="f8e9e-391">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-391">Requirement</span></span>| <span data-ttu-id="f8e9e-392">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-392">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-393">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-394">1.3</span><span class="sxs-lookup"><span data-stu-id="f8e9e-394">1.3</span></span>|
|[<span data-ttu-id="f8e9e-395">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-395">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-396">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-397">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-398">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-398">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-399">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-399">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f8e9e-400">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-400">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f8e9e-401">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-401">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f8e9e-402">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-402">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8e9e-403">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="f8e9e-403">Read mode</span></span>

<span data-ttu-id="f8e9e-404">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-404">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f8e9e-405">Режим создания</span><span class="sxs-lookup"><span data-stu-id="f8e9e-405">Compose mode</span></span>

<span data-ttu-id="f8e9e-406">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-406">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f8e9e-407">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-407">Type</span></span>

*   <span data-ttu-id="f8e9e-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-409">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-409">Requirements</span></span>

|<span data-ttu-id="f8e9e-410">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-410">Requirement</span></span>| <span data-ttu-id="f8e9e-411">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-412">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-413">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-413">1.0</span></span>|
|[<span data-ttu-id="f8e9e-414">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-415">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-416">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-417">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-417">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="f8e9e-418">Организатор: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-418">organizer: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="f8e9e-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-421">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-421">Type</span></span>

*   [<span data-ttu-id="f8e9e-422">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f8e9e-422">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f8e9e-423">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-423">Requirements</span></span>

|<span data-ttu-id="f8e9e-424">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-424">Requirement</span></span>| <span data-ttu-id="f8e9e-425">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-426">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-427">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-427">1.0</span></span>|
|[<span data-ttu-id="f8e9e-428">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-429">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-430">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-431">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-432">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-432">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f8e9e-433">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-433">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f8e9e-434">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-434">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f8e9e-435">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-435">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8e9e-436">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="f8e9e-436">Read mode</span></span>

<span data-ttu-id="f8e9e-437">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-437">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f8e9e-438">Режим создания</span><span class="sxs-lookup"><span data-stu-id="f8e9e-438">Compose mode</span></span>

<span data-ttu-id="f8e9e-439">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-439">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="f8e9e-440">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-440">Type</span></span>

*   <span data-ttu-id="f8e9e-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-442">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-442">Requirements</span></span>

|<span data-ttu-id="f8e9e-443">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-443">Requirement</span></span>| <span data-ttu-id="f8e9e-444">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-445">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-446">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-446">1.0</span></span>|
|[<span data-ttu-id="f8e9e-447">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-447">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-448">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-449">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-449">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-450">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-450">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="f8e9e-451">Отправитель: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-451">sender: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="f8e9e-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f8e9e-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-456">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-456">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f8e9e-457">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-457">Type</span></span>

*   [<span data-ttu-id="f8e9e-458">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f8e9e-458">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f8e9e-459">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-459">Requirements</span></span>

|<span data-ttu-id="f8e9e-460">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-460">Requirement</span></span>| <span data-ttu-id="f8e9e-461">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-462">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-463">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-463">1.0</span></span>|
|[<span data-ttu-id="f8e9e-464">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-465">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-466">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-467">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-468">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-468">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="f8e9e-469">Начало: Дата | [Time (время](/javascript/api/outlook_1_4/office.time) )</span><span class="sxs-lookup"><span data-stu-id="f8e9e-469">start: Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="f8e9e-470">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-470">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f8e9e-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8e9e-473">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="f8e9e-473">Read mode</span></span>

<span data-ttu-id="f8e9e-474">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-474">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="f8e9e-475">Режим создания</span><span class="sxs-lookup"><span data-stu-id="f8e9e-475">Compose mode</span></span>

<span data-ttu-id="f8e9e-476">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-476">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f8e9e-477">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-477">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f8e9e-478">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-478">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f8e9e-479">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-479">Type</span></span>

*   <span data-ttu-id="f8e9e-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-481">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-481">Requirements</span></span>

|<span data-ttu-id="f8e9e-482">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-482">Requirement</span></span>| <span data-ttu-id="f8e9e-483">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-484">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-485">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-485">1.0</span></span>|
|[<span data-ttu-id="f8e9e-486">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-487">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-488">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-489">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-489">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="f8e9e-490">Тема: строка | [Subject (тема](/javascript/api/outlook_1_4/office.subject) )</span><span class="sxs-lookup"><span data-stu-id="f8e9e-490">subject: String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="f8e9e-491">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f8e9e-492">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8e9e-493">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="f8e9e-493">Read mode</span></span>

<span data-ttu-id="f8e9e-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="f8e9e-496">Режим создания</span><span class="sxs-lookup"><span data-stu-id="f8e9e-496">Compose mode</span></span>

<span data-ttu-id="f8e9e-497">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="f8e9e-498">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-498">Type</span></span>

*   <span data-ttu-id="f8e9e-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-500">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-500">Requirements</span></span>

|<span data-ttu-id="f8e9e-501">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-501">Requirement</span></span>| <span data-ttu-id="f8e9e-502">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-503">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-504">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-504">1.0</span></span>|
|[<span data-ttu-id="f8e9e-505">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-506">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-507">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-508">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-508">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f8e9e-509">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-509">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f8e9e-510">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f8e9e-511">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8e9e-512">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="f8e9e-512">Read mode</span></span>

<span data-ttu-id="f8e9e-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="f8e9e-515">Режим создания</span><span class="sxs-lookup"><span data-stu-id="f8e9e-515">Compose mode</span></span>

<span data-ttu-id="f8e9e-516">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f8e9e-517">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-517">Type</span></span>

*   <span data-ttu-id="f8e9e-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-519">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-519">Requirements</span></span>

|<span data-ttu-id="f8e9e-520">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-520">Requirement</span></span>| <span data-ttu-id="f8e9e-521">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-522">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-523">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-523">1.0</span></span>|
|[<span data-ttu-id="f8e9e-524">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-525">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-526">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-527">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-527">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="f8e9e-528">Методы</span><span class="sxs-lookup"><span data-stu-id="f8e9e-528">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f8e9e-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f8e9e-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f8e9e-530">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f8e9e-531">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f8e9e-532">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-533">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-533">Parameters</span></span>

|<span data-ttu-id="f8e9e-534">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-534">Name</span></span>| <span data-ttu-id="f8e9e-535">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-535">Type</span></span>| <span data-ttu-id="f8e9e-536">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f8e9e-536">Attributes</span></span>| <span data-ttu-id="f8e9e-537">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="f8e9e-538">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-538">String</span></span>||<span data-ttu-id="f8e9e-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f8e9e-541">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-541">String</span></span>||<span data-ttu-id="f8e9e-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f8e9e-544">Объект</span><span class="sxs-lookup"><span data-stu-id="f8e9e-544">Object</span></span>| <span data-ttu-id="f8e9e-545">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-545">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-546">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8e9e-547">Object</span><span class="sxs-lookup"><span data-stu-id="f8e9e-547">Object</span></span>| <span data-ttu-id="f8e9e-548">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-548">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-549">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f8e9e-550">функция</span><span class="sxs-lookup"><span data-stu-id="f8e9e-550">function</span></span>| <span data-ttu-id="f8e9e-551">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-551">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-552">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f8e9e-553">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f8e9e-554">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f8e9e-555">Ошибки</span><span class="sxs-lookup"><span data-stu-id="f8e9e-555">Errors</span></span>

| <span data-ttu-id="f8e9e-556">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="f8e9e-556">Error code</span></span> | <span data-ttu-id="f8e9e-557">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="f8e9e-558">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="f8e9e-559">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f8e9e-560">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8e9e-561">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-561">Requirements</span></span>

|<span data-ttu-id="f8e9e-562">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-562">Requirement</span></span>| <span data-ttu-id="f8e9e-563">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-564">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-565">1.1</span><span class="sxs-lookup"><span data-stu-id="f8e9e-565">1.1</span></span>|
|[<span data-ttu-id="f8e9e-566">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8e9e-568">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-569">Создание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-570">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-570">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f8e9e-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f8e9e-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f8e9e-572">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f8e9e-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f8e9e-576">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f8e9e-577">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-577">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-578">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-578">Parameters</span></span>

|<span data-ttu-id="f8e9e-579">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-579">Name</span></span>| <span data-ttu-id="f8e9e-580">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-580">Type</span></span>| <span data-ttu-id="f8e9e-581">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f8e9e-581">Attributes</span></span>| <span data-ttu-id="f8e9e-582">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="f8e9e-583">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-583">String</span></span>||<span data-ttu-id="f8e9e-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f8e9e-586">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-586">String</span></span>||<span data-ttu-id="f8e9e-587">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-587">The subject of the item to be attached.</span></span> <span data-ttu-id="f8e9e-588">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-588">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f8e9e-589">Object</span><span class="sxs-lookup"><span data-stu-id="f8e9e-589">Object</span></span>| <span data-ttu-id="f8e9e-590">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-590">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-591">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8e9e-592">Объект</span><span class="sxs-lookup"><span data-stu-id="f8e9e-592">Object</span></span>| <span data-ttu-id="f8e9e-593">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-593">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-594">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f8e9e-595">функция</span><span class="sxs-lookup"><span data-stu-id="f8e9e-595">function</span></span>| <span data-ttu-id="f8e9e-596">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-596">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-597">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f8e9e-598">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f8e9e-599">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f8e9e-600">Ошибки</span><span class="sxs-lookup"><span data-stu-id="f8e9e-600">Errors</span></span>

| <span data-ttu-id="f8e9e-601">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="f8e9e-601">Error code</span></span> | <span data-ttu-id="f8e9e-602">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f8e9e-603">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8e9e-604">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-604">Requirements</span></span>

|<span data-ttu-id="f8e9e-605">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-605">Requirement</span></span>| <span data-ttu-id="f8e9e-606">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-607">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-608">1.1</span><span class="sxs-lookup"><span data-stu-id="f8e9e-608">1.1</span></span>|
|[<span data-ttu-id="f8e9e-609">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8e9e-611">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-612">Создание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-613">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-613">Example</span></span>

<span data-ttu-id="f8e9e-614">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="f8e9e-615">close()</span><span class="sxs-lookup"><span data-stu-id="f8e9e-615">close()</span></span>

<span data-ttu-id="f8e9e-616">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f8e9e-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-619">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f8e9e-620">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-621">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-621">Requirements</span></span>

|<span data-ttu-id="f8e9e-622">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-622">Requirement</span></span>| <span data-ttu-id="f8e9e-623">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-624">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-625">1.3</span><span class="sxs-lookup"><span data-stu-id="f8e9e-625">1.3</span></span>|
|[<span data-ttu-id="f8e9e-626">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-626">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-627">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="f8e9e-627">Restricted</span></span>|
|[<span data-ttu-id="f8e9e-628">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-628">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-629">Создание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-629">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="f8e9e-630">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f8e9e-630">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="f8e9e-631">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-632">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-632">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f8e9e-633">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-633">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f8e9e-634">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f8e9e-635">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-635">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="f8e9e-636">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-636">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="f8e9e-637">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-637">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-638">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-638">Parameters</span></span>

|<span data-ttu-id="f8e9e-639">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-639">Name</span></span>| <span data-ttu-id="f8e9e-640">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-640">Type</span></span>| <span data-ttu-id="f8e9e-641">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f8e9e-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f8e9e-642">String &#124; Object</span></span>| |<span data-ttu-id="f8e9e-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f8e9e-645">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="f8e9e-645">**OR**</span></span><br/><span data-ttu-id="f8e9e-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f8e9e-648">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-648">String</span></span> | <span data-ttu-id="f8e9e-649">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-649">&lt;optional&gt;</span></span> | <span data-ttu-id="f8e9e-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f8e9e-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f8e9e-653">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-653">&lt;optional&gt;</span></span> | <span data-ttu-id="f8e9e-654">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f8e9e-655">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-655">String</span></span> | | <span data-ttu-id="f8e9e-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f8e9e-658">Строка</span><span class="sxs-lookup"><span data-stu-id="f8e9e-658">String</span></span> | | <span data-ttu-id="f8e9e-659">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f8e9e-660">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-660">String</span></span> | | <span data-ttu-id="f8e9e-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f8e9e-663">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-663">String</span></span> | | <span data-ttu-id="f8e9e-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f8e9e-667">function</span><span class="sxs-lookup"><span data-stu-id="f8e9e-667">function</span></span> | <span data-ttu-id="f8e9e-668">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-668">&lt;optional&gt;</span></span> | <span data-ttu-id="f8e9e-669">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8e9e-670">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-670">Requirements</span></span>

|<span data-ttu-id="f8e9e-671">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-671">Requirement</span></span>| <span data-ttu-id="f8e9e-672">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-673">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-674">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-674">1.0</span></span>|
|[<span data-ttu-id="f8e9e-675">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-676">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-677">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-678">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f8e9e-679">Примеры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-679">Examples</span></span>

<span data-ttu-id="f8e9e-680">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f8e9e-681">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-681">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f8e9e-682">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-682">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f8e9e-683">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f8e9e-684">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f8e9e-685">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="f8e9e-686">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f8e9e-686">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="f8e9e-687">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-688">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-688">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f8e9e-689">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-689">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f8e9e-690">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f8e9e-691">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-691">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="f8e9e-692">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-692">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="f8e9e-693">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-693">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-694">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-694">Parameters</span></span>

|<span data-ttu-id="f8e9e-695">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-695">Name</span></span>| <span data-ttu-id="f8e9e-696">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-696">Type</span></span>| <span data-ttu-id="f8e9e-697">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f8e9e-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f8e9e-698">String &#124; Object</span></span>| | <span data-ttu-id="f8e9e-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f8e9e-701">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="f8e9e-701">**OR**</span></span><br/><span data-ttu-id="f8e9e-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f8e9e-704">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-704">String</span></span> | <span data-ttu-id="f8e9e-705">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-705">&lt;optional&gt;</span></span> | <span data-ttu-id="f8e9e-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f8e9e-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f8e9e-709">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-709">&lt;optional&gt;</span></span> | <span data-ttu-id="f8e9e-710">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f8e9e-711">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-711">String</span></span> | | <span data-ttu-id="f8e9e-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f8e9e-714">Строка</span><span class="sxs-lookup"><span data-stu-id="f8e9e-714">String</span></span> | | <span data-ttu-id="f8e9e-715">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f8e9e-716">Строка</span><span class="sxs-lookup"><span data-stu-id="f8e9e-716">String</span></span> | | <span data-ttu-id="f8e9e-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f8e9e-719">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-719">String</span></span> | | <span data-ttu-id="f8e9e-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f8e9e-723">function</span><span class="sxs-lookup"><span data-stu-id="f8e9e-723">function</span></span> | <span data-ttu-id="f8e9e-724">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-724">&lt;optional&gt;</span></span> | <span data-ttu-id="f8e9e-725">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8e9e-726">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-726">Requirements</span></span>

|<span data-ttu-id="f8e9e-727">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-727">Requirement</span></span>| <span data-ttu-id="f8e9e-728">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-729">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-730">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-730">1.0</span></span>|
|[<span data-ttu-id="f8e9e-731">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-732">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-733">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-734">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f8e9e-735">Примеры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-735">Examples</span></span>

<span data-ttu-id="f8e9e-736">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f8e9e-737">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-737">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f8e9e-738">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-738">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f8e9e-739">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f8e9e-740">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f8e9e-741">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="f8e9e-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f8e9e-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="f8e9e-743">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-744">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-744">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-745">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-745">Requirements</span></span>

|<span data-ttu-id="f8e9e-746">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-746">Requirement</span></span>| <span data-ttu-id="f8e9e-747">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-748">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-749">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-749">1.0</span></span>|
|[<span data-ttu-id="f8e9e-750">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-751">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-752">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-753">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8e9e-754">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f8e9e-754">Returns:</span></span>

<span data-ttu-id="f8e9e-755">Тип: [Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f8e9e-756">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-756">Example</span></span>

<span data-ttu-id="f8e9e-757">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-757">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="f8e9e-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f8e9e-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f8e9e-759">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-760">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-760">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-761">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-761">Parameters</span></span>

|<span data-ttu-id="f8e9e-762">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-762">Name</span></span>| <span data-ttu-id="f8e9e-763">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-763">Type</span></span>| <span data-ttu-id="f8e9e-764">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="f8e9e-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f8e9e-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="f8e9e-766">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8e9e-767">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-767">Requirements</span></span>

|<span data-ttu-id="f8e9e-768">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-768">Requirement</span></span>| <span data-ttu-id="f8e9e-769">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-770">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-771">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-771">1.0</span></span>|
|[<span data-ttu-id="f8e9e-772">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-772">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-773">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="f8e9e-773">Restricted</span></span>|
|[<span data-ttu-id="f8e9e-774">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-774">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-775">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8e9e-776">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f8e9e-776">Returns:</span></span>

<span data-ttu-id="f8e9e-777">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f8e9e-778">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="f8e9e-779">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f8e9e-780">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="f8e9e-781">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="f8e9e-781">Value of `entityType`</span></span> | <span data-ttu-id="f8e9e-782">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="f8e9e-782">Type of objects in returned array</span></span> | <span data-ttu-id="f8e9e-783">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="f8e9e-784">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-784">String</span></span> | <span data-ttu-id="f8e9e-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f8e9e-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="f8e9e-786">Contact</span><span class="sxs-lookup"><span data-stu-id="f8e9e-786">Contact</span></span> | <span data-ttu-id="f8e9e-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f8e9e-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="f8e9e-788">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-788">String</span></span> | <span data-ttu-id="f8e9e-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f8e9e-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="f8e9e-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f8e9e-790">MeetingSuggestion</span></span> | <span data-ttu-id="f8e9e-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f8e9e-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="f8e9e-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f8e9e-792">PhoneNumber</span></span> | <span data-ttu-id="f8e9e-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f8e9e-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="f8e9e-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f8e9e-794">TaskSuggestion</span></span> | <span data-ttu-id="f8e9e-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f8e9e-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="f8e9e-796">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-796">String</span></span> | <span data-ttu-id="f8e9e-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f8e9e-797">**Restricted**</span></span> |

<span data-ttu-id="f8e9e-798">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f8e9e-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f8e9e-799">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-799">Example</span></span>

<span data-ttu-id="f8e9e-800">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="f8e9e-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f8e9e-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f8e9e-802">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-803">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-803">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f8e9e-804">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-805">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-805">Parameters</span></span>

|<span data-ttu-id="f8e9e-806">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-806">Name</span></span>| <span data-ttu-id="f8e9e-807">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-807">Type</span></span>| <span data-ttu-id="f8e9e-808">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f8e9e-809">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-809">String</span></span>|<span data-ttu-id="f8e9e-810">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8e9e-811">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-811">Requirements</span></span>

|<span data-ttu-id="f8e9e-812">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-812">Requirement</span></span>| <span data-ttu-id="f8e9e-813">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-814">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-815">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-815">1.0</span></span>|
|[<span data-ttu-id="f8e9e-816">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-816">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-817">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-818">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-818">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-819">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8e9e-820">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f8e9e-820">Returns:</span></span>

<span data-ttu-id="f8e9e-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f8e9e-823">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f8e9e-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="f8e9e-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f8e9e-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f8e9e-825">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-826">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-826">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f8e9e-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f8e9e-830">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f8e9e-831">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f8e9e-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8e9e-835">Requirements</span><span class="sxs-lookup"><span data-stu-id="f8e9e-835">Requirements</span></span>

|<span data-ttu-id="f8e9e-836">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-836">Requirement</span></span>| <span data-ttu-id="f8e9e-837">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-838">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-839">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-839">1.0</span></span>|
|[<span data-ttu-id="f8e9e-840">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-840">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-841">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-842">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-842">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-843">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8e9e-844">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f8e9e-844">Returns:</span></span>

<span data-ttu-id="f8e9e-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f8e9e-847">Тип:</span><span class="sxs-lookup"><span data-stu-id="f8e9e-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f8e9e-848">Object</span><span class="sxs-lookup"><span data-stu-id="f8e9e-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f8e9e-849">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-849">Example</span></span>

<span data-ttu-id="f8e9e-850">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f8e9e-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="f8e9e-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f8e9e-852">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-853">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-853">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f8e9e-854">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f8e9e-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-857">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-857">Parameters</span></span>

|<span data-ttu-id="f8e9e-858">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-858">Name</span></span>| <span data-ttu-id="f8e9e-859">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-859">Type</span></span>| <span data-ttu-id="f8e9e-860">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f8e9e-861">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-861">String</span></span>|<span data-ttu-id="f8e9e-862">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8e9e-863">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-863">Requirements</span></span>

|<span data-ttu-id="f8e9e-864">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-864">Requirement</span></span>| <span data-ttu-id="f8e9e-865">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-866">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-867">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-867">1.0</span></span>|
|[<span data-ttu-id="f8e9e-868">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-869">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-870">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-871">Чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8e9e-872">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f8e9e-872">Returns:</span></span>

<span data-ttu-id="f8e9e-873">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="f8e9e-874">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="f8e9e-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f8e9e-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="f8e9e-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f8e9e-876">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-876">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f8e9e-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f8e9e-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f8e9e-878">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f8e9e-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-881">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-881">Parameters</span></span>

|<span data-ttu-id="f8e9e-882">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-882">Name</span></span>| <span data-ttu-id="f8e9e-883">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-883">Type</span></span>| <span data-ttu-id="f8e9e-884">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f8e9e-884">Attributes</span></span>| <span data-ttu-id="f8e9e-885">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="f8e9e-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f8e9e-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f8e9e-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="f8e9e-890">Объект</span><span class="sxs-lookup"><span data-stu-id="f8e9e-890">Object</span></span>| <span data-ttu-id="f8e9e-891">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-891">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-892">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8e9e-893">Объект</span><span class="sxs-lookup"><span data-stu-id="f8e9e-893">Object</span></span>| <span data-ttu-id="f8e9e-894">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-894">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-895">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f8e9e-896">функция</span><span class="sxs-lookup"><span data-stu-id="f8e9e-896">function</span></span>||<span data-ttu-id="f8e9e-897">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f8e9e-898">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f8e9e-899">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8e9e-900">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-900">Requirements</span></span>

|<span data-ttu-id="f8e9e-901">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-901">Requirement</span></span>| <span data-ttu-id="f8e9e-902">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-903">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-904">1.2</span><span class="sxs-lookup"><span data-stu-id="f8e9e-904">1.2</span></span>|
|[<span data-ttu-id="f8e9e-905">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8e9e-907">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-908">Создание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8e9e-909">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f8e9e-909">Returns:</span></span>

<span data-ttu-id="f8e9e-910">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="f8e9e-911">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="f8e9e-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f8e9e-912">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f8e9e-913">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-913">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f8e9e-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f8e9e-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f8e9e-915">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f8e9e-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-919">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-919">Parameters</span></span>

|<span data-ttu-id="f8e9e-920">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-920">Name</span></span>| <span data-ttu-id="f8e9e-921">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-921">Type</span></span>| <span data-ttu-id="f8e9e-922">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f8e9e-922">Attributes</span></span>| <span data-ttu-id="f8e9e-923">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f8e9e-924">function</span><span class="sxs-lookup"><span data-stu-id="f8e9e-924">function</span></span>||<span data-ttu-id="f8e9e-925">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f8e9e-926">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f8e9e-927">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="f8e9e-928">Объект</span><span class="sxs-lookup"><span data-stu-id="f8e9e-928">Object</span></span>| <span data-ttu-id="f8e9e-929">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-929">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-930">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="f8e9e-931">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8e9e-932">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-932">Requirements</span></span>

|<span data-ttu-id="f8e9e-933">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-933">Requirement</span></span>| <span data-ttu-id="f8e9e-934">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-935">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-936">1.0</span><span class="sxs-lookup"><span data-stu-id="f8e9e-936">1.0</span></span>|
|[<span data-ttu-id="f8e9e-937">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-938">ReadItem</span></span>|
|[<span data-ttu-id="f8e9e-939">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-940">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-940">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-941">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-941">Example</span></span>

<span data-ttu-id="f8e9e-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f8e9e-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f8e9e-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f8e9e-946">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f8e9e-947">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-947">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="f8e9e-948">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-948">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="f8e9e-949">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-949">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="f8e9e-950">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-950">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-951">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-951">Parameters</span></span>

|<span data-ttu-id="f8e9e-952">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-952">Name</span></span>| <span data-ttu-id="f8e9e-953">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-953">Type</span></span>| <span data-ttu-id="f8e9e-954">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f8e9e-954">Attributes</span></span>| <span data-ttu-id="f8e9e-955">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="f8e9e-956">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-956">String</span></span>||<span data-ttu-id="f8e9e-957">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-957">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="f8e9e-958">Объект</span><span class="sxs-lookup"><span data-stu-id="f8e9e-958">Object</span></span>| <span data-ttu-id="f8e9e-959">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-959">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-960">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8e9e-961">Объект</span><span class="sxs-lookup"><span data-stu-id="f8e9e-961">Object</span></span>| <span data-ttu-id="f8e9e-962">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-962">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-963">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f8e9e-964">функция</span><span class="sxs-lookup"><span data-stu-id="f8e9e-964">function</span></span>| <span data-ttu-id="f8e9e-965">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-965">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-966">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-966">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f8e9e-967">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-967">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f8e9e-968">Ошибки</span><span class="sxs-lookup"><span data-stu-id="f8e9e-968">Errors</span></span>

| <span data-ttu-id="f8e9e-969">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="f8e9e-969">Error code</span></span> | <span data-ttu-id="f8e9e-970">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-970">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="f8e9e-971">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-971">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8e9e-972">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-972">Requirements</span></span>

|<span data-ttu-id="f8e9e-973">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-973">Requirement</span></span>| <span data-ttu-id="f8e9e-974">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-974">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-975">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f8e9e-975">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-976">1.1</span><span class="sxs-lookup"><span data-stu-id="f8e9e-976">1.1</span></span>|
|[<span data-ttu-id="f8e9e-977">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-977">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-978">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-978">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8e9e-979">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-979">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-980">Создание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-980">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-981">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-981">Example</span></span>

<span data-ttu-id="f8e9e-982">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="f8e9e-982">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="f8e9e-983">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-983">saveAsync([options], callback)</span></span>

<span data-ttu-id="f8e9e-984">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-984">Asynchronously saves an item.</span></span>

<span data-ttu-id="f8e9e-985">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-985">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="f8e9e-986">В Outlook в Интернете или Outlook в интерактивном режиме элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-986">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="f8e9e-987">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-987">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-988">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-988">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="f8e9e-989">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-989">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f8e9e-p168">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f8e9e-993">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="f8e9e-993">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f8e9e-994">Outlook в Mac не поддерживает сохранение собраний.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-994">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="f8e9e-995">`saveAsync` Метод завершается с ошибкой при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-995">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="f8e9e-996">Просмотреть [не удается сохранить собрание в виде черновика в Outlook для Mac с помощью API Office JS](https://support.microsoft.com/help/4505745) для обхода.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-996">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="f8e9e-997">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-998">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-998">Parameters</span></span>

|<span data-ttu-id="f8e9e-999">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-999">Name</span></span>| <span data-ttu-id="f8e9e-1000">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1000">Type</span></span>| <span data-ttu-id="f8e9e-1001">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1001">Attributes</span></span>| <span data-ttu-id="f8e9e-1002">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="f8e9e-1003">Object</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1003">Object</span></span>| <span data-ttu-id="f8e9e-1004">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-1005">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8e9e-1006">Объект</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1006">Object</span></span>| <span data-ttu-id="f8e9e-1007">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-1008">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="f8e9e-1009">функция</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1009">function</span></span>||<span data-ttu-id="f8e9e-1010">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f8e9e-1011">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8e9e-1012">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1012">Requirements</span></span>

|<span data-ttu-id="f8e9e-1013">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1013">Requirement</span></span>| <span data-ttu-id="f8e9e-1014">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-1015">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1016">1.3</span></span>|
|[<span data-ttu-id="f8e9e-1017">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1017">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8e9e-1019">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1019">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-1020">Создание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f8e9e-1021">Примеры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1021">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="f8e9e-p170">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f8e9e-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f8e9e-1025">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f8e9e-p171">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8e9e-1029">Параметры</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1029">Parameters</span></span>

|<span data-ttu-id="f8e9e-1030">Имя</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1030">Name</span></span>| <span data-ttu-id="f8e9e-1031">Тип</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1031">Type</span></span>| <span data-ttu-id="f8e9e-1032">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1032">Attributes</span></span>| <span data-ttu-id="f8e9e-1033">Описание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f8e9e-1034">String</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1034">String</span></span>||<span data-ttu-id="f8e9e-p172">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="f8e9e-1038">Object</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1038">Object</span></span>| <span data-ttu-id="f8e9e-1039">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-1040">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8e9e-1041">Объект</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1041">Object</span></span>| <span data-ttu-id="f8e9e-1042">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-1043">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="f8e9e-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="f8e9e-1045">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="f8e9e-1046">Если `text`текущий стиль применяется в Outlook для веб-клиентов и клиентов для настольных ПК.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1046">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="f8e9e-1047">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1047">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f8e9e-1048">Если `html` и поле поддерживает HTML (тема не используется), текущий стиль применяется в Outlook в Интернете, а в настольных клиентах Outlook применяется стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1048">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="f8e9e-1049">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1049">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f8e9e-1050">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="f8e9e-1051">функция</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1051">function</span></span>||<span data-ttu-id="f8e9e-1052">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8e9e-1053">Требования</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1053">Requirements</span></span>

|<span data-ttu-id="f8e9e-1054">Требование</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1054">Requirement</span></span>| <span data-ttu-id="f8e9e-1055">Значение</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8e9e-1056">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8e9e-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1057">1.2</span></span>|
|[<span data-ttu-id="f8e9e-1058">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8e9e-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8e9e-1060">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f8e9e-1061">Создание</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8e9e-1062">Пример</span><span class="sxs-lookup"><span data-stu-id="f8e9e-1062">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
