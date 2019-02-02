---
title: Office.Context.Mailbox.Item - требование задать 1.3
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: 545525a0d3c32718f063b7d249cd0a7cea2d27d5
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701906"
---
# <a name="item"></a><span data-ttu-id="5d72c-102">item</span><span class="sxs-lookup"><span data-stu-id="5d72c-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="5d72c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="5d72c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="5d72c-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="5d72c-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="5d72c-106">Requirements</span></span>

|<span data-ttu-id="5d72c-107">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-107">Requirement</span></span>| <span data-ttu-id="5d72c-108">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-110">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-110">1.0</span></span>|
|[<span data-ttu-id="5d72c-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5d72c-112">Restricted</span></span>|
|[<span data-ttu-id="5d72c-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="5d72c-115">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-115">Example</span></span>

<span data-ttu-id="5d72c-116">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="5d72c-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
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
}
```

### <a name="members"></a><span data-ttu-id="5d72c-117">Элементы</span><span class="sxs-lookup"><span data-stu-id="5d72c-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="5d72c-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5d72c-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="5d72c-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-121">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="5d72c-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="5d72c-122">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="5d72c-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-123">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-123">Type:</span></span>

*   <span data-ttu-id="5d72c-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5d72c-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-125">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-125">Requirements</span></span>

|<span data-ttu-id="5d72c-126">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-126">Requirement</span></span>| <span data-ttu-id="5d72c-127">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-128">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-129">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-129">1.0</span></span>|
|[<span data-ttu-id="5d72c-130">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-131">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-133">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-134">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-134">Example</span></span>

<span data-ttu-id="5d72c-135">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="5d72c-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d72c-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="5d72c-137">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="5d72c-138">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-139">Type:</span></span>

*   [<span data-ttu-id="5d72c-140">Recipients</span><span class="sxs-lookup"><span data-stu-id="5d72c-140">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="5d72c-141">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-141">Requirements</span></span>

|<span data-ttu-id="5d72c-142">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-142">Requirement</span></span>| <span data-ttu-id="5d72c-143">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-145">1.1</span><span class="sxs-lookup"><span data-stu-id="5d72c-145">1.1</span></span>|
|[<span data-ttu-id="5d72c-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-147">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-149">Создание</span><span class="sxs-lookup"><span data-stu-id="5d72c-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-150">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-150">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="5d72c-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="5d72c-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="5d72c-152">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-153">Type:</span></span>

*   [<span data-ttu-id="5d72c-154">Body</span><span class="sxs-lookup"><span data-stu-id="5d72c-154">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="5d72c-155">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-155">Requirements</span></span>

|<span data-ttu-id="5d72c-156">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-156">Requirement</span></span>| <span data-ttu-id="5d72c-157">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-159">1.1</span><span class="sxs-lookup"><span data-stu-id="5d72c-159">1.1</span></span>|
|[<span data-ttu-id="5d72c-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-161">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="5d72c-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d72c-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="5d72c-165">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="5d72c-166">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d72c-167">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d72c-167">Read mode</span></span>

<span data-ttu-id="5d72c-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d72c-170">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d72c-170">Compose mode</span></span>

<span data-ttu-id="5d72c-171">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-172">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-172">Type:</span></span>

*   <span data-ttu-id="5d72c-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d72c-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-174">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-174">Requirements</span></span>

|<span data-ttu-id="5d72c-175">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-175">Requirement</span></span>| <span data-ttu-id="5d72c-176">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-177">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-178">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-178">1.0</span></span>|
|[<span data-ttu-id="5d72c-179">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-180">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-181">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-182">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-183">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-183">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="5d72c-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="5d72c-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="5d72c-185">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="5d72c-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="5d72c-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="5d72c-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-190">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-190">Type:</span></span>

*   <span data-ttu-id="5d72c-191">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-192">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-192">Requirements</span></span>

|<span data-ttu-id="5d72c-193">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-193">Requirement</span></span>| <span data-ttu-id="5d72c-194">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-195">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-196">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-196">1.0</span></span>|
|[<span data-ttu-id="5d72c-197">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-198">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="5d72c-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="5d72c-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="5d72c-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-204">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-204">Type:</span></span>

*   <span data-ttu-id="5d72c-205">Date</span><span class="sxs-lookup"><span data-stu-id="5d72c-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-206">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-206">Requirements</span></span>

|<span data-ttu-id="5d72c-207">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-207">Requirement</span></span>| <span data-ttu-id="5d72c-208">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-209">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-210">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-210">1.0</span></span>|
|[<span data-ttu-id="5d72c-211">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-212">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-214">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-215">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-215">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="5d72c-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="5d72c-216">dateTimeModified :Date</span></span>

<span data-ttu-id="5d72c-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-219">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d72c-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-220">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-220">Type:</span></span>

*   <span data-ttu-id="5d72c-221">Date</span><span class="sxs-lookup"><span data-stu-id="5d72c-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-222">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-222">Requirements</span></span>

|<span data-ttu-id="5d72c-223">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-223">Requirement</span></span>| <span data-ttu-id="5d72c-224">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-225">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-226">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-226">1.0</span></span>|
|[<span data-ttu-id="5d72c-227">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-228">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-229">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-230">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-231">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-231">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="5d72c-232">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="5d72c-232">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="5d72c-233">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="5d72c-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="5d72c-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="5d72c-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d72c-236">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d72c-236">Read mode</span></span>

<span data-ttu-id="5d72c-237">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d72c-238">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d72c-238">Compose mode</span></span>

<span data-ttu-id="5d72c-239">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="5d72c-240">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5d72c-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-241">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-241">Type:</span></span>

*   <span data-ttu-id="5d72c-242">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="5d72c-242">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-243">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-243">Requirements</span></span>

|<span data-ttu-id="5d72c-244">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-244">Requirement</span></span>| <span data-ttu-id="5d72c-245">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-246">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-247">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-247">1.0</span></span>|
|[<span data-ttu-id="5d72c-248">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-249">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-250">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-251">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-252">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-252">Example</span></span>

<span data-ttu-id="5d72c-253">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="5d72c-254">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5d72c-254">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="5d72c-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="5d72c-p113">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-259">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-260">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-260">Type:</span></span>

*   [<span data-ttu-id="5d72c-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5d72c-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5d72c-262">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-262">Requirements</span></span>

|<span data-ttu-id="5d72c-263">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-263">Requirement</span></span>| <span data-ttu-id="5d72c-264">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-265">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-266">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-266">1.0</span></span>|
|[<span data-ttu-id="5d72c-267">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-268">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-269">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-270">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="5d72c-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="5d72c-271">internetMessageId :String</span></span>

<span data-ttu-id="5d72c-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-274">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-274">Type:</span></span>

*   <span data-ttu-id="5d72c-275">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-276">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-276">Requirements</span></span>

|<span data-ttu-id="5d72c-277">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-277">Requirement</span></span>| <span data-ttu-id="5d72c-278">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-279">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-280">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-280">1.0</span></span>|
|[<span data-ttu-id="5d72c-281">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-282">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-283">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-284">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-285">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-285">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="5d72c-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="5d72c-286">itemClass :String</span></span>

<span data-ttu-id="5d72c-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="5d72c-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="5d72c-291">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-291">Type</span></span> | <span data-ttu-id="5d72c-292">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-292">Description</span></span> | <span data-ttu-id="5d72c-293">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="5d72c-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="5d72c-294">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="5d72c-294">Appointment items</span></span> | <span data-ttu-id="5d72c-295">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="5d72c-296">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="5d72c-296">Message items</span></span> | <span data-ttu-id="5d72c-297">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="5d72c-298">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-299">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-299">Type:</span></span>

*   <span data-ttu-id="5d72c-300">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-301">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-301">Requirements</span></span>

|<span data-ttu-id="5d72c-302">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-302">Requirement</span></span>| <span data-ttu-id="5d72c-303">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-304">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-305">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-305">1.0</span></span>|
|[<span data-ttu-id="5d72c-306">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-307">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-308">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-309">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-310">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-310">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="5d72c-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="5d72c-311">(nullable) itemId :String</span></span>

<span data-ttu-id="5d72c-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-314">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="5d72c-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="5d72c-315">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="5d72c-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="5d72c-316">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="5d72c-316">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="5d72c-317">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="5d72c-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="5d72c-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-320">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-320">Type:</span></span>

*   <span data-ttu-id="5d72c-321">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-321">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-322">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-322">Requirements</span></span>

|<span data-ttu-id="5d72c-323">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-323">Requirement</span></span>| <span data-ttu-id="5d72c-324">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-325">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-326">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-326">1.0</span></span>|
|[<span data-ttu-id="5d72c-327">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-328">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-329">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-330">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-330">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-331">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-331">Example</span></span>

<span data-ttu-id="5d72c-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="5d72c-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="5d72c-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="5d72c-335">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="5d72c-335">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="5d72c-336">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="5d72c-336">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-337">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-337">Type:</span></span>

*   [<span data-ttu-id="5d72c-338">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="5d72c-338">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="5d72c-339">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-339">Requirements</span></span>

|<span data-ttu-id="5d72c-340">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-340">Requirement</span></span>| <span data-ttu-id="5d72c-341">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-342">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-343">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-343">1.0</span></span>|
|[<span data-ttu-id="5d72c-344">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-345">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-346">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-347">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-348">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-348">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="5d72c-349">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="5d72c-349">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="5d72c-350">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="5d72c-350">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d72c-351">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d72c-351">Read mode</span></span>

<span data-ttu-id="5d72c-352">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="5d72c-352">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d72c-353">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d72c-353">Compose mode</span></span>

<span data-ttu-id="5d72c-354">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="5d72c-354">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-355">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-355">Type:</span></span>

*   <span data-ttu-id="5d72c-356">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="5d72c-356">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-357">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-357">Requirements</span></span>

|<span data-ttu-id="5d72c-358">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-358">Requirement</span></span>| <span data-ttu-id="5d72c-359">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-360">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-361">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-361">1.0</span></span>|
|[<span data-ttu-id="5d72c-362">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-363">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-364">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-365">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-366">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-366">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="5d72c-367">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="5d72c-367">normalizedSubject :String</span></span>

<span data-ttu-id="5d72c-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="5d72c-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject).</span><span class="sxs-lookup"><span data-stu-id="5d72c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-372">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-372">Type:</span></span>

*   <span data-ttu-id="5d72c-373">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-374">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-374">Requirements</span></span>

|<span data-ttu-id="5d72c-375">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-375">Requirement</span></span>| <span data-ttu-id="5d72c-376">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-377">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-378">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-378">1.0</span></span>|
|[<span data-ttu-id="5d72c-379">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-380">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-381">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-382">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-383">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-383">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="5d72c-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="5d72c-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="5d72c-385">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-385">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-386">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-386">Type:</span></span>

*   [<span data-ttu-id="5d72c-387">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="5d72c-387">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="5d72c-388">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-388">Requirements</span></span>

|<span data-ttu-id="5d72c-389">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-389">Requirement</span></span>| <span data-ttu-id="5d72c-390">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-390">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-391">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-392">1.3</span><span class="sxs-lookup"><span data-stu-id="5d72c-392">1.3</span></span>|
|[<span data-ttu-id="5d72c-393">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-393">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-394">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-395">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-395">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-396">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-396">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="5d72c-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d72c-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="5d72c-398">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5d72c-398">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="5d72c-399">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-399">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d72c-400">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d72c-400">Read mode</span></span>

<span data-ttu-id="5d72c-401">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-401">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d72c-402">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d72c-402">Compose mode</span></span>

<span data-ttu-id="5d72c-403">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-403">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-404">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-404">Type:</span></span>

*   <span data-ttu-id="5d72c-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d72c-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-406">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-406">Requirements</span></span>

|<span data-ttu-id="5d72c-407">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-407">Requirement</span></span>| <span data-ttu-id="5d72c-408">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-409">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-410">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-410">1.0</span></span>|
|[<span data-ttu-id="5d72c-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-412">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-414">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-414">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-415">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-415">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="5d72c-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5d72c-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="5d72c-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-419">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-419">Type:</span></span>

*   [<span data-ttu-id="5d72c-420">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5d72c-420">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5d72c-421">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-421">Requirements</span></span>

|<span data-ttu-id="5d72c-422">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-422">Requirement</span></span>| <span data-ttu-id="5d72c-423">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-424">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-425">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-425">1.0</span></span>|
|[<span data-ttu-id="5d72c-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-427">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-429">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-430">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-430">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="5d72c-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d72c-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="5d72c-432">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5d72c-432">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="5d72c-433">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-433">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d72c-434">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d72c-434">Read mode</span></span>

<span data-ttu-id="5d72c-435">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-435">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d72c-436">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d72c-436">Compose mode</span></span>

<span data-ttu-id="5d72c-437">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-437">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-438">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-438">Type:</span></span>

*   <span data-ttu-id="5d72c-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d72c-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-440">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-440">Requirements</span></span>

|<span data-ttu-id="5d72c-441">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-441">Requirement</span></span>| <span data-ttu-id="5d72c-442">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-443">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-444">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-444">1.0</span></span>|
|[<span data-ttu-id="5d72c-445">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-446">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-447">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-448">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-449">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-449">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="5d72c-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5d72c-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="5d72c-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="5d72c-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-455">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-455">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-456">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-456">Type:</span></span>

*   [<span data-ttu-id="5d72c-457">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5d72c-457">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5d72c-458">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-458">Requirements</span></span>

|<span data-ttu-id="5d72c-459">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-459">Requirement</span></span>| <span data-ttu-id="5d72c-460">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-461">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-461">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-462">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-462">1.0</span></span>|
|[<span data-ttu-id="5d72c-463">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-464">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-465">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-466">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-466">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-467">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-467">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="5d72c-468">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="5d72c-468">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="5d72c-469">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="5d72c-469">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="5d72c-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="5d72c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d72c-472">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d72c-472">Read mode</span></span>

<span data-ttu-id="5d72c-473">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-473">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d72c-474">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d72c-474">Compose mode</span></span>

<span data-ttu-id="5d72c-475">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-475">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="5d72c-476">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5d72c-476">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-477">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-477">Type:</span></span>

*   <span data-ttu-id="5d72c-478">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="5d72c-478">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-479">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-479">Requirements</span></span>

|<span data-ttu-id="5d72c-480">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-480">Requirement</span></span>| <span data-ttu-id="5d72c-481">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-482">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-482">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-483">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-483">1.0</span></span>|
|[<span data-ttu-id="5d72c-484">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-485">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-486">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-487">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-487">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-488">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-488">Example</span></span>

<span data-ttu-id="5d72c-489">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-489">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="5d72c-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5d72c-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="5d72c-491">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="5d72c-492">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="5d72c-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d72c-493">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d72c-493">Read mode</span></span>

<span data-ttu-id="5d72c-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="5d72c-496">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d72c-496">Compose mode</span></span>

<span data-ttu-id="5d72c-497">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="5d72c-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5d72c-498">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-498">Type:</span></span>

*   <span data-ttu-id="5d72c-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5d72c-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-500">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-500">Requirements</span></span>

|<span data-ttu-id="5d72c-501">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-501">Requirement</span></span>| <span data-ttu-id="5d72c-502">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-503">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-504">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-504">1.0</span></span>|
|[<span data-ttu-id="5d72c-505">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-506">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-507">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-508">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-508">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="5d72c-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d72c-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="5d72c-510">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="5d72c-511">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5d72c-512">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5d72c-512">Read mode</span></span>

<span data-ttu-id="5d72c-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5d72c-515">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5d72c-515">Compose mode</span></span>

<span data-ttu-id="5d72c-516">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="5d72c-517">Тип:</span><span class="sxs-lookup"><span data-stu-id="5d72c-517">Type:</span></span>

*   <span data-ttu-id="5d72c-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5d72c-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-519">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-519">Requirements</span></span>

|<span data-ttu-id="5d72c-520">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-520">Requirement</span></span>| <span data-ttu-id="5d72c-521">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-522">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-523">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-523">1.0</span></span>|
|[<span data-ttu-id="5d72c-524">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-525">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-526">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-527">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-528">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-528">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="5d72c-529">Методы</span><span class="sxs-lookup"><span data-stu-id="5d72c-529">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="5d72c-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5d72c-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5d72c-531">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-531">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="5d72c-532">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-532">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="5d72c-533">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5d72c-533">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-534">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d72c-534">Parameters:</span></span>

|<span data-ttu-id="5d72c-535">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-535">Name</span></span>| <span data-ttu-id="5d72c-536">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-536">Type</span></span>| <span data-ttu-id="5d72c-537">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d72c-537">Attributes</span></span>| <span data-ttu-id="5d72c-538">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-538">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="5d72c-539">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-539">String</span></span>||<span data-ttu-id="5d72c-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5d72c-542">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-542">String</span></span>||<span data-ttu-id="5d72c-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5d72c-545">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-545">Object</span></span>| <span data-ttu-id="5d72c-546">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-546">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-547">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d72c-547">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d72c-548">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-548">Object</span></span>| <span data-ttu-id="5d72c-549">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-549">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-550">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d72c-550">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5d72c-551">функция</span><span class="sxs-lookup"><span data-stu-id="5d72c-551">function</span></span>| <span data-ttu-id="5d72c-552">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-552">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-553">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d72c-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5d72c-554">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-554">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5d72c-555">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5d72c-555">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5d72c-556">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5d72c-556">Errors</span></span>

| <span data-ttu-id="5d72c-557">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5d72c-557">Error code</span></span> | <span data-ttu-id="5d72c-558">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-558">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="5d72c-559">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="5d72c-559">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="5d72c-560">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5d72c-560">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5d72c-561">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5d72c-561">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d72c-562">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-562">Requirements</span></span>

|<span data-ttu-id="5d72c-563">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-563">Requirement</span></span>| <span data-ttu-id="5d72c-564">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-565">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-566">1.1</span><span class="sxs-lookup"><span data-stu-id="5d72c-566">1.1</span></span>|
|[<span data-ttu-id="5d72c-567">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-568">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-568">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d72c-569">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-570">Создание</span><span class="sxs-lookup"><span data-stu-id="5d72c-570">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-571">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-571">Example</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="5d72c-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5d72c-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5d72c-573">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-573">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="5d72c-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="5d72c-577">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5d72c-577">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="5d72c-578">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5d72c-578">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-579">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5d72c-579">Parameters:</span></span>

|<span data-ttu-id="5d72c-580">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-580">Name</span></span>| <span data-ttu-id="5d72c-581">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-581">Type</span></span>| <span data-ttu-id="5d72c-582">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d72c-582">Attributes</span></span>| <span data-ttu-id="5d72c-583">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-583">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="5d72c-584">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-584">String</span></span>||<span data-ttu-id="5d72c-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5d72c-587">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-587">String</span></span>||<span data-ttu-id="5d72c-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5d72c-590">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-590">Object</span></span>| <span data-ttu-id="5d72c-591">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-591">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-592">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d72c-592">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d72c-593">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-593">Object</span></span>| <span data-ttu-id="5d72c-594">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-594">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-595">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d72c-595">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5d72c-596">функция</span><span class="sxs-lookup"><span data-stu-id="5d72c-596">function</span></span>| <span data-ttu-id="5d72c-597">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-597">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-598">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d72c-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5d72c-599">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-599">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5d72c-600">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5d72c-600">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5d72c-601">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5d72c-601">Errors</span></span>

| <span data-ttu-id="5d72c-602">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5d72c-602">Error code</span></span> | <span data-ttu-id="5d72c-603">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-603">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5d72c-604">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5d72c-604">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d72c-605">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-605">Requirements</span></span>

|<span data-ttu-id="5d72c-606">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-606">Requirement</span></span>| <span data-ttu-id="5d72c-607">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-608">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-609">1.1</span><span class="sxs-lookup"><span data-stu-id="5d72c-609">1.1</span></span>|
|[<span data-ttu-id="5d72c-610">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d72c-612">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-613">Создание</span><span class="sxs-lookup"><span data-stu-id="5d72c-613">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-614">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-614">Example</span></span>

<span data-ttu-id="5d72c-615">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-615">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="5d72c-616">close()</span><span class="sxs-lookup"><span data-stu-id="5d72c-616">close()</span></span>

<span data-ttu-id="5d72c-617">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="5d72c-617">Closes the current item that is being composed.</span></span>

<span data-ttu-id="5d72c-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-620">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-620">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="5d72c-621">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="5d72c-621">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-622">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-622">Requirements</span></span>

|<span data-ttu-id="5d72c-623">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-623">Requirement</span></span>| <span data-ttu-id="5d72c-624">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-624">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-625">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-625">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-626">1.3</span><span class="sxs-lookup"><span data-stu-id="5d72c-626">1.3</span></span>|
|[<span data-ttu-id="5d72c-627">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-627">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-628">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5d72c-628">Restricted</span></span>|
|[<span data-ttu-id="5d72c-629">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-629">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-630">Создание</span><span class="sxs-lookup"><span data-stu-id="5d72c-630">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="5d72c-631">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="5d72c-631">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="5d72c-632">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5d72c-632">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-633">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d72c-633">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d72c-634">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5d72c-634">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5d72c-635">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5d72c-635">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="5d72c-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-639">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5d72c-639">Parameters:</span></span>

|<span data-ttu-id="5d72c-640">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-640">Name</span></span>| <span data-ttu-id="5d72c-641">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-641">Type</span></span>| <span data-ttu-id="5d72c-642">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-642">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="5d72c-643">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-643">String &#124; Object</span></span>| |<span data-ttu-id="5d72c-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5d72c-646">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5d72c-646">**OR**</span></span><br/><span data-ttu-id="5d72c-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5d72c-649">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-649">String</span></span> | <span data-ttu-id="5d72c-650">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-650">&lt;optional&gt;</span></span> | <span data-ttu-id="5d72c-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5d72c-653">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-653">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5d72c-654">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-654">&lt;optional&gt;</span></span> | <span data-ttu-id="5d72c-655">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5d72c-655">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5d72c-656">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-656">String</span></span> | | <span data-ttu-id="5d72c-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5d72c-659">Строка</span><span class="sxs-lookup"><span data-stu-id="5d72c-659">String</span></span> | | <span data-ttu-id="5d72c-660">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-660">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5d72c-661">Строка</span><span class="sxs-lookup"><span data-stu-id="5d72c-661">String</span></span> | | <span data-ttu-id="5d72c-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5d72c-664">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-664">String</span></span> | | <span data-ttu-id="5d72c-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5d72c-668">функция</span><span class="sxs-lookup"><span data-stu-id="5d72c-668">function</span></span> | <span data-ttu-id="5d72c-669">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-669">&lt;optional&gt;</span></span> | <span data-ttu-id="5d72c-670">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d72c-670">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d72c-671">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-671">Requirements</span></span>

|<span data-ttu-id="5d72c-672">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-672">Requirement</span></span>| <span data-ttu-id="5d72c-673">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-673">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-674">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-674">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-675">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-675">1.0</span></span>|
|[<span data-ttu-id="5d72c-676">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-676">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-677">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-677">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-678">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-678">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-679">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-679">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5d72c-680">Примеры</span><span class="sxs-lookup"><span data-stu-id="5d72c-680">Examples</span></span>

<span data-ttu-id="5d72c-681">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-681">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="5d72c-682">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-682">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="5d72c-683">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-683">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5d72c-684">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5d72c-684">Reply with a body and a file attachment.</span></span>

```js
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

<span data-ttu-id="5d72c-685">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5d72c-685">Reply with a body and an item attachment.</span></span>

```js
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

<span data-ttu-id="5d72c-686">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5d72c-686">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="5d72c-687">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="5d72c-687">displayReplyForm(formData)</span></span>

<span data-ttu-id="5d72c-688">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5d72c-688">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-689">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d72c-689">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d72c-690">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5d72c-690">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5d72c-691">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5d72c-691">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="5d72c-p145">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-695">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5d72c-695">Parameters:</span></span>

|<span data-ttu-id="5d72c-696">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-696">Name</span></span>| <span data-ttu-id="5d72c-697">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-697">Type</span></span>| <span data-ttu-id="5d72c-698">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-698">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="5d72c-699">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-699">String &#124; Object</span></span>| | <span data-ttu-id="5d72c-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5d72c-702">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5d72c-702">**OR**</span></span><br/><span data-ttu-id="5d72c-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5d72c-705">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-705">String</span></span> | <span data-ttu-id="5d72c-706">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-706">&lt;optional&gt;</span></span> | <span data-ttu-id="5d72c-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5d72c-709">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-709">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5d72c-710">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-710">&lt;optional&gt;</span></span> | <span data-ttu-id="5d72c-711">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5d72c-711">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5d72c-712">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-712">String</span></span> | | <span data-ttu-id="5d72c-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5d72c-715">Строка</span><span class="sxs-lookup"><span data-stu-id="5d72c-715">String</span></span> | | <span data-ttu-id="5d72c-716">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-716">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5d72c-717">Строка</span><span class="sxs-lookup"><span data-stu-id="5d72c-717">String</span></span> | | <span data-ttu-id="5d72c-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5d72c-720">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-720">String</span></span> | | <span data-ttu-id="5d72c-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5d72c-724">function</span><span class="sxs-lookup"><span data-stu-id="5d72c-724">function</span></span> | <span data-ttu-id="5d72c-725">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-725">&lt;optional&gt;</span></span> | <span data-ttu-id="5d72c-726">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d72c-726">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d72c-727">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-727">Requirements</span></span>

|<span data-ttu-id="5d72c-728">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-728">Requirement</span></span>| <span data-ttu-id="5d72c-729">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-729">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-730">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-730">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-731">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-731">1.0</span></span>|
|[<span data-ttu-id="5d72c-732">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-732">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-733">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-733">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-734">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-734">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-735">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-735">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5d72c-736">Примеры</span><span class="sxs-lookup"><span data-stu-id="5d72c-736">Examples</span></span>

<span data-ttu-id="5d72c-737">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-737">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="5d72c-738">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-738">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="5d72c-739">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-739">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5d72c-740">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5d72c-740">Reply with a body and a file attachment.</span></span>

```js
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

<span data-ttu-id="5d72c-741">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5d72c-741">Reply with a body and an item attachment.</span></span>

```js
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

<span data-ttu-id="5d72c-742">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5d72c-742">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="5d72c-743">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="5d72c-743">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="5d72c-744">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-744">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-745">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d72c-745">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-746">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-746">Requirements</span></span>

|<span data-ttu-id="5d72c-747">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-747">Requirement</span></span>| <span data-ttu-id="5d72c-748">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-749">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-749">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-750">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-750">1.0</span></span>|
|[<span data-ttu-id="5d72c-751">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-752">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-753">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-754">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d72c-755">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d72c-755">Returns:</span></span>

<span data-ttu-id="5d72c-756">Тип: [Entities](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="5d72c-756">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="5d72c-757">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-757">Example</span></span>

<span data-ttu-id="5d72c-758">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-758">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="5d72c-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5d72c-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5d72c-760">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-760">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-761">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d72c-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-762">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d72c-762">Parameters:</span></span>

|<span data-ttu-id="5d72c-763">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-763">Name</span></span>| <span data-ttu-id="5d72c-764">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-764">Type</span></span>| <span data-ttu-id="5d72c-765">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-765">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="5d72c-766">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="5d72c-766">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="5d72c-767">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="5d72c-767">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d72c-768">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-768">Requirements</span></span>

|<span data-ttu-id="5d72c-769">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-769">Requirement</span></span>| <span data-ttu-id="5d72c-770">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-771">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-772">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-772">1.0</span></span>|
|[<span data-ttu-id="5d72c-773">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-773">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-774">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5d72c-774">Restricted</span></span>|
|[<span data-ttu-id="5d72c-775">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-775">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-776">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d72c-777">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d72c-777">Returns:</span></span>

<span data-ttu-id="5d72c-778">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="5d72c-778">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="5d72c-779">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5d72c-779">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="5d72c-780">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-780">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="5d72c-781">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="5d72c-781">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="5d72c-782">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="5d72c-782">Value of `entityType`</span></span> | <span data-ttu-id="5d72c-783">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="5d72c-783">Type of objects in returned array</span></span> | <span data-ttu-id="5d72c-784">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-784">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="5d72c-785">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-785">String</span></span> | <span data-ttu-id="5d72c-786">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5d72c-786">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="5d72c-787">Contact</span><span class="sxs-lookup"><span data-stu-id="5d72c-787">Contact</span></span> | <span data-ttu-id="5d72c-788">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5d72c-788">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="5d72c-789">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-789">String</span></span> | <span data-ttu-id="5d72c-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5d72c-790">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="5d72c-791">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="5d72c-791">MeetingSuggestion</span></span> | <span data-ttu-id="5d72c-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5d72c-792">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="5d72c-793">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="5d72c-793">PhoneNumber</span></span> | <span data-ttu-id="5d72c-794">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5d72c-794">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="5d72c-795">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="5d72c-795">TaskSuggestion</span></span> | <span data-ttu-id="5d72c-796">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5d72c-796">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="5d72c-797">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-797">String</span></span> | <span data-ttu-id="5d72c-798">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5d72c-798">**Restricted**</span></span> |

<span data-ttu-id="5d72c-799">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5d72c-799">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="5d72c-800">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-800">Example</span></span>

<span data-ttu-id="5d72c-801">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-801">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="5d72c-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5d72c-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5d72c-803">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5d72c-803">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-804">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d72c-804">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d72c-805">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-805">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-806">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5d72c-806">Parameters:</span></span>

|<span data-ttu-id="5d72c-807">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-807">Name</span></span>| <span data-ttu-id="5d72c-808">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-808">Type</span></span>| <span data-ttu-id="5d72c-809">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-809">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5d72c-810">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-810">String</span></span>|<span data-ttu-id="5d72c-811">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5d72c-811">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d72c-812">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-812">Requirements</span></span>

|<span data-ttu-id="5d72c-813">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-813">Requirement</span></span>| <span data-ttu-id="5d72c-814">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-814">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-815">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-815">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-816">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-816">1.0</span></span>|
|[<span data-ttu-id="5d72c-817">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-817">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-818">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-818">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-819">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-819">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-820">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-820">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d72c-821">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d72c-821">Returns:</span></span>

<span data-ttu-id="5d72c-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="5d72c-824">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5d72c-824">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="5d72c-825">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="5d72c-825">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="5d72c-826">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5d72c-826">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-827">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d72c-827">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d72c-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="5d72c-831">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-831">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="5d72c-832">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-832">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="5d72c-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d72c-836">Requirements</span><span class="sxs-lookup"><span data-stu-id="5d72c-836">Requirements</span></span>

|<span data-ttu-id="5d72c-837">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-837">Requirement</span></span>| <span data-ttu-id="5d72c-838">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-839">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-840">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-840">1.0</span></span>|
|[<span data-ttu-id="5d72c-841">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-842">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-843">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-844">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d72c-845">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d72c-845">Returns:</span></span>

<span data-ttu-id="5d72c-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="5d72c-848">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5d72c-848">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5d72c-849">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-849">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5d72c-850">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-850">Example</span></span>

<span data-ttu-id="5d72c-851">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="5d72c-851">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="5d72c-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="5d72c-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="5d72c-853">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5d72c-853">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-854">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5d72c-854">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5d72c-855">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-855">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="5d72c-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-858">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5d72c-858">Parameters:</span></span>

|<span data-ttu-id="5d72c-859">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-859">Name</span></span>| <span data-ttu-id="5d72c-860">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-860">Type</span></span>| <span data-ttu-id="5d72c-861">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-861">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5d72c-862">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-862">String</span></span>|<span data-ttu-id="5d72c-863">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5d72c-863">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d72c-864">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-864">Requirements</span></span>

|<span data-ttu-id="5d72c-865">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-865">Requirement</span></span>| <span data-ttu-id="5d72c-866">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-866">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-867">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-867">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-868">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-868">1.0</span></span>|
|[<span data-ttu-id="5d72c-869">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-869">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-870">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-870">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-871">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-871">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-872">Чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-872">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d72c-873">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d72c-873">Returns:</span></span>

<span data-ttu-id="5d72c-874">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5d72c-874">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="5d72c-875">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5d72c-875">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5d72c-876">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="5d72c-876">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5d72c-877">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-877">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="5d72c-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="5d72c-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="5d72c-879">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-879">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="5d72c-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-882">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5d72c-882">Parameters:</span></span>

|<span data-ttu-id="5d72c-883">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-883">Name</span></span>| <span data-ttu-id="5d72c-884">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-884">Type</span></span>| <span data-ttu-id="5d72c-885">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d72c-885">Attributes</span></span>| <span data-ttu-id="5d72c-886">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-886">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="5d72c-887">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5d72c-887">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="5d72c-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="5d72c-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="5d72c-891">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-891">Object</span></span>| <span data-ttu-id="5d72c-892">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-892">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-893">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d72c-893">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d72c-894">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-894">Object</span></span>| <span data-ttu-id="5d72c-895">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-895">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-896">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d72c-896">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5d72c-897">функция</span><span class="sxs-lookup"><span data-stu-id="5d72c-897">function</span></span>||<span data-ttu-id="5d72c-898">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d72c-898">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5d72c-899">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-899">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="5d72c-900">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-900">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d72c-901">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-901">Requirements</span></span>

|<span data-ttu-id="5d72c-902">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-902">Requirement</span></span>| <span data-ttu-id="5d72c-903">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-904">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-905">1.2</span><span class="sxs-lookup"><span data-stu-id="5d72c-905">1.2</span></span>|
|[<span data-ttu-id="5d72c-906">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-906">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-907">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-907">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d72c-908">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-908">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-909">Создание</span><span class="sxs-lookup"><span data-stu-id="5d72c-909">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="5d72c-910">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5d72c-910">Returns:</span></span>

<span data-ttu-id="5d72c-911">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-911">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="5d72c-912">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5d72c-912">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5d72c-913">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-913">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5d72c-914">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-914">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="5d72c-915">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5d72c-915">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="5d72c-916">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-916">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="5d72c-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-920">Параметры</span><span class="sxs-lookup"><span data-stu-id="5d72c-920">Parameters:</span></span>

|<span data-ttu-id="5d72c-921">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-921">Name</span></span>| <span data-ttu-id="5d72c-922">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-922">Type</span></span>| <span data-ttu-id="5d72c-923">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d72c-923">Attributes</span></span>| <span data-ttu-id="5d72c-924">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-924">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5d72c-925">функция</span><span class="sxs-lookup"><span data-stu-id="5d72c-925">function</span></span>||<span data-ttu-id="5d72c-926">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d72c-926">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5d72c-927">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-927">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="5d72c-928">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="5d72c-928">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="5d72c-929">Объект</span><span class="sxs-lookup"><span data-stu-id="5d72c-929">Object</span></span>| <span data-ttu-id="5d72c-930">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-930">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-931">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d72c-931">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="5d72c-932">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d72c-932">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d72c-933">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-933">Requirements</span></span>

|<span data-ttu-id="5d72c-934">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-934">Requirement</span></span>| <span data-ttu-id="5d72c-935">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-936">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-937">1.0</span><span class="sxs-lookup"><span data-stu-id="5d72c-937">1.0</span></span>|
|[<span data-ttu-id="5d72c-938">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-939">ReadItem</span></span>|
|[<span data-ttu-id="5d72c-940">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-941">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5d72c-941">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-942">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-942">Example</span></span>

<span data-ttu-id="5d72c-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="5d72c-946">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5d72c-946">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="5d72c-947">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5d72c-947">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="5d72c-p165">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-952">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5d72c-952">Parameters:</span></span>

|<span data-ttu-id="5d72c-953">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-953">Name</span></span>| <span data-ttu-id="5d72c-954">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-954">Type</span></span>| <span data-ttu-id="5d72c-955">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d72c-955">Attributes</span></span>| <span data-ttu-id="5d72c-956">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-956">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="5d72c-957">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-957">String</span></span>||<span data-ttu-id="5d72c-958">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-958">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="5d72c-959">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-959">Object</span></span>| <span data-ttu-id="5d72c-960">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-960">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-961">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d72c-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d72c-962">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-962">Object</span></span>| <span data-ttu-id="5d72c-963">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-963">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-964">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d72c-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5d72c-965">функция</span><span class="sxs-lookup"><span data-stu-id="5d72c-965">function</span></span>| <span data-ttu-id="5d72c-966">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-966">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-967">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d72c-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5d72c-968">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="5d72c-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5d72c-969">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5d72c-969">Errors</span></span>

| <span data-ttu-id="5d72c-970">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5d72c-970">Error code</span></span> | <span data-ttu-id="5d72c-971">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="5d72c-972">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="5d72c-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d72c-973">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-973">Requirements</span></span>

|<span data-ttu-id="5d72c-974">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-974">Requirement</span></span>| <span data-ttu-id="5d72c-975">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-976">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5d72c-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-977">1.1</span><span class="sxs-lookup"><span data-stu-id="5d72c-977">1.1</span></span>|
|[<span data-ttu-id="5d72c-978">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d72c-980">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-981">Создание</span><span class="sxs-lookup"><span data-stu-id="5d72c-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-982">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-982">Example</span></span>

<span data-ttu-id="5d72c-983">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="5d72c-983">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="5d72c-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="5d72c-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="5d72c-985">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="5d72c-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="5d72c-p166">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-989">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="5d72c-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="5d72c-990">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="5d72c-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="5d72c-p168">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="5d72c-994">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="5d72c-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="5d72c-995">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="5d72c-996">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="5d72c-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="5d72c-997">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5d72c-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-998">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5d72c-998">Parameters:</span></span>

|<span data-ttu-id="5d72c-999">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-999">Name</span></span>| <span data-ttu-id="5d72c-1000">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-1000">Type</span></span>| <span data-ttu-id="5d72c-1001">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d72c-1001">Attributes</span></span>| <span data-ttu-id="5d72c-1002">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="5d72c-1003">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-1003">Object</span></span>| <span data-ttu-id="5d72c-1004">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-1005">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d72c-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d72c-1006">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-1006">Object</span></span>| <span data-ttu-id="5d72c-1007">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-1008">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5d72c-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5d72c-1009">функция</span><span class="sxs-lookup"><span data-stu-id="5d72c-1009">function</span></span>||<span data-ttu-id="5d72c-1010">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d72c-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5d72c-1011">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d72c-1012">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-1012">Requirements</span></span>

|<span data-ttu-id="5d72c-1013">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-1013">Requirement</span></span>| <span data-ttu-id="5d72c-1014">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-1015">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="5d72c-1016">1.3</span></span>|
|[<span data-ttu-id="5d72c-1017">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d72c-1019">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-1020">Создание</span><span class="sxs-lookup"><span data-stu-id="5d72c-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="5d72c-1021">Примеры</span><span class="sxs-lookup"><span data-stu-id="5d72c-1021">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="5d72c-p170">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="5d72c-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="5d72c-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="5d72c-1025">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="5d72c-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="5d72c-p171">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5d72c-1029">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5d72c-1029">Parameters:</span></span>

|<span data-ttu-id="5d72c-1030">Имя</span><span class="sxs-lookup"><span data-stu-id="5d72c-1030">Name</span></span>| <span data-ttu-id="5d72c-1031">Тип</span><span class="sxs-lookup"><span data-stu-id="5d72c-1031">Type</span></span>| <span data-ttu-id="5d72c-1032">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5d72c-1032">Attributes</span></span>| <span data-ttu-id="5d72c-1033">Описание</span><span class="sxs-lookup"><span data-stu-id="5d72c-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="5d72c-1034">String</span><span class="sxs-lookup"><span data-stu-id="5d72c-1034">String</span></span>||<span data-ttu-id="5d72c-p172">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="5d72c-1038">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-1038">Object</span></span>| <span data-ttu-id="5d72c-1039">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-1040">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5d72c-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5d72c-1041">Object</span><span class="sxs-lookup"><span data-stu-id="5d72c-1041">Object</span></span>| <span data-ttu-id="5d72c-1042">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-1043">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="5d72c-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="5d72c-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5d72c-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="5d72c-1045">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5d72c-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="5d72c-p173">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="5d72c-p174">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="5d72c-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="5d72c-1050">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="5d72c-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="5d72c-1051">функция</span><span class="sxs-lookup"><span data-stu-id="5d72c-1051">function</span></span>||<span data-ttu-id="5d72c-1052">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5d72c-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d72c-1053">Требования</span><span class="sxs-lookup"><span data-stu-id="5d72c-1053">Requirements</span></span>

|<span data-ttu-id="5d72c-1054">Требование</span><span class="sxs-lookup"><span data-stu-id="5d72c-1054">Requirement</span></span>| <span data-ttu-id="5d72c-1055">Значение</span><span class="sxs-lookup"><span data-stu-id="5d72c-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d72c-1056">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5d72c-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d72c-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="5d72c-1057">1.2</span></span>|
|[<span data-ttu-id="5d72c-1058">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5d72c-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d72c-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5d72c-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="5d72c-1060">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5d72c-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d72c-1061">Создание</span><span class="sxs-lookup"><span data-stu-id="5d72c-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5d72c-1062">Пример</span><span class="sxs-lookup"><span data-stu-id="5d72c-1062">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
