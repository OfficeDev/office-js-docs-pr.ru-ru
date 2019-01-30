---
title: Office.Context.Mailbox.Item - требование задать 1.2 (en)
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: d58a38ce045a179a7e5cdd2e15b4e16c2ac03c91
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388600"
---
# <a name="item"></a><span data-ttu-id="6bad0-102">item</span><span class="sxs-lookup"><span data-stu-id="6bad0-102">item</span></span>

### <span data-ttu-id="6bad0-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="6bad0-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="6bad0-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="6bad0-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="6bad0-107">Requirements</span></span>

|<span data-ttu-id="6bad0-108">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-108">Requirement</span></span>| <span data-ttu-id="6bad0-109">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-111">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-111">1.0</span></span>|
|[<span data-ttu-id="6bad0-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="6bad0-113">Restricted</span></span>|
|[<span data-ttu-id="6bad0-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="6bad0-116">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-116">Example</span></span>

<span data-ttu-id="6bad0-117">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="6bad0-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="6bad0-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="6bad0-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="6bad0-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="6bad0-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="6bad0-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-122">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="6bad0-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="6bad0-123">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="6bad0-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-124">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-124">Type:</span></span>

*   <span data-ttu-id="6bad0-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="6bad0-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-126">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-126">Requirements</span></span>

|<span data-ttu-id="6bad0-127">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-127">Requirement</span></span>| <span data-ttu-id="6bad0-128">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-129">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-130">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-130">1.0</span></span>|
|[<span data-ttu-id="6bad0-131">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-132">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-134">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-135">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-135">Example</span></span>

<span data-ttu-id="6bad0-136">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6bad0-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6bad0-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6bad0-138">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="6bad0-139">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="6bad0-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-140">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-140">Type:</span></span>

*   [<span data-ttu-id="6bad0-141">Recipients</span><span class="sxs-lookup"><span data-stu-id="6bad0-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="6bad0-142">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-142">Requirements</span></span>

|<span data-ttu-id="6bad0-143">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-143">Requirement</span></span>| <span data-ttu-id="6bad0-144">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-145">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-146">1.1</span><span class="sxs-lookup"><span data-stu-id="6bad0-146">1.1</span></span>|
|[<span data-ttu-id="6bad0-147">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-148">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-149">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-150">Создание</span><span class="sxs-lookup"><span data-stu-id="6bad0-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-151">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="6bad0-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="6bad0-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="6bad0-153">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-154">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-154">Type:</span></span>

*   [<span data-ttu-id="6bad0-155">Body</span><span class="sxs-lookup"><span data-stu-id="6bad0-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="6bad0-156">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-156">Requirements</span></span>

|<span data-ttu-id="6bad0-157">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-157">Requirement</span></span>| <span data-ttu-id="6bad0-158">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-160">1.1</span><span class="sxs-lookup"><span data-stu-id="6bad0-160">1.1</span></span>|
|[<span data-ttu-id="6bad0-161">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-162">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6bad0-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6bad0-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6bad0-166">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="6bad0-167">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6bad0-168">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="6bad0-168">Read mode</span></span>

<span data-ttu-id="6bad0-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6bad0-171">Режим создания</span><span class="sxs-lookup"><span data-stu-id="6bad0-171">Compose mode</span></span>

<span data-ttu-id="6bad0-172">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-173">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-173">Type:</span></span>

*   <span data-ttu-id="6bad0-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6bad0-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-175">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-175">Requirements</span></span>

|<span data-ttu-id="6bad0-176">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-176">Requirement</span></span>| <span data-ttu-id="6bad0-177">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-178">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-179">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-179">1.0</span></span>|
|[<span data-ttu-id="6bad0-180">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-181">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-182">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-183">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-184">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="6bad0-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="6bad0-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="6bad0-186">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="6bad0-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="6bad0-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="6bad0-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-191">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-191">Type:</span></span>

*   <span data-ttu-id="6bad0-192">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-193">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-193">Requirements</span></span>

|<span data-ttu-id="6bad0-194">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-194">Requirement</span></span>| <span data-ttu-id="6bad0-195">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-196">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-197">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-197">1.0</span></span>|
|[<span data-ttu-id="6bad0-198">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-199">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-200">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-201">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="6bad0-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="6bad0-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="6bad0-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-205">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-205">Type:</span></span>

*   <span data-ttu-id="6bad0-206">Date</span><span class="sxs-lookup"><span data-stu-id="6bad0-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-207">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-207">Requirements</span></span>

|<span data-ttu-id="6bad0-208">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-208">Requirement</span></span>| <span data-ttu-id="6bad0-209">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-210">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-211">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-211">1.0</span></span>|
|[<span data-ttu-id="6bad0-212">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-213">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-214">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-215">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-216">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="6bad0-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="6bad0-217">dateTimeModified :Date</span></span>

<span data-ttu-id="6bad0-p111">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-220">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="6bad0-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-221">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-221">Type:</span></span>

*   <span data-ttu-id="6bad0-222">Date</span><span class="sxs-lookup"><span data-stu-id="6bad0-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-223">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-223">Requirements</span></span>

|<span data-ttu-id="6bad0-224">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-224">Requirement</span></span>| <span data-ttu-id="6bad0-225">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-226">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-227">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-227">1.0</span></span>|
|[<span data-ttu-id="6bad0-228">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-229">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-230">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-231">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-232">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="6bad0-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="6bad0-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="6bad0-234">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="6bad0-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="6bad0-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="6bad0-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6bad0-237">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="6bad0-237">Read mode</span></span>

<span data-ttu-id="6bad0-238">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6bad0-239">Режим создания</span><span class="sxs-lookup"><span data-stu-id="6bad0-239">Compose mode</span></span>

<span data-ttu-id="6bad0-240">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="6bad0-241">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="6bad0-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-242">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-242">Type:</span></span>

*   <span data-ttu-id="6bad0-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="6bad0-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-244">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-244">Requirements</span></span>

|<span data-ttu-id="6bad0-245">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-245">Requirement</span></span>| <span data-ttu-id="6bad0-246">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-247">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-248">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-248">1.0</span></span>|
|[<span data-ttu-id="6bad0-249">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-250">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-251">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-252">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-253">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-253">Example</span></span>

<span data-ttu-id="6bad0-254">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="6bad0-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="6bad0-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6bad0-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="6bad0-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="6bad0-p114">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-260">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-261">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-261">Type:</span></span>

*   [<span data-ttu-id="6bad0-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6bad0-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6bad0-263">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-263">Requirements</span></span>

|<span data-ttu-id="6bad0-264">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-264">Requirement</span></span>| <span data-ttu-id="6bad0-265">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-267">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-267">1.0</span></span>|
|[<span data-ttu-id="6bad0-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-269">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-271">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="6bad0-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="6bad0-272">internetMessageId :String</span></span>

<span data-ttu-id="6bad0-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-275">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-275">Type:</span></span>

*   <span data-ttu-id="6bad0-276">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-277">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-277">Requirements</span></span>

|<span data-ttu-id="6bad0-278">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-278">Requirement</span></span>| <span data-ttu-id="6bad0-279">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-280">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-281">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-281">1.0</span></span>|
|[<span data-ttu-id="6bad0-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-283">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-285">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-286">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="6bad0-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="6bad0-287">itemClass :String</span></span>

<span data-ttu-id="6bad0-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="6bad0-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="6bad0-292">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-292">Type</span></span> | <span data-ttu-id="6bad0-293">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-293">Description</span></span> | <span data-ttu-id="6bad0-294">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="6bad0-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="6bad0-295">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="6bad0-295">Appointment items</span></span> | <span data-ttu-id="6bad0-296">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="6bad0-297">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="6bad0-297">Message items</span></span> | <span data-ttu-id="6bad0-298">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="6bad0-299">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-300">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-300">Type:</span></span>

*   <span data-ttu-id="6bad0-301">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-302">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-302">Requirements</span></span>

|<span data-ttu-id="6bad0-303">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-303">Requirement</span></span>| <span data-ttu-id="6bad0-304">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-305">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-306">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-306">1.0</span></span>|
|[<span data-ttu-id="6bad0-307">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-308">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-309">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-310">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-311">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="6bad0-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="6bad0-312">(nullable) itemId :String</span></span>

<span data-ttu-id="6bad0-p118">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-315">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="6bad0-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="6bad0-316">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="6bad0-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="6bad0-317">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью метода `Office.context.mailbox.convertToRestId`, который доступен в наборе обязательных элементов, начиная с версии 1.3.</span><span class="sxs-lookup"><span data-stu-id="6bad0-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="6bad0-318">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="6bad0-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-319">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-319">Type:</span></span>

*   <span data-ttu-id="6bad0-320">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-321">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-321">Requirements</span></span>

|<span data-ttu-id="6bad0-322">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-322">Requirement</span></span>| <span data-ttu-id="6bad0-323">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-324">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-325">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-325">1.0</span></span>|
|[<span data-ttu-id="6bad0-326">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-327">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-328">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-329">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-330">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-330">Example</span></span>

<span data-ttu-id="6bad0-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="6bad0-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="6bad0-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="6bad0-334">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="6bad0-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="6bad0-335">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="6bad0-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-336">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-336">Type:</span></span>

*   [<span data-ttu-id="6bad0-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="6bad0-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="6bad0-338">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-338">Requirements</span></span>

|<span data-ttu-id="6bad0-339">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-339">Requirement</span></span>| <span data-ttu-id="6bad0-340">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-341">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-342">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-342">1.0</span></span>|
|[<span data-ttu-id="6bad0-343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-344">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-346">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-347">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="6bad0-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="6bad0-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="6bad0-349">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="6bad0-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6bad0-350">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="6bad0-350">Read mode</span></span>

<span data-ttu-id="6bad0-351">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="6bad0-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6bad0-352">Режим создания</span><span class="sxs-lookup"><span data-stu-id="6bad0-352">Compose mode</span></span>

<span data-ttu-id="6bad0-353">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="6bad0-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-354">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-354">Type:</span></span>

*   <span data-ttu-id="6bad0-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="6bad0-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-356">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-356">Requirements</span></span>

|<span data-ttu-id="6bad0-357">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-357">Requirement</span></span>| <span data-ttu-id="6bad0-358">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-359">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-360">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-360">1.0</span></span>|
|[<span data-ttu-id="6bad0-361">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-362">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-363">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-364">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-365">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="6bad0-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="6bad0-366">normalizedSubject :String</span></span>

<span data-ttu-id="6bad0-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="6bad0-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="6bad0-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-371">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-371">Type:</span></span>

*   <span data-ttu-id="6bad0-372">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-373">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-373">Requirements</span></span>

|<span data-ttu-id="6bad0-374">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-374">Requirement</span></span>| <span data-ttu-id="6bad0-375">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-376">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-377">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-377">1.0</span></span>|
|[<span data-ttu-id="6bad0-378">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-379">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-380">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-381">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-382">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6bad0-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6bad0-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6bad0-384">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="6bad0-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="6bad0-385">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6bad0-386">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="6bad0-386">Read mode</span></span>

<span data-ttu-id="6bad0-387">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="6bad0-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6bad0-388">Режим создания</span><span class="sxs-lookup"><span data-stu-id="6bad0-388">Compose mode</span></span>

<span data-ttu-id="6bad0-389">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="6bad0-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-390">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-390">Type:</span></span>

*   <span data-ttu-id="6bad0-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6bad0-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-392">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-392">Requirements</span></span>

|<span data-ttu-id="6bad0-393">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-393">Requirement</span></span>| <span data-ttu-id="6bad0-394">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-395">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-396">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-396">1.0</span></span>|
|[<span data-ttu-id="6bad0-397">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-398">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-399">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-400">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-401">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="6bad0-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6bad0-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="6bad0-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-405">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-405">Type:</span></span>

*   [<span data-ttu-id="6bad0-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6bad0-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6bad0-407">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-407">Requirements</span></span>

|<span data-ttu-id="6bad0-408">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-408">Requirement</span></span>| <span data-ttu-id="6bad0-409">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-410">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-411">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-411">1.0</span></span>|
|[<span data-ttu-id="6bad0-412">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-413">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-414">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-415">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-416">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6bad0-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6bad0-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6bad0-418">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="6bad0-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="6bad0-419">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6bad0-420">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="6bad0-420">Read mode</span></span>

<span data-ttu-id="6bad0-421">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="6bad0-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6bad0-422">Режим создания</span><span class="sxs-lookup"><span data-stu-id="6bad0-422">Compose mode</span></span>

<span data-ttu-id="6bad0-423">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="6bad0-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-424">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-424">Type:</span></span>

*   <span data-ttu-id="6bad0-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6bad0-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-426">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-426">Requirements</span></span>

|<span data-ttu-id="6bad0-427">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-427">Requirement</span></span>| <span data-ttu-id="6bad0-428">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-429">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-430">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-430">1.0</span></span>|
|[<span data-ttu-id="6bad0-431">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-432">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-433">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-434">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-435">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="6bad0-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6bad0-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="6bad0-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="6bad0-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-441">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-442">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-442">Type:</span></span>

*   [<span data-ttu-id="6bad0-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6bad0-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6bad0-444">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-444">Requirements</span></span>

|<span data-ttu-id="6bad0-445">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-445">Requirement</span></span>| <span data-ttu-id="6bad0-446">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-447">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-448">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-448">1.0</span></span>|
|[<span data-ttu-id="6bad0-449">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-450">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-451">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-452">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-453">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="6bad0-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="6bad0-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="6bad0-455">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="6bad0-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="6bad0-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="6bad0-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6bad0-458">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="6bad0-458">Read mode</span></span>

<span data-ttu-id="6bad0-459">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6bad0-460">Режим создания</span><span class="sxs-lookup"><span data-stu-id="6bad0-460">Compose mode</span></span>

<span data-ttu-id="6bad0-461">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="6bad0-462">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="6bad0-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-463">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-463">Type:</span></span>

*   <span data-ttu-id="6bad0-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="6bad0-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-465">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-465">Requirements</span></span>

|<span data-ttu-id="6bad0-466">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-466">Requirement</span></span>| <span data-ttu-id="6bad0-467">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-468">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-469">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-469">1.0</span></span>|
|[<span data-ttu-id="6bad0-470">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-471">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-472">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-473">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-474">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-474">Example</span></span>

<span data-ttu-id="6bad0-475">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="6bad0-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="6bad0-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="6bad0-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="6bad0-477">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="6bad0-478">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="6bad0-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6bad0-479">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="6bad0-479">Read mode</span></span>

<span data-ttu-id="6bad0-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="6bad0-482">Режим создания</span><span class="sxs-lookup"><span data-stu-id="6bad0-482">Compose mode</span></span>

<span data-ttu-id="6bad0-483">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="6bad0-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6bad0-484">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-484">Type:</span></span>

*   <span data-ttu-id="6bad0-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="6bad0-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-486">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-486">Requirements</span></span>

|<span data-ttu-id="6bad0-487">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-487">Requirement</span></span>| <span data-ttu-id="6bad0-488">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-489">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-490">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-490">1.0</span></span>|
|[<span data-ttu-id="6bad0-491">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-492">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-493">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-494">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="6bad0-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6bad0-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="6bad0-496">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="6bad0-497">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6bad0-498">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="6bad0-498">Read mode</span></span>

<span data-ttu-id="6bad0-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6bad0-501">Режим создания</span><span class="sxs-lookup"><span data-stu-id="6bad0-501">Compose mode</span></span>

<span data-ttu-id="6bad0-502">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="6bad0-503">Тип:</span><span class="sxs-lookup"><span data-stu-id="6bad0-503">Type:</span></span>

*   <span data-ttu-id="6bad0-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6bad0-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-505">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-505">Requirements</span></span>

|<span data-ttu-id="6bad0-506">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-506">Requirement</span></span>| <span data-ttu-id="6bad0-507">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-508">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-509">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-509">1.0</span></span>|
|[<span data-ttu-id="6bad0-510">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-511">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-512">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-513">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-514">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="6bad0-515">Методы</span><span class="sxs-lookup"><span data-stu-id="6bad0-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="6bad0-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6bad0-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6bad0-517">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="6bad0-518">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="6bad0-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="6bad0-519">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="6bad0-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-520">Параметры</span><span class="sxs-lookup"><span data-stu-id="6bad0-520">Parameters:</span></span>

|<span data-ttu-id="6bad0-521">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-521">Name</span></span>| <span data-ttu-id="6bad0-522">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-522">Type</span></span>| <span data-ttu-id="6bad0-523">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6bad0-523">Attributes</span></span>| <span data-ttu-id="6bad0-524">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="6bad0-525">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-525">String</span></span>||<span data-ttu-id="6bad0-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6bad0-528">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-528">String</span></span>||<span data-ttu-id="6bad0-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6bad0-531">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-531">Object</span></span>| <span data-ttu-id="6bad0-532">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-532">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-533">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="6bad0-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6bad0-534">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-534">Object</span></span>| <span data-ttu-id="6bad0-535">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-535">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-536">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="6bad0-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6bad0-537">функция</span><span class="sxs-lookup"><span data-stu-id="6bad0-537">function</span></span>| <span data-ttu-id="6bad0-538">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-538">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-539">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6bad0-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6bad0-540">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6bad0-541">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="6bad0-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6bad0-542">Ошибки</span><span class="sxs-lookup"><span data-stu-id="6bad0-542">Errors</span></span>

| <span data-ttu-id="6bad0-543">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="6bad0-543">Error code</span></span> | <span data-ttu-id="6bad0-544">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="6bad0-545">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="6bad0-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="6bad0-546">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="6bad0-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6bad0-547">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="6bad0-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6bad0-548">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-548">Requirements</span></span>

|<span data-ttu-id="6bad0-549">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-549">Requirement</span></span>| <span data-ttu-id="6bad0-550">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-551">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-552">1.1</span><span class="sxs-lookup"><span data-stu-id="6bad0-552">1.1</span></span>|
|[<span data-ttu-id="6bad0-553">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="6bad0-555">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-556">Создание</span><span class="sxs-lookup"><span data-stu-id="6bad0-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-557">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-557">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="6bad0-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6bad0-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6bad0-559">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="6bad0-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="6bad0-563">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="6bad0-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="6bad0-564">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="6bad0-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-565">Параметры:</span><span class="sxs-lookup"><span data-stu-id="6bad0-565">Parameters:</span></span>

|<span data-ttu-id="6bad0-566">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-566">Name</span></span>| <span data-ttu-id="6bad0-567">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-567">Type</span></span>| <span data-ttu-id="6bad0-568">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6bad0-568">Attributes</span></span>| <span data-ttu-id="6bad0-569">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="6bad0-570">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-570">String</span></span>||<span data-ttu-id="6bad0-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6bad0-573">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-573">String</span></span>||<span data-ttu-id="6bad0-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6bad0-576">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-576">Object</span></span>| <span data-ttu-id="6bad0-577">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-577">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-578">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="6bad0-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6bad0-579">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-579">Object</span></span>| <span data-ttu-id="6bad0-580">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-580">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-581">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="6bad0-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6bad0-582">функция</span><span class="sxs-lookup"><span data-stu-id="6bad0-582">function</span></span>| <span data-ttu-id="6bad0-583">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-583">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-584">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6bad0-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6bad0-585">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6bad0-586">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="6bad0-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6bad0-587">Ошибки</span><span class="sxs-lookup"><span data-stu-id="6bad0-587">Errors</span></span>

| <span data-ttu-id="6bad0-588">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="6bad0-588">Error code</span></span> | <span data-ttu-id="6bad0-589">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6bad0-590">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="6bad0-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6bad0-591">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-591">Requirements</span></span>

|<span data-ttu-id="6bad0-592">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-592">Requirement</span></span>| <span data-ttu-id="6bad0-593">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-594">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-595">1.1</span><span class="sxs-lookup"><span data-stu-id="6bad0-595">1.1</span></span>|
|[<span data-ttu-id="6bad0-596">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="6bad0-598">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-599">Создание</span><span class="sxs-lookup"><span data-stu-id="6bad0-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-600">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-600">Example</span></span>

<span data-ttu-id="6bad0-601">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="6bad0-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="6bad0-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="6bad0-603">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="6bad0-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-604">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="6bad0-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6bad0-605">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="6bad0-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6bad0-606">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="6bad0-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="6bad0-p137">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-610">Параметры</span><span class="sxs-lookup"><span data-stu-id="6bad0-610">Parameters:</span></span>

|<span data-ttu-id="6bad0-611">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-611">Name</span></span>| <span data-ttu-id="6bad0-612">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-612">Type</span></span>| <span data-ttu-id="6bad0-613">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-613">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="6bad0-614">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-614">String &#124; Object</span></span>| |<span data-ttu-id="6bad0-p138">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6bad0-617">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="6bad0-617">**OR**</span></span><br/><span data-ttu-id="6bad0-p139">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6bad0-620">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-620">String</span></span> | <span data-ttu-id="6bad0-621">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-621">&lt;optional&gt;</span></span> | <span data-ttu-id="6bad0-p140">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6bad0-624">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-624">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6bad0-625">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-625">&lt;optional&gt;</span></span> | <span data-ttu-id="6bad0-626">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="6bad0-626">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6bad0-627">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-627">String</span></span> | | <span data-ttu-id="6bad0-p141">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6bad0-630">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-630">String</span></span> | | <span data-ttu-id="6bad0-631">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-631">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6bad0-632">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-632">String</span></span> | | <span data-ttu-id="6bad0-p142">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6bad0-635">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-635">String</span></span> | | <span data-ttu-id="6bad0-p143">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6bad0-639">function</span><span class="sxs-lookup"><span data-stu-id="6bad0-639">function</span></span> | <span data-ttu-id="6bad0-640">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-640">&lt;optional&gt;</span></span> | <span data-ttu-id="6bad0-641">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6bad0-641">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6bad0-642">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-642">Requirements</span></span>

|<span data-ttu-id="6bad0-643">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-643">Requirement</span></span>| <span data-ttu-id="6bad0-644">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-645">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-646">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-646">1.0</span></span>|
|[<span data-ttu-id="6bad0-647">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-648">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-649">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-650">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-650">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6bad0-651">Примеры</span><span class="sxs-lookup"><span data-stu-id="6bad0-651">Examples</span></span>

<span data-ttu-id="6bad0-652">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-652">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="6bad0-653">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-653">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="6bad0-654">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-654">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6bad0-655">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="6bad0-655">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="6bad0-656">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="6bad0-656">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="6bad0-657">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="6bad0-657">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="6bad0-658">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="6bad0-658">displayReplyForm(formData)</span></span>

<span data-ttu-id="6bad0-659">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="6bad0-659">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-660">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="6bad0-660">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6bad0-661">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="6bad0-661">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6bad0-662">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="6bad0-662">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="6bad0-p144">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-666">Параметры</span><span class="sxs-lookup"><span data-stu-id="6bad0-666">Parameters:</span></span>

|<span data-ttu-id="6bad0-667">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-667">Name</span></span>| <span data-ttu-id="6bad0-668">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-668">Type</span></span>| <span data-ttu-id="6bad0-669">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-669">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="6bad0-670">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-670">String &#124; Object</span></span>| | <span data-ttu-id="6bad0-p145">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6bad0-673">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="6bad0-673">**OR**</span></span><br/><span data-ttu-id="6bad0-p146">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6bad0-676">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-676">String</span></span> | <span data-ttu-id="6bad0-677">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-677">&lt;optional&gt;</span></span> | <span data-ttu-id="6bad0-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6bad0-680">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-680">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6bad0-681">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-681">&lt;optional&gt;</span></span> | <span data-ttu-id="6bad0-682">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="6bad0-682">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6bad0-683">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-683">String</span></span> | | <span data-ttu-id="6bad0-p148">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6bad0-686">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-686">String</span></span> | | <span data-ttu-id="6bad0-687">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-687">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6bad0-688">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-688">String</span></span> | | <span data-ttu-id="6bad0-p149">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6bad0-691">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-691">String</span></span> | | <span data-ttu-id="6bad0-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6bad0-695">function</span><span class="sxs-lookup"><span data-stu-id="6bad0-695">function</span></span> | <span data-ttu-id="6bad0-696">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-696">&lt;optional&gt;</span></span> | <span data-ttu-id="6bad0-697">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6bad0-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6bad0-698">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-698">Requirements</span></span>

|<span data-ttu-id="6bad0-699">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-699">Requirement</span></span>| <span data-ttu-id="6bad0-700">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-700">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-701">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-701">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-702">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-702">1.0</span></span>|
|[<span data-ttu-id="6bad0-703">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-703">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-704">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-704">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-705">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-705">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-706">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-706">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6bad0-707">Примеры</span><span class="sxs-lookup"><span data-stu-id="6bad0-707">Examples</span></span>

<span data-ttu-id="6bad0-708">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-708">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="6bad0-709">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-709">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="6bad0-710">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-710">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6bad0-711">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="6bad0-711">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="6bad0-712">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="6bad0-712">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="6bad0-713">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="6bad0-713">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="6bad0-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="6bad0-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="6bad0-715">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-715">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-716">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="6bad0-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-717">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-717">Requirements</span></span>

|<span data-ttu-id="6bad0-718">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-718">Requirement</span></span>| <span data-ttu-id="6bad0-719">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-720">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-721">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-721">1.0</span></span>|
|[<span data-ttu-id="6bad0-722">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-722">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-723">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-723">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-724">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-724">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-725">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-725">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6bad0-726">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="6bad0-726">Returns:</span></span>

<span data-ttu-id="6bad0-727">Тип: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="6bad0-727">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="6bad0-728">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-728">Example</span></span>

<span data-ttu-id="6bad0-729">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-729">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="6bad0-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="6bad0-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="6bad0-731">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-731">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-732">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="6bad0-732">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-733">Параметры</span><span class="sxs-lookup"><span data-stu-id="6bad0-733">Parameters:</span></span>

|<span data-ttu-id="6bad0-734">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-734">Name</span></span>| <span data-ttu-id="6bad0-735">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-735">Type</span></span>| <span data-ttu-id="6bad0-736">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-736">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="6bad0-737">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="6bad0-737">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="6bad0-738">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="6bad0-738">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6bad0-739">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-739">Requirements</span></span>

|<span data-ttu-id="6bad0-740">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-740">Requirement</span></span>| <span data-ttu-id="6bad0-741">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-742">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-743">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-743">1.0</span></span>|
|[<span data-ttu-id="6bad0-744">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-744">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-745">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="6bad0-745">Restricted</span></span>|
|[<span data-ttu-id="6bad0-746">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-746">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-747">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-747">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6bad0-748">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="6bad0-748">Returns:</span></span>

<span data-ttu-id="6bad0-749">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="6bad0-749">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="6bad0-750">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="6bad0-750">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="6bad0-751">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-751">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="6bad0-752">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="6bad0-752">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="6bad0-753">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="6bad0-753">Value of `entityType`</span></span> | <span data-ttu-id="6bad0-754">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="6bad0-754">Type of objects in returned array</span></span> | <span data-ttu-id="6bad0-755">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-755">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="6bad0-756">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-756">String</span></span> | <span data-ttu-id="6bad0-757">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="6bad0-757">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="6bad0-758">Contact</span><span class="sxs-lookup"><span data-stu-id="6bad0-758">Contact</span></span> | <span data-ttu-id="6bad0-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6bad0-759">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="6bad0-760">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-760">String</span></span> | <span data-ttu-id="6bad0-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6bad0-761">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="6bad0-762">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="6bad0-762">MeetingSuggestion</span></span> | <span data-ttu-id="6bad0-763">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6bad0-763">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="6bad0-764">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="6bad0-764">PhoneNumber</span></span> | <span data-ttu-id="6bad0-765">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="6bad0-765">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="6bad0-766">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="6bad0-766">TaskSuggestion</span></span> | <span data-ttu-id="6bad0-767">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6bad0-767">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="6bad0-768">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-768">String</span></span> | <span data-ttu-id="6bad0-769">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="6bad0-769">**Restricted**</span></span> |

<span data-ttu-id="6bad0-770">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="6bad0-770">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="6bad0-771">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-771">Example</span></span>

<span data-ttu-id="6bad0-772">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-772">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="6bad0-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="6bad0-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="6bad0-774">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="6bad0-774">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-775">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="6bad0-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6bad0-776">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-776">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-777">Параметры</span><span class="sxs-lookup"><span data-stu-id="6bad0-777">Parameters:</span></span>

|<span data-ttu-id="6bad0-778">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-778">Name</span></span>| <span data-ttu-id="6bad0-779">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-779">Type</span></span>| <span data-ttu-id="6bad0-780">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-780">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6bad0-781">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-781">String</span></span>|<span data-ttu-id="6bad0-782">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="6bad0-782">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6bad0-783">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-783">Requirements</span></span>

|<span data-ttu-id="6bad0-784">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-784">Requirement</span></span>| <span data-ttu-id="6bad0-785">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-786">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-787">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-787">1.0</span></span>|
|[<span data-ttu-id="6bad0-788">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-789">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-789">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-790">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-791">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-791">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6bad0-792">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="6bad0-792">Returns:</span></span>

<span data-ttu-id="6bad0-p152">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="6bad0-795">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="6bad0-795">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="6bad0-796">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="6bad0-796">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="6bad0-797">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="6bad0-797">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-798">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="6bad0-798">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6bad0-p153">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="6bad0-802">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-802">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="6bad0-803">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-803">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="6bad0-p154">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bad0-806">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-806">Requirements</span></span>

|<span data-ttu-id="6bad0-807">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-807">Requirement</span></span>| <span data-ttu-id="6bad0-808">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-808">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-809">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-809">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-810">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-810">1.0</span></span>|
|[<span data-ttu-id="6bad0-811">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-811">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-812">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-812">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-813">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-813">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-814">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-814">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6bad0-815">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="6bad0-815">Returns:</span></span>

<span data-ttu-id="6bad0-p155">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="6bad0-818">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="6bad0-818">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6bad0-819">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-819">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6bad0-820">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-820">Example</span></span>

<span data-ttu-id="6bad0-821">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="6bad0-821">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="6bad0-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="6bad0-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="6bad0-823">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="6bad0-823">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6bad0-824">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="6bad0-824">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6bad0-825">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-825">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="6bad0-p156">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-828">Параметры</span><span class="sxs-lookup"><span data-stu-id="6bad0-828">Parameters:</span></span>

|<span data-ttu-id="6bad0-829">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-829">Name</span></span>| <span data-ttu-id="6bad0-830">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-830">Type</span></span>| <span data-ttu-id="6bad0-831">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-831">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6bad0-832">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-832">String</span></span>|<span data-ttu-id="6bad0-833">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="6bad0-833">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6bad0-834">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-834">Requirements</span></span>

|<span data-ttu-id="6bad0-835">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-835">Requirement</span></span>| <span data-ttu-id="6bad0-836">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-836">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-837">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-837">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-838">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-838">1.0</span></span>|
|[<span data-ttu-id="6bad0-839">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-839">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-840">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-840">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-841">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-841">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-842">Чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-842">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6bad0-843">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="6bad0-843">Returns:</span></span>

<span data-ttu-id="6bad0-844">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="6bad0-844">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="6bad0-845">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="6bad0-845">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6bad0-846">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="6bad0-846">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6bad0-847">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-847">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="6bad0-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="6bad0-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="6bad0-849">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-849">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="6bad0-p157">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-852">Параметры</span><span class="sxs-lookup"><span data-stu-id="6bad0-852">Parameters:</span></span>

|<span data-ttu-id="6bad0-853">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-853">Name</span></span>| <span data-ttu-id="6bad0-854">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-854">Type</span></span>| <span data-ttu-id="6bad0-855">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6bad0-855">Attributes</span></span>| <span data-ttu-id="6bad0-856">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-856">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="6bad0-857">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6bad0-857">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="6bad0-p158">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="6bad0-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="6bad0-861">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-861">Object</span></span>| <span data-ttu-id="6bad0-862">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-862">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-863">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="6bad0-863">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6bad0-864">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-864">Object</span></span>| <span data-ttu-id="6bad0-865">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-865">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-866">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="6bad0-866">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6bad0-867">функция</span><span class="sxs-lookup"><span data-stu-id="6bad0-867">function</span></span>||<span data-ttu-id="6bad0-868">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6bad0-868">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6bad0-869">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-869">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="6bad0-870">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-870">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6bad0-871">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-871">Requirements</span></span>

|<span data-ttu-id="6bad0-872">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-872">Requirement</span></span>| <span data-ttu-id="6bad0-873">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-874">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="6bad0-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-875">1.2</span><span class="sxs-lookup"><span data-stu-id="6bad0-875">1.2</span></span>|
|[<span data-ttu-id="6bad0-876">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-876">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-877">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-877">ReadWriteItem</span></span>|
|[<span data-ttu-id="6bad0-878">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-878">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-879">Создание</span><span class="sxs-lookup"><span data-stu-id="6bad0-879">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="6bad0-880">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="6bad0-880">Returns:</span></span>

<span data-ttu-id="6bad0-881">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-881">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="6bad0-882">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="6bad0-882">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6bad0-883">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-883">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6bad0-884">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-884">Example</span></span>

```JavaScript
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="6bad0-885">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6bad0-885">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="6bad0-886">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="6bad0-886">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="6bad0-p160">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-890">Параметры</span><span class="sxs-lookup"><span data-stu-id="6bad0-890">Parameters:</span></span>

|<span data-ttu-id="6bad0-891">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-891">Name</span></span>| <span data-ttu-id="6bad0-892">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-892">Type</span></span>| <span data-ttu-id="6bad0-893">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6bad0-893">Attributes</span></span>| <span data-ttu-id="6bad0-894">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-894">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="6bad0-895">function</span><span class="sxs-lookup"><span data-stu-id="6bad0-895">function</span></span>||<span data-ttu-id="6bad0-896">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6bad0-896">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6bad0-897">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-897">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="6bad0-898">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="6bad0-898">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="6bad0-899">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-899">Object</span></span>| <span data-ttu-id="6bad0-900">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-900">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-901">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="6bad0-901">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="6bad0-902">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="6bad0-902">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6bad0-903">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-903">Requirements</span></span>

|<span data-ttu-id="6bad0-904">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-904">Requirement</span></span>| <span data-ttu-id="6bad0-905">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-906">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-907">1.0</span><span class="sxs-lookup"><span data-stu-id="6bad0-907">1.0</span></span>|
|[<span data-ttu-id="6bad0-908">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-909">ReadItem</span></span>|
|[<span data-ttu-id="6bad0-910">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-911">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6bad0-911">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-912">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-912">Example</span></span>

<span data-ttu-id="6bad0-p163">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="6bad0-916">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6bad0-916">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="6bad0-917">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="6bad0-917">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="6bad0-p164">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-922">Параметры</span><span class="sxs-lookup"><span data-stu-id="6bad0-922">Parameters:</span></span>

|<span data-ttu-id="6bad0-923">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-923">Name</span></span>| <span data-ttu-id="6bad0-924">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-924">Type</span></span>| <span data-ttu-id="6bad0-925">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6bad0-925">Attributes</span></span>| <span data-ttu-id="6bad0-926">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-926">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="6bad0-927">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-927">String</span></span>||<span data-ttu-id="6bad0-928">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-928">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="6bad0-929">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-929">Object</span></span>| <span data-ttu-id="6bad0-930">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-930">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-931">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="6bad0-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6bad0-932">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-932">Object</span></span>| <span data-ttu-id="6bad0-933">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-933">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-934">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="6bad0-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6bad0-935">функция</span><span class="sxs-lookup"><span data-stu-id="6bad0-935">function</span></span>| <span data-ttu-id="6bad0-936">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-936">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-937">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6bad0-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6bad0-938">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="6bad0-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6bad0-939">Ошибки</span><span class="sxs-lookup"><span data-stu-id="6bad0-939">Errors</span></span>

| <span data-ttu-id="6bad0-940">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="6bad0-940">Error code</span></span> | <span data-ttu-id="6bad0-941">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="6bad0-942">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="6bad0-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6bad0-943">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-943">Requirements</span></span>

|<span data-ttu-id="6bad0-944">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-944">Requirement</span></span>| <span data-ttu-id="6bad0-945">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-946">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6bad0-946">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-947">1.1</span><span class="sxs-lookup"><span data-stu-id="6bad0-947">1.1</span></span>|
|[<span data-ttu-id="6bad0-948">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="6bad0-950">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-951">Создание</span><span class="sxs-lookup"><span data-stu-id="6bad0-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-952">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-952">Example</span></span>

<span data-ttu-id="6bad0-953">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="6bad0-953">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="6bad0-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="6bad0-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="6bad0-955">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="6bad0-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="6bad0-p165">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p165">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6bad0-959">Параметры:</span><span class="sxs-lookup"><span data-stu-id="6bad0-959">Parameters:</span></span>

|<span data-ttu-id="6bad0-960">Имя</span><span class="sxs-lookup"><span data-stu-id="6bad0-960">Name</span></span>| <span data-ttu-id="6bad0-961">Тип</span><span class="sxs-lookup"><span data-stu-id="6bad0-961">Type</span></span>| <span data-ttu-id="6bad0-962">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6bad0-962">Attributes</span></span>| <span data-ttu-id="6bad0-963">Описание</span><span class="sxs-lookup"><span data-stu-id="6bad0-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="6bad0-964">String</span><span class="sxs-lookup"><span data-stu-id="6bad0-964">String</span></span>||<span data-ttu-id="6bad0-p166">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p166">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="6bad0-968">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-968">Object</span></span>| <span data-ttu-id="6bad0-969">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-969">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-970">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="6bad0-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6bad0-971">Object</span><span class="sxs-lookup"><span data-stu-id="6bad0-971">Object</span></span>| <span data-ttu-id="6bad0-972">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-972">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-973">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="6bad0-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="6bad0-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6bad0-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="6bad0-975">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="6bad0-975">&lt;optional&gt;</span></span>|<span data-ttu-id="6bad0-p167">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p167">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="6bad0-p168">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="6bad0-p168">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="6bad0-980">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="6bad0-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="6bad0-981">функция</span><span class="sxs-lookup"><span data-stu-id="6bad0-981">function</span></span>||<span data-ttu-id="6bad0-982">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6bad0-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6bad0-983">Требования</span><span class="sxs-lookup"><span data-stu-id="6bad0-983">Requirements</span></span>

|<span data-ttu-id="6bad0-984">Требование</span><span class="sxs-lookup"><span data-stu-id="6bad0-984">Requirement</span></span>| <span data-ttu-id="6bad0-985">Значение</span><span class="sxs-lookup"><span data-stu-id="6bad0-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bad0-986">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="6bad0-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bad0-987">1.2</span><span class="sxs-lookup"><span data-stu-id="6bad0-987">1.2</span></span>|
|[<span data-ttu-id="6bad0-988">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6bad0-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6bad0-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6bad0-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="6bad0-990">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6bad0-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bad0-991">Создание</span><span class="sxs-lookup"><span data-stu-id="6bad0-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6bad0-992">Пример</span><span class="sxs-lookup"><span data-stu-id="6bad0-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
