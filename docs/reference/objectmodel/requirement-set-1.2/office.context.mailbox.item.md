---
title: Office.Context.Mailbox.Item - требование задать 1.2 (en)
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: 2ac3df2a8daae00e64bb66247e66834f9da4243c
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701871"
---
# <a name="item"></a><span data-ttu-id="9866c-102">item</span><span class="sxs-lookup"><span data-stu-id="9866c-102">item</span></span>

### <span data-ttu-id="9866c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="9866c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="9866c-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="9866c-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="9866c-107">Requirements</span></span>

|<span data-ttu-id="9866c-108">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-108">Requirement</span></span>| <span data-ttu-id="9866c-109">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-110">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-111">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-111">1.0</span></span>|
|[<span data-ttu-id="9866c-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9866c-113">Restricted</span></span>|
|[<span data-ttu-id="9866c-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="9866c-116">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-116">Example</span></span>

<span data-ttu-id="9866c-117">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9866c-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="9866c-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="9866c-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="9866c-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9866c-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="9866c-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-122">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="9866c-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="9866c-123">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="9866c-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-124">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-124">Type:</span></span>

*   <span data-ttu-id="9866c-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9866c-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-126">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-126">Requirements</span></span>

|<span data-ttu-id="9866c-127">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-127">Requirement</span></span>| <span data-ttu-id="9866c-128">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-129">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-130">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-130">1.0</span></span>|
|[<span data-ttu-id="9866c-131">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-132">ReadItem</span></span>|
|[<span data-ttu-id="9866c-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-134">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-135">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-135">Example</span></span>

<span data-ttu-id="9866c-136">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9866c-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9866c-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9866c-138">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="9866c-139">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9866c-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-140">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-140">Type:</span></span>

*   [<span data-ttu-id="9866c-141">Recipients</span><span class="sxs-lookup"><span data-stu-id="9866c-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="9866c-142">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-142">Requirements</span></span>

|<span data-ttu-id="9866c-143">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-143">Requirement</span></span>| <span data-ttu-id="9866c-144">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-145">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-146">1.1</span><span class="sxs-lookup"><span data-stu-id="9866c-146">1.1</span></span>|
|[<span data-ttu-id="9866c-147">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-148">ReadItem</span></span>|
|[<span data-ttu-id="9866c-149">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-150">Создание</span><span class="sxs-lookup"><span data-stu-id="9866c-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-151">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="9866c-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="9866c-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="9866c-153">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-154">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-154">Type:</span></span>

*   [<span data-ttu-id="9866c-155">Body</span><span class="sxs-lookup"><span data-stu-id="9866c-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="9866c-156">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-156">Requirements</span></span>

|<span data-ttu-id="9866c-157">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-157">Requirement</span></span>| <span data-ttu-id="9866c-158">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-160">1.1</span><span class="sxs-lookup"><span data-stu-id="9866c-160">1.1</span></span>|
|[<span data-ttu-id="9866c-161">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-162">ReadItem</span></span>|
|[<span data-ttu-id="9866c-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9866c-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9866c-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9866c-166">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="9866c-167">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9866c-168">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9866c-168">Read mode</span></span>

<span data-ttu-id="9866c-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="9866c-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9866c-171">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9866c-171">Compose mode</span></span>

<span data-ttu-id="9866c-172">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-173">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-173">Type:</span></span>

*   <span data-ttu-id="9866c-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9866c-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-175">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-175">Requirements</span></span>

|<span data-ttu-id="9866c-176">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-176">Requirement</span></span>| <span data-ttu-id="9866c-177">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-179">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-179">1.0</span></span>|
|[<span data-ttu-id="9866c-180">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-181">ReadItem</span></span>|
|[<span data-ttu-id="9866c-182">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-183">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-184">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="9866c-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="9866c-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="9866c-186">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="9866c-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="9866c-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="9866c-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="9866c-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="9866c-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-191">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-191">Type:</span></span>

*   <span data-ttu-id="9866c-192">String</span><span class="sxs-lookup"><span data-stu-id="9866c-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-193">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-193">Requirements</span></span>

|<span data-ttu-id="9866c-194">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-194">Requirement</span></span>| <span data-ttu-id="9866c-195">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-196">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-197">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-197">1.0</span></span>|
|[<span data-ttu-id="9866c-198">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-199">ReadItem</span></span>|
|[<span data-ttu-id="9866c-200">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-201">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="9866c-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="9866c-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="9866c-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-205">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-205">Type:</span></span>

*   <span data-ttu-id="9866c-206">Date</span><span class="sxs-lookup"><span data-stu-id="9866c-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-207">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-207">Requirements</span></span>

|<span data-ttu-id="9866c-208">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-208">Requirement</span></span>| <span data-ttu-id="9866c-209">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-210">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-211">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-211">1.0</span></span>|
|[<span data-ttu-id="9866c-212">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-213">ReadItem</span></span>|
|[<span data-ttu-id="9866c-214">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-215">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-216">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="9866c-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="9866c-217">dateTimeModified :Date</span></span>

<span data-ttu-id="9866c-p111">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-220">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9866c-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-221">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-221">Type:</span></span>

*   <span data-ttu-id="9866c-222">Date</span><span class="sxs-lookup"><span data-stu-id="9866c-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-223">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-223">Requirements</span></span>

|<span data-ttu-id="9866c-224">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-224">Requirement</span></span>| <span data-ttu-id="9866c-225">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-226">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-227">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-227">1.0</span></span>|
|[<span data-ttu-id="9866c-228">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-229">ReadItem</span></span>|
|[<span data-ttu-id="9866c-230">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-231">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-232">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="9866c-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="9866c-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="9866c-234">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="9866c-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="9866c-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="9866c-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9866c-237">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9866c-237">Read mode</span></span>

<span data-ttu-id="9866c-238">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="9866c-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9866c-239">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9866c-239">Compose mode</span></span>

<span data-ttu-id="9866c-240">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="9866c-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="9866c-241">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="9866c-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-242">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-242">Type:</span></span>

*   <span data-ttu-id="9866c-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="9866c-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-244">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-244">Requirements</span></span>

|<span data-ttu-id="9866c-245">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-245">Requirement</span></span>| <span data-ttu-id="9866c-246">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-247">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-248">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-248">1.0</span></span>|
|[<span data-ttu-id="9866c-249">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-250">ReadItem</span></span>|
|[<span data-ttu-id="9866c-251">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-252">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-253">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-253">Example</span></span>

<span data-ttu-id="9866c-254">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9866c-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="9866c-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9866c-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="9866c-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="9866c-p114">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="9866c-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-260">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9866c-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-261">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-261">Type:</span></span>

*   [<span data-ttu-id="9866c-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9866c-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9866c-263">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-263">Requirements</span></span>

|<span data-ttu-id="9866c-264">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-264">Requirement</span></span>| <span data-ttu-id="9866c-265">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-267">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-267">1.0</span></span>|
|[<span data-ttu-id="9866c-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-269">ReadItem</span></span>|
|[<span data-ttu-id="9866c-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-271">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="9866c-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="9866c-272">internetMessageId :String</span></span>

<span data-ttu-id="9866c-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-275">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-275">Type:</span></span>

*   <span data-ttu-id="9866c-276">String</span><span class="sxs-lookup"><span data-stu-id="9866c-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-277">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-277">Requirements</span></span>

|<span data-ttu-id="9866c-278">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-278">Requirement</span></span>| <span data-ttu-id="9866c-279">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-280">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-281">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-281">1.0</span></span>|
|[<span data-ttu-id="9866c-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-283">ReadItem</span></span>|
|[<span data-ttu-id="9866c-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-285">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-286">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="9866c-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="9866c-287">itemClass :String</span></span>

<span data-ttu-id="9866c-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="9866c-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="9866c-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="9866c-292">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-292">Type</span></span> | <span data-ttu-id="9866c-293">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-293">Description</span></span> | <span data-ttu-id="9866c-294">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="9866c-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="9866c-295">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="9866c-295">Appointment items</span></span> | <span data-ttu-id="9866c-296">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="9866c-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="9866c-297">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="9866c-297">Message items</span></span> | <span data-ttu-id="9866c-298">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="9866c-299">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="9866c-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-300">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-300">Type:</span></span>

*   <span data-ttu-id="9866c-301">String</span><span class="sxs-lookup"><span data-stu-id="9866c-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-302">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-302">Requirements</span></span>

|<span data-ttu-id="9866c-303">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-303">Requirement</span></span>| <span data-ttu-id="9866c-304">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-305">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-306">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-306">1.0</span></span>|
|[<span data-ttu-id="9866c-307">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-308">ReadItem</span></span>|
|[<span data-ttu-id="9866c-309">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-310">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-311">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="9866c-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="9866c-312">(nullable) itemId :String</span></span>

<span data-ttu-id="9866c-p118">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-315">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="9866c-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9866c-316">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="9866c-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="9866c-317">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью метода `Office.context.mailbox.convertToRestId`, который доступен в наборе обязательных элементов, начиная с версии 1.3.</span><span class="sxs-lookup"><span data-stu-id="9866c-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="9866c-318">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="9866c-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-319">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-319">Type:</span></span>

*   <span data-ttu-id="9866c-320">String</span><span class="sxs-lookup"><span data-stu-id="9866c-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-321">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-321">Requirements</span></span>

|<span data-ttu-id="9866c-322">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-322">Requirement</span></span>| <span data-ttu-id="9866c-323">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-324">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-325">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-325">1.0</span></span>|
|[<span data-ttu-id="9866c-326">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-327">ReadItem</span></span>|
|[<span data-ttu-id="9866c-328">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-329">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-330">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-330">Example</span></span>

<span data-ttu-id="9866c-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="9866c-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="9866c-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="9866c-334">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="9866c-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="9866c-335">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="9866c-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-336">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-336">Type:</span></span>

*   [<span data-ttu-id="9866c-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="9866c-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="9866c-338">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-338">Requirements</span></span>

|<span data-ttu-id="9866c-339">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-339">Requirement</span></span>| <span data-ttu-id="9866c-340">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-341">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-342">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-342">1.0</span></span>|
|[<span data-ttu-id="9866c-343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-344">ReadItem</span></span>|
|[<span data-ttu-id="9866c-345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-346">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-347">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="9866c-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="9866c-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="9866c-349">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="9866c-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9866c-350">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9866c-350">Read mode</span></span>

<span data-ttu-id="9866c-351">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="9866c-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9866c-352">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9866c-352">Compose mode</span></span>

<span data-ttu-id="9866c-353">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="9866c-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-354">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-354">Type:</span></span>

*   <span data-ttu-id="9866c-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="9866c-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-356">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-356">Requirements</span></span>

|<span data-ttu-id="9866c-357">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-357">Requirement</span></span>| <span data-ttu-id="9866c-358">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-359">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-360">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-360">1.0</span></span>|
|[<span data-ttu-id="9866c-361">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-362">ReadItem</span></span>|
|[<span data-ttu-id="9866c-363">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-364">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-365">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="9866c-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="9866c-366">normalizedSubject :String</span></span>

<span data-ttu-id="9866c-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="9866c-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="9866c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-371">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-371">Type:</span></span>

*   <span data-ttu-id="9866c-372">String</span><span class="sxs-lookup"><span data-stu-id="9866c-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-373">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-373">Requirements</span></span>

|<span data-ttu-id="9866c-374">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-374">Requirement</span></span>| <span data-ttu-id="9866c-375">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-376">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-377">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-377">1.0</span></span>|
|[<span data-ttu-id="9866c-378">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-379">ReadItem</span></span>|
|[<span data-ttu-id="9866c-380">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-381">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-382">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9866c-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9866c-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9866c-384">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="9866c-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="9866c-385">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9866c-386">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9866c-386">Read mode</span></span>

<span data-ttu-id="9866c-387">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="9866c-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9866c-388">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9866c-388">Compose mode</span></span>

<span data-ttu-id="9866c-389">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="9866c-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-390">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-390">Type:</span></span>

*   <span data-ttu-id="9866c-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9866c-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-392">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-392">Requirements</span></span>

|<span data-ttu-id="9866c-393">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-393">Requirement</span></span>| <span data-ttu-id="9866c-394">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-395">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-396">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-396">1.0</span></span>|
|[<span data-ttu-id="9866c-397">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-398">ReadItem</span></span>|
|[<span data-ttu-id="9866c-399">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-400">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-401">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="9866c-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9866c-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="9866c-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-405">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-405">Type:</span></span>

*   [<span data-ttu-id="9866c-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9866c-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9866c-407">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-407">Requirements</span></span>

|<span data-ttu-id="9866c-408">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-408">Requirement</span></span>| <span data-ttu-id="9866c-409">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-410">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-411">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-411">1.0</span></span>|
|[<span data-ttu-id="9866c-412">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-413">ReadItem</span></span>|
|[<span data-ttu-id="9866c-414">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-415">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-416">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9866c-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9866c-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9866c-418">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="9866c-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="9866c-419">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9866c-420">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9866c-420">Read mode</span></span>

<span data-ttu-id="9866c-421">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="9866c-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9866c-422">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9866c-422">Compose mode</span></span>

<span data-ttu-id="9866c-423">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="9866c-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-424">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-424">Type:</span></span>

*   <span data-ttu-id="9866c-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9866c-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-426">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-426">Requirements</span></span>

|<span data-ttu-id="9866c-427">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-427">Requirement</span></span>| <span data-ttu-id="9866c-428">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-429">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-430">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-430">1.0</span></span>|
|[<span data-ttu-id="9866c-431">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-432">ReadItem</span></span>|
|[<span data-ttu-id="9866c-433">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-434">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-435">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="9866c-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9866c-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="9866c-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9866c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="9866c-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="9866c-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-441">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9866c-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-442">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-442">Type:</span></span>

*   [<span data-ttu-id="9866c-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9866c-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9866c-444">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-444">Requirements</span></span>

|<span data-ttu-id="9866c-445">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-445">Requirement</span></span>| <span data-ttu-id="9866c-446">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-447">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-448">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-448">1.0</span></span>|
|[<span data-ttu-id="9866c-449">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-450">ReadItem</span></span>|
|[<span data-ttu-id="9866c-451">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-452">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-453">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="9866c-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="9866c-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="9866c-455">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="9866c-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="9866c-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="9866c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9866c-458">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9866c-458">Read mode</span></span>

<span data-ttu-id="9866c-459">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="9866c-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9866c-460">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9866c-460">Compose mode</span></span>

<span data-ttu-id="9866c-461">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="9866c-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="9866c-462">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="9866c-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-463">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-463">Type:</span></span>

*   <span data-ttu-id="9866c-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="9866c-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-465">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-465">Requirements</span></span>

|<span data-ttu-id="9866c-466">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-466">Requirement</span></span>| <span data-ttu-id="9866c-467">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-468">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-469">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-469">1.0</span></span>|
|[<span data-ttu-id="9866c-470">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-471">ReadItem</span></span>|
|[<span data-ttu-id="9866c-472">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-473">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-474">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-474">Example</span></span>

<span data-ttu-id="9866c-475">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9866c-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="9866c-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9866c-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="9866c-477">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="9866c-478">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="9866c-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9866c-479">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9866c-479">Read mode</span></span>

<span data-ttu-id="9866c-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="9866c-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="9866c-482">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9866c-482">Compose mode</span></span>

<span data-ttu-id="9866c-483">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="9866c-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9866c-484">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-484">Type:</span></span>

*   <span data-ttu-id="9866c-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9866c-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-486">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-486">Requirements</span></span>

|<span data-ttu-id="9866c-487">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-487">Requirement</span></span>| <span data-ttu-id="9866c-488">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-489">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-490">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-490">1.0</span></span>|
|[<span data-ttu-id="9866c-491">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-492">ReadItem</span></span>|
|[<span data-ttu-id="9866c-493">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-494">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="9866c-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9866c-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="9866c-496">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="9866c-497">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9866c-498">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9866c-498">Read mode</span></span>

<span data-ttu-id="9866c-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="9866c-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9866c-501">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9866c-501">Compose mode</span></span>

<span data-ttu-id="9866c-502">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="9866c-503">Тип:</span><span class="sxs-lookup"><span data-stu-id="9866c-503">Type:</span></span>

*   <span data-ttu-id="9866c-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9866c-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-505">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-505">Requirements</span></span>

|<span data-ttu-id="9866c-506">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-506">Requirement</span></span>| <span data-ttu-id="9866c-507">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-508">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-509">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-509">1.0</span></span>|
|[<span data-ttu-id="9866c-510">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-511">ReadItem</span></span>|
|[<span data-ttu-id="9866c-512">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-513">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-514">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="9866c-515">Методы</span><span class="sxs-lookup"><span data-stu-id="9866c-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="9866c-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9866c-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9866c-517">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="9866c-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9866c-518">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="9866c-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="9866c-519">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="9866c-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-520">Параметры</span><span class="sxs-lookup"><span data-stu-id="9866c-520">Parameters:</span></span>

|<span data-ttu-id="9866c-521">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-521">Name</span></span>| <span data-ttu-id="9866c-522">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-522">Type</span></span>| <span data-ttu-id="9866c-523">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9866c-523">Attributes</span></span>| <span data-ttu-id="9866c-524">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="9866c-525">String</span><span class="sxs-lookup"><span data-stu-id="9866c-525">String</span></span>||<span data-ttu-id="9866c-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="9866c-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9866c-528">String</span><span class="sxs-lookup"><span data-stu-id="9866c-528">String</span></span>||<span data-ttu-id="9866c-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9866c-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9866c-531">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-531">Object</span></span>| <span data-ttu-id="9866c-532">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-532">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-533">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9866c-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9866c-534">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-534">Object</span></span>| <span data-ttu-id="9866c-535">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-535">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-536">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9866c-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9866c-537">функция</span><span class="sxs-lookup"><span data-stu-id="9866c-537">function</span></span>| <span data-ttu-id="9866c-538">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-538">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-539">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9866c-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9866c-540">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9866c-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9866c-541">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="9866c-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9866c-542">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9866c-542">Errors</span></span>

| <span data-ttu-id="9866c-543">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9866c-543">Error code</span></span> | <span data-ttu-id="9866c-544">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="9866c-545">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="9866c-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="9866c-546">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="9866c-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9866c-547">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="9866c-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9866c-548">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-548">Requirements</span></span>

|<span data-ttu-id="9866c-549">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-549">Requirement</span></span>| <span data-ttu-id="9866c-550">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-551">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-552">1.1</span><span class="sxs-lookup"><span data-stu-id="9866c-552">1.1</span></span>|
|[<span data-ttu-id="9866c-553">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9866c-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="9866c-555">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-556">Создание</span><span class="sxs-lookup"><span data-stu-id="9866c-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-557">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="9866c-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9866c-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9866c-559">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="9866c-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="9866c-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9866c-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="9866c-563">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="9866c-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="9866c-564">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="9866c-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-565">Параметры:</span><span class="sxs-lookup"><span data-stu-id="9866c-565">Parameters:</span></span>

|<span data-ttu-id="9866c-566">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-566">Name</span></span>| <span data-ttu-id="9866c-567">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-567">Type</span></span>| <span data-ttu-id="9866c-568">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9866c-568">Attributes</span></span>| <span data-ttu-id="9866c-569">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="9866c-570">String</span><span class="sxs-lookup"><span data-stu-id="9866c-570">String</span></span>||<span data-ttu-id="9866c-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="9866c-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9866c-573">String</span><span class="sxs-lookup"><span data-stu-id="9866c-573">String</span></span>||<span data-ttu-id="9866c-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9866c-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9866c-576">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-576">Object</span></span>| <span data-ttu-id="9866c-577">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-577">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-578">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9866c-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9866c-579">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-579">Object</span></span>| <span data-ttu-id="9866c-580">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-580">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-581">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9866c-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9866c-582">функция</span><span class="sxs-lookup"><span data-stu-id="9866c-582">function</span></span>| <span data-ttu-id="9866c-583">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-583">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-584">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9866c-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9866c-585">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9866c-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9866c-586">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="9866c-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9866c-587">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9866c-587">Errors</span></span>

| <span data-ttu-id="9866c-588">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9866c-588">Error code</span></span> | <span data-ttu-id="9866c-589">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9866c-590">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="9866c-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9866c-591">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-591">Requirements</span></span>

|<span data-ttu-id="9866c-592">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-592">Requirement</span></span>| <span data-ttu-id="9866c-593">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-594">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-595">1.1</span><span class="sxs-lookup"><span data-stu-id="9866c-595">1.1</span></span>|
|[<span data-ttu-id="9866c-596">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9866c-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="9866c-598">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-599">Создание</span><span class="sxs-lookup"><span data-stu-id="9866c-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-600">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-600">Example</span></span>

<span data-ttu-id="9866c-601">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="9866c-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="9866c-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="9866c-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="9866c-603">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="9866c-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-604">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9866c-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9866c-605">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="9866c-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9866c-606">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="9866c-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="9866c-p137">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="9866c-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-610">Параметры:</span><span class="sxs-lookup"><span data-stu-id="9866c-610">Parameters:</span></span>

|<span data-ttu-id="9866c-611">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-611">Name</span></span>| <span data-ttu-id="9866c-612">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-612">Type</span></span>| <span data-ttu-id="9866c-613">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-613">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9866c-614">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9866c-614">String &#124; Object</span></span>| |<span data-ttu-id="9866c-p138">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9866c-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9866c-617">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="9866c-617">**OR**</span></span><br/><span data-ttu-id="9866c-p139">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="9866c-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9866c-620">String</span><span class="sxs-lookup"><span data-stu-id="9866c-620">String</span></span> | <span data-ttu-id="9866c-621">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-621">&lt;optional&gt;</span></span> | <span data-ttu-id="9866c-p140">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9866c-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="9866c-624">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-624">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9866c-625">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-625">&lt;optional&gt;</span></span> | <span data-ttu-id="9866c-626">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="9866c-626">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="9866c-627">String</span><span class="sxs-lookup"><span data-stu-id="9866c-627">String</span></span> | | <span data-ttu-id="9866c-p141">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="9866c-630">Строка</span><span class="sxs-lookup"><span data-stu-id="9866c-630">String</span></span> | | <span data-ttu-id="9866c-631">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9866c-631">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="9866c-632">Строка</span><span class="sxs-lookup"><span data-stu-id="9866c-632">String</span></span> | | <span data-ttu-id="9866c-p142">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="9866c-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="9866c-635">String</span><span class="sxs-lookup"><span data-stu-id="9866c-635">String</span></span> | | <span data-ttu-id="9866c-p143">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="9866c-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="9866c-639">функция</span><span class="sxs-lookup"><span data-stu-id="9866c-639">function</span></span> | <span data-ttu-id="9866c-640">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-640">&lt;optional&gt;</span></span> | <span data-ttu-id="9866c-641">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9866c-641">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9866c-642">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-642">Requirements</span></span>

|<span data-ttu-id="9866c-643">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-643">Requirement</span></span>| <span data-ttu-id="9866c-644">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-645">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-646">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-646">1.0</span></span>|
|[<span data-ttu-id="9866c-647">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-648">ReadItem</span></span>|
|[<span data-ttu-id="9866c-649">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-650">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-650">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9866c-651">Примеры</span><span class="sxs-lookup"><span data-stu-id="9866c-651">Examples</span></span>

<span data-ttu-id="9866c-652">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="9866c-652">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="9866c-653">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-653">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="9866c-654">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-654">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9866c-655">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="9866c-655">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9866c-656">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="9866c-656">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9866c-657">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="9866c-657">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="9866c-658">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="9866c-658">displayReplyForm(formData)</span></span>

<span data-ttu-id="9866c-659">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="9866c-659">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-660">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9866c-660">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9866c-661">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="9866c-661">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9866c-662">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="9866c-662">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="9866c-p144">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="9866c-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-666">Параметры:</span><span class="sxs-lookup"><span data-stu-id="9866c-666">Parameters:</span></span>

|<span data-ttu-id="9866c-667">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-667">Name</span></span>| <span data-ttu-id="9866c-668">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-668">Type</span></span>| <span data-ttu-id="9866c-669">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-669">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9866c-670">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9866c-670">String &#124; Object</span></span>| | <span data-ttu-id="9866c-p145">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9866c-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9866c-673">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="9866c-673">**OR**</span></span><br/><span data-ttu-id="9866c-p146">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="9866c-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9866c-676">String</span><span class="sxs-lookup"><span data-stu-id="9866c-676">String</span></span> | <span data-ttu-id="9866c-677">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-677">&lt;optional&gt;</span></span> | <span data-ttu-id="9866c-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9866c-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="9866c-680">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-680">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9866c-681">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-681">&lt;optional&gt;</span></span> | <span data-ttu-id="9866c-682">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="9866c-682">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="9866c-683">String</span><span class="sxs-lookup"><span data-stu-id="9866c-683">String</span></span> | | <span data-ttu-id="9866c-p148">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="9866c-686">Строка</span><span class="sxs-lookup"><span data-stu-id="9866c-686">String</span></span> | | <span data-ttu-id="9866c-687">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9866c-687">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="9866c-688">Строка</span><span class="sxs-lookup"><span data-stu-id="9866c-688">String</span></span> | | <span data-ttu-id="9866c-p149">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="9866c-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="9866c-691">String</span><span class="sxs-lookup"><span data-stu-id="9866c-691">String</span></span> | | <span data-ttu-id="9866c-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="9866c-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="9866c-695">function</span><span class="sxs-lookup"><span data-stu-id="9866c-695">function</span></span> | <span data-ttu-id="9866c-696">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-696">&lt;optional&gt;</span></span> | <span data-ttu-id="9866c-697">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9866c-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9866c-698">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-698">Requirements</span></span>

|<span data-ttu-id="9866c-699">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-699">Requirement</span></span>| <span data-ttu-id="9866c-700">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-700">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-701">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-701">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-702">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-702">1.0</span></span>|
|[<span data-ttu-id="9866c-703">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-703">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-704">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-704">ReadItem</span></span>|
|[<span data-ttu-id="9866c-705">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-705">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-706">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-706">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9866c-707">Примеры</span><span class="sxs-lookup"><span data-stu-id="9866c-707">Examples</span></span>

<span data-ttu-id="9866c-708">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="9866c-708">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="9866c-709">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-709">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="9866c-710">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-710">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9866c-711">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="9866c-711">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9866c-712">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="9866c-712">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9866c-713">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="9866c-713">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="9866c-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="9866c-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="9866c-715">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-715">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-716">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9866c-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-717">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-717">Requirements</span></span>

|<span data-ttu-id="9866c-718">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-718">Requirement</span></span>| <span data-ttu-id="9866c-719">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-720">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-721">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-721">1.0</span></span>|
|[<span data-ttu-id="9866c-722">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-722">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-723">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-723">ReadItem</span></span>|
|[<span data-ttu-id="9866c-724">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-724">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-725">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-725">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9866c-726">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9866c-726">Returns:</span></span>

<span data-ttu-id="9866c-727">Тип: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="9866c-727">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="9866c-728">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-728">Example</span></span>

<span data-ttu-id="9866c-729">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-729">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="9866c-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9866c-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9866c-731">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-731">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-732">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9866c-732">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-733">Параметры</span><span class="sxs-lookup"><span data-stu-id="9866c-733">Parameters:</span></span>

|<span data-ttu-id="9866c-734">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-734">Name</span></span>| <span data-ttu-id="9866c-735">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-735">Type</span></span>| <span data-ttu-id="9866c-736">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-736">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="9866c-737">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="9866c-737">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="9866c-738">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="9866c-738">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9866c-739">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-739">Requirements</span></span>

|<span data-ttu-id="9866c-740">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-740">Requirement</span></span>| <span data-ttu-id="9866c-741">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-742">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-743">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-743">1.0</span></span>|
|[<span data-ttu-id="9866c-744">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-744">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-745">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9866c-745">Restricted</span></span>|
|[<span data-ttu-id="9866c-746">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-746">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-747">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-747">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9866c-748">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9866c-748">Returns:</span></span>

<span data-ttu-id="9866c-749">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="9866c-749">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="9866c-750">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="9866c-750">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="9866c-751">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="9866c-751">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="9866c-752">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="9866c-752">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="9866c-753">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="9866c-753">Value of `entityType`</span></span> | <span data-ttu-id="9866c-754">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="9866c-754">Type of objects in returned array</span></span> | <span data-ttu-id="9866c-755">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-755">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="9866c-756">String</span><span class="sxs-lookup"><span data-stu-id="9866c-756">String</span></span> | <span data-ttu-id="9866c-757">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9866c-757">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="9866c-758">Contact</span><span class="sxs-lookup"><span data-stu-id="9866c-758">Contact</span></span> | <span data-ttu-id="9866c-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9866c-759">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="9866c-760">String</span><span class="sxs-lookup"><span data-stu-id="9866c-760">String</span></span> | <span data-ttu-id="9866c-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9866c-761">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="9866c-762">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9866c-762">MeetingSuggestion</span></span> | <span data-ttu-id="9866c-763">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9866c-763">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="9866c-764">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9866c-764">PhoneNumber</span></span> | <span data-ttu-id="9866c-765">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9866c-765">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="9866c-766">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9866c-766">TaskSuggestion</span></span> | <span data-ttu-id="9866c-767">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9866c-767">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="9866c-768">String</span><span class="sxs-lookup"><span data-stu-id="9866c-768">String</span></span> | <span data-ttu-id="9866c-769">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9866c-769">**Restricted**</span></span> |

<span data-ttu-id="9866c-770">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9866c-770">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="9866c-771">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-771">Example</span></span>

<span data-ttu-id="9866c-772">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-772">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="9866c-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9866c-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9866c-774">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9866c-774">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-775">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9866c-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9866c-776">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="9866c-776">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-777">Параметры:</span><span class="sxs-lookup"><span data-stu-id="9866c-777">Parameters:</span></span>

|<span data-ttu-id="9866c-778">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-778">Name</span></span>| <span data-ttu-id="9866c-779">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-779">Type</span></span>| <span data-ttu-id="9866c-780">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-780">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9866c-781">String</span><span class="sxs-lookup"><span data-stu-id="9866c-781">String</span></span>|<span data-ttu-id="9866c-782">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="9866c-782">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9866c-783">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-783">Requirements</span></span>

|<span data-ttu-id="9866c-784">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-784">Requirement</span></span>| <span data-ttu-id="9866c-785">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-786">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-787">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-787">1.0</span></span>|
|[<span data-ttu-id="9866c-788">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-789">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-789">ReadItem</span></span>|
|[<span data-ttu-id="9866c-790">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-791">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-791">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9866c-792">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9866c-792">Returns:</span></span>

<span data-ttu-id="9866c-p152">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="9866c-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="9866c-795">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9866c-795">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="9866c-796">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9866c-796">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="9866c-797">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9866c-797">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-798">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9866c-798">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9866c-p153">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="9866c-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9866c-802">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="9866c-802">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9866c-803">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="9866c-803">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="9866c-p154">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="9866c-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9866c-806">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-806">Requirements</span></span>

|<span data-ttu-id="9866c-807">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-807">Requirement</span></span>| <span data-ttu-id="9866c-808">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-808">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-809">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-809">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-810">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-810">1.0</span></span>|
|[<span data-ttu-id="9866c-811">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-811">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-812">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-812">ReadItem</span></span>|
|[<span data-ttu-id="9866c-813">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-813">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-814">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-814">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9866c-815">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9866c-815">Returns:</span></span>

<span data-ttu-id="9866c-p155">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="9866c-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="9866c-818">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="9866c-818">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9866c-819">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-819">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9866c-820">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-820">Example</span></span>

<span data-ttu-id="9866c-821">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="9866c-821">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="9866c-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="9866c-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="9866c-823">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9866c-823">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9866c-824">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9866c-824">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9866c-825">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="9866c-825">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="9866c-p156">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="9866c-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-828">Параметры:</span><span class="sxs-lookup"><span data-stu-id="9866c-828">Parameters:</span></span>

|<span data-ttu-id="9866c-829">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-829">Name</span></span>| <span data-ttu-id="9866c-830">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-830">Type</span></span>| <span data-ttu-id="9866c-831">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-831">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9866c-832">String</span><span class="sxs-lookup"><span data-stu-id="9866c-832">String</span></span>|<span data-ttu-id="9866c-833">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="9866c-833">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9866c-834">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-834">Requirements</span></span>

|<span data-ttu-id="9866c-835">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-835">Requirement</span></span>| <span data-ttu-id="9866c-836">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-836">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-837">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-837">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-838">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-838">1.0</span></span>|
|[<span data-ttu-id="9866c-839">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-839">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-840">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-840">ReadItem</span></span>|
|[<span data-ttu-id="9866c-841">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-841">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-842">Чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-842">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9866c-843">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9866c-843">Returns:</span></span>

<span data-ttu-id="9866c-844">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9866c-844">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="9866c-845">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="9866c-845">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9866c-846">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="9866c-846">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9866c-847">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-847">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="9866c-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="9866c-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="9866c-849">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-849">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="9866c-p157">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="9866c-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-852">Параметры:</span><span class="sxs-lookup"><span data-stu-id="9866c-852">Parameters:</span></span>

|<span data-ttu-id="9866c-853">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-853">Name</span></span>| <span data-ttu-id="9866c-854">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-854">Type</span></span>| <span data-ttu-id="9866c-855">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9866c-855">Attributes</span></span>| <span data-ttu-id="9866c-856">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-856">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="9866c-857">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9866c-857">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="9866c-p158">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="9866c-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="9866c-861">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-861">Object</span></span>| <span data-ttu-id="9866c-862">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-862">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-863">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9866c-863">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9866c-864">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-864">Object</span></span>| <span data-ttu-id="9866c-865">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-865">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-866">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9866c-866">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9866c-867">функция</span><span class="sxs-lookup"><span data-stu-id="9866c-867">function</span></span>||<span data-ttu-id="9866c-868">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9866c-868">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9866c-869">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="9866c-869">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="9866c-870">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="9866c-870">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9866c-871">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-871">Requirements</span></span>

|<span data-ttu-id="9866c-872">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-872">Requirement</span></span>| <span data-ttu-id="9866c-873">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-874">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-875">1.2</span><span class="sxs-lookup"><span data-stu-id="9866c-875">1.2</span></span>|
|[<span data-ttu-id="9866c-876">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-876">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-877">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9866c-877">ReadWriteItem</span></span>|
|[<span data-ttu-id="9866c-878">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-878">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-879">Создание</span><span class="sxs-lookup"><span data-stu-id="9866c-879">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="9866c-880">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9866c-880">Returns:</span></span>

<span data-ttu-id="9866c-881">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="9866c-881">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="9866c-882">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="9866c-882">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9866c-883">String</span><span class="sxs-lookup"><span data-stu-id="9866c-883">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9866c-884">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-884">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="9866c-885">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9866c-885">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="9866c-886">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="9866c-886">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="9866c-p160">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="9866c-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-890">Параметры</span><span class="sxs-lookup"><span data-stu-id="9866c-890">Parameters:</span></span>

|<span data-ttu-id="9866c-891">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-891">Name</span></span>| <span data-ttu-id="9866c-892">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-892">Type</span></span>| <span data-ttu-id="9866c-893">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9866c-893">Attributes</span></span>| <span data-ttu-id="9866c-894">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-894">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9866c-895">функция</span><span class="sxs-lookup"><span data-stu-id="9866c-895">function</span></span>||<span data-ttu-id="9866c-896">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9866c-896">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9866c-897">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9866c-897">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9866c-898">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="9866c-898">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="9866c-899">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-899">Object</span></span>| <span data-ttu-id="9866c-900">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-900">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-901">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9866c-901">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="9866c-902">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9866c-902">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9866c-903">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-903">Requirements</span></span>

|<span data-ttu-id="9866c-904">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-904">Requirement</span></span>| <span data-ttu-id="9866c-905">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-906">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-907">1.0</span><span class="sxs-lookup"><span data-stu-id="9866c-907">1.0</span></span>|
|[<span data-ttu-id="9866c-908">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9866c-909">ReadItem</span></span>|
|[<span data-ttu-id="9866c-910">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-911">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9866c-911">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-912">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-912">Example</span></span>

<span data-ttu-id="9866c-p163">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="9866c-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="9866c-916">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9866c-916">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="9866c-917">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="9866c-917">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="9866c-p164">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="9866c-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-922">Параметры:</span><span class="sxs-lookup"><span data-stu-id="9866c-922">Parameters:</span></span>

|<span data-ttu-id="9866c-923">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-923">Name</span></span>| <span data-ttu-id="9866c-924">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-924">Type</span></span>| <span data-ttu-id="9866c-925">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9866c-925">Attributes</span></span>| <span data-ttu-id="9866c-926">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-926">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="9866c-927">String</span><span class="sxs-lookup"><span data-stu-id="9866c-927">String</span></span>||<span data-ttu-id="9866c-928">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="9866c-928">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="9866c-929">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-929">Object</span></span>| <span data-ttu-id="9866c-930">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-930">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-931">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9866c-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9866c-932">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-932">Object</span></span>| <span data-ttu-id="9866c-933">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-933">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-934">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9866c-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9866c-935">функция</span><span class="sxs-lookup"><span data-stu-id="9866c-935">function</span></span>| <span data-ttu-id="9866c-936">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-936">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-937">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9866c-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9866c-938">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="9866c-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9866c-939">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9866c-939">Errors</span></span>

| <span data-ttu-id="9866c-940">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9866c-940">Error code</span></span> | <span data-ttu-id="9866c-941">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="9866c-942">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="9866c-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9866c-943">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-943">Requirements</span></span>

|<span data-ttu-id="9866c-944">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-944">Requirement</span></span>| <span data-ttu-id="9866c-945">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-946">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9866c-946">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-947">1.1</span><span class="sxs-lookup"><span data-stu-id="9866c-947">1.1</span></span>|
|[<span data-ttu-id="9866c-948">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9866c-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="9866c-950">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-951">Создание</span><span class="sxs-lookup"><span data-stu-id="9866c-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-952">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-952">Example</span></span>

<span data-ttu-id="9866c-953">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="9866c-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="9866c-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="9866c-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="9866c-955">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="9866c-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="9866c-p165">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="9866c-p165">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9866c-959">Параметры:</span><span class="sxs-lookup"><span data-stu-id="9866c-959">Parameters:</span></span>

|<span data-ttu-id="9866c-960">Имя</span><span class="sxs-lookup"><span data-stu-id="9866c-960">Name</span></span>| <span data-ttu-id="9866c-961">Тип</span><span class="sxs-lookup"><span data-stu-id="9866c-961">Type</span></span>| <span data-ttu-id="9866c-962">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9866c-962">Attributes</span></span>| <span data-ttu-id="9866c-963">Описание</span><span class="sxs-lookup"><span data-stu-id="9866c-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="9866c-964">String</span><span class="sxs-lookup"><span data-stu-id="9866c-964">String</span></span>||<span data-ttu-id="9866c-p166">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="9866c-p166">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="9866c-968">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-968">Object</span></span>| <span data-ttu-id="9866c-969">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-969">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-970">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9866c-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9866c-971">Object</span><span class="sxs-lookup"><span data-stu-id="9866c-971">Object</span></span>| <span data-ttu-id="9866c-972">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-972">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-973">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="9866c-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="9866c-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9866c-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="9866c-975">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9866c-975">&lt;optional&gt;</span></span>|<span data-ttu-id="9866c-p167">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="9866c-p167">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="9866c-p168">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="9866c-p168">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="9866c-980">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="9866c-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="9866c-981">функция</span><span class="sxs-lookup"><span data-stu-id="9866c-981">function</span></span>||<span data-ttu-id="9866c-982">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9866c-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9866c-983">Требования</span><span class="sxs-lookup"><span data-stu-id="9866c-983">Requirements</span></span>

|<span data-ttu-id="9866c-984">Требование</span><span class="sxs-lookup"><span data-stu-id="9866c-984">Requirement</span></span>| <span data-ttu-id="9866c-985">Значение</span><span class="sxs-lookup"><span data-stu-id="9866c-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="9866c-986">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9866c-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9866c-987">1.2</span><span class="sxs-lookup"><span data-stu-id="9866c-987">1.2</span></span>|
|[<span data-ttu-id="9866c-988">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9866c-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9866c-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9866c-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="9866c-990">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9866c-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9866c-991">Создание</span><span class="sxs-lookup"><span data-stu-id="9866c-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9866c-992">Пример</span><span class="sxs-lookup"><span data-stu-id="9866c-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
