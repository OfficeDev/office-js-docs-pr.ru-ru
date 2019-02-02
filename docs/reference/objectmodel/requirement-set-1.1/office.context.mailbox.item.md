---
title: Office.Context.Mailbox.Item - требование задать 1.1
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: ce8c10987c08609eba90a3a957b372114e62cd81
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701878"
---
# <a name="item"></a><span data-ttu-id="b877b-102">item</span><span class="sxs-lookup"><span data-stu-id="b877b-102">item</span></span>

### <span data-ttu-id="b877b-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="b877b-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="b877b-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="b877b-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="b877b-107">Requirements</span></span>

|<span data-ttu-id="b877b-108">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-108">Requirement</span></span>| <span data-ttu-id="b877b-109">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-110">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-111">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-111">1.0</span></span>|
|[<span data-ttu-id="b877b-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="b877b-113">Restricted</span></span>|
|[<span data-ttu-id="b877b-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="b877b-116">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-116">Example</span></span>

<span data-ttu-id="b877b-117">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="b877b-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="b877b-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="b877b-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="b877b-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b877b-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="b877b-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-122">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="b877b-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b877b-123">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="b877b-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-124">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-124">Type:</span></span>

*   <span data-ttu-id="b877b-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b877b-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-126">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-126">Requirements</span></span>

|<span data-ttu-id="b877b-127">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-127">Requirement</span></span>| <span data-ttu-id="b877b-128">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-129">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-130">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-130">1.0</span></span>|
|[<span data-ttu-id="b877b-131">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-132">ReadItem</span></span>|
|[<span data-ttu-id="b877b-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-134">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-135">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-135">Example</span></span>

<span data-ttu-id="b877b-136">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="b877b-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b877b-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="b877b-138">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b877b-139">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b877b-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-140">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-140">Type:</span></span>

*   [<span data-ttu-id="b877b-141">Recipients</span><span class="sxs-lookup"><span data-stu-id="b877b-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b877b-142">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-142">Requirements</span></span>

|<span data-ttu-id="b877b-143">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-143">Requirement</span></span>| <span data-ttu-id="b877b-144">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-145">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-146">1.1</span><span class="sxs-lookup"><span data-stu-id="b877b-146">1.1</span></span>|
|[<span data-ttu-id="b877b-147">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-148">ReadItem</span></span>|
|[<span data-ttu-id="b877b-149">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-150">Создание</span><span class="sxs-lookup"><span data-stu-id="b877b-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-151">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="b877b-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="b877b-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="b877b-153">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-154">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-154">Type:</span></span>

*   [<span data-ttu-id="b877b-155">Body</span><span class="sxs-lookup"><span data-stu-id="b877b-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="b877b-156">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-156">Requirements</span></span>

|<span data-ttu-id="b877b-157">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-157">Requirement</span></span>| <span data-ttu-id="b877b-158">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-159">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-160">1.1</span><span class="sxs-lookup"><span data-stu-id="b877b-160">1.1</span></span>|
|[<span data-ttu-id="b877b-161">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-162">ReadItem</span></span>|
|[<span data-ttu-id="b877b-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="b877b-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b877b-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="b877b-166">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b877b-167">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b877b-168">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b877b-168">Read mode</span></span>

<span data-ttu-id="b877b-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="b877b-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b877b-171">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b877b-171">Compose mode</span></span>

<span data-ttu-id="b877b-172">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-173">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-173">Type:</span></span>

*   <span data-ttu-id="b877b-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b877b-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-175">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-175">Requirements</span></span>

|<span data-ttu-id="b877b-176">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-176">Requirement</span></span>| <span data-ttu-id="b877b-177">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-179">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-179">1.0</span></span>|
|[<span data-ttu-id="b877b-180">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-181">ReadItem</span></span>|
|[<span data-ttu-id="b877b-182">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-183">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-184">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b877b-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b877b-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="b877b-186">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="b877b-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b877b-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="b877b-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b877b-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="b877b-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-191">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-191">Type:</span></span>

*   <span data-ttu-id="b877b-192">String</span><span class="sxs-lookup"><span data-stu-id="b877b-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-193">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-193">Requirements</span></span>

|<span data-ttu-id="b877b-194">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-194">Requirement</span></span>| <span data-ttu-id="b877b-195">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-196">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-197">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-197">1.0</span></span>|
|[<span data-ttu-id="b877b-198">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-199">ReadItem</span></span>|
|[<span data-ttu-id="b877b-200">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-201">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="b877b-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b877b-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="b877b-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-205">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-205">Type:</span></span>

*   <span data-ttu-id="b877b-206">Date</span><span class="sxs-lookup"><span data-stu-id="b877b-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-207">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-207">Requirements</span></span>

|<span data-ttu-id="b877b-208">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-208">Requirement</span></span>| <span data-ttu-id="b877b-209">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-210">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-211">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-211">1.0</span></span>|
|[<span data-ttu-id="b877b-212">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-213">ReadItem</span></span>|
|[<span data-ttu-id="b877b-214">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-215">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-216">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b877b-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b877b-217">dateTimeModified :Date</span></span>

<span data-ttu-id="b877b-p111">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-220">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b877b-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-221">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-221">Type:</span></span>

*   <span data-ttu-id="b877b-222">Date</span><span class="sxs-lookup"><span data-stu-id="b877b-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-223">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-223">Requirements</span></span>

|<span data-ttu-id="b877b-224">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-224">Requirement</span></span>| <span data-ttu-id="b877b-225">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-226">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-227">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-227">1.0</span></span>|
|[<span data-ttu-id="b877b-228">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-229">ReadItem</span></span>|
|[<span data-ttu-id="b877b-230">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-231">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-232">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="b877b-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="b877b-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="b877b-234">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="b877b-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b877b-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="b877b-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b877b-237">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b877b-237">Read mode</span></span>

<span data-ttu-id="b877b-238">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="b877b-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b877b-239">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b877b-239">Compose mode</span></span>

<span data-ttu-id="b877b-240">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="b877b-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b877b-241">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="b877b-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-242">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-242">Type:</span></span>

*   <span data-ttu-id="b877b-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="b877b-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-244">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-244">Requirements</span></span>

|<span data-ttu-id="b877b-245">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-245">Requirement</span></span>| <span data-ttu-id="b877b-246">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-247">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-248">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-248">1.0</span></span>|
|[<span data-ttu-id="b877b-249">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-250">ReadItem</span></span>|
|[<span data-ttu-id="b877b-251">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-252">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-253">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-253">Example</span></span>

<span data-ttu-id="b877b-254">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b877b-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="b877b-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b877b-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="b877b-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="b877b-p114">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="b877b-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-260">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b877b-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-261">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-261">Type:</span></span>

*   [<span data-ttu-id="b877b-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b877b-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b877b-263">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-263">Requirements</span></span>

|<span data-ttu-id="b877b-264">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-264">Requirement</span></span>| <span data-ttu-id="b877b-265">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-267">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-267">1.0</span></span>|
|[<span data-ttu-id="b877b-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-269">ReadItem</span></span>|
|[<span data-ttu-id="b877b-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-271">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="b877b-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b877b-272">internetMessageId :String</span></span>

<span data-ttu-id="b877b-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-275">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-275">Type:</span></span>

*   <span data-ttu-id="b877b-276">String</span><span class="sxs-lookup"><span data-stu-id="b877b-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-277">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-277">Requirements</span></span>

|<span data-ttu-id="b877b-278">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-278">Requirement</span></span>| <span data-ttu-id="b877b-279">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-280">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-281">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-281">1.0</span></span>|
|[<span data-ttu-id="b877b-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-283">ReadItem</span></span>|
|[<span data-ttu-id="b877b-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-285">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-286">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b877b-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b877b-287">itemClass :String</span></span>

<span data-ttu-id="b877b-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b877b-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="b877b-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="b877b-292">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-292">Type</span></span> | <span data-ttu-id="b877b-293">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-293">Description</span></span> | <span data-ttu-id="b877b-294">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="b877b-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="b877b-295">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="b877b-295">Appointment items</span></span> | <span data-ttu-id="b877b-296">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="b877b-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="b877b-297">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="b877b-297">Message items</span></span> | <span data-ttu-id="b877b-298">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="b877b-299">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="b877b-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-300">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-300">Type:</span></span>

*   <span data-ttu-id="b877b-301">String</span><span class="sxs-lookup"><span data-stu-id="b877b-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-302">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-302">Requirements</span></span>

|<span data-ttu-id="b877b-303">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-303">Requirement</span></span>| <span data-ttu-id="b877b-304">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-305">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-306">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-306">1.0</span></span>|
|[<span data-ttu-id="b877b-307">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-308">ReadItem</span></span>|
|[<span data-ttu-id="b877b-309">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-310">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-311">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b877b-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b877b-312">(nullable) itemId :String</span></span>

<span data-ttu-id="b877b-p118">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-315">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="b877b-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b877b-316">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="b877b-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b877b-317">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью метода `Office.context.mailbox.convertToRestId`, который доступен в наборе обязательных элементов, начиная с версии 1.3.</span><span class="sxs-lookup"><span data-stu-id="b877b-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="b877b-318">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="b877b-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-319">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-319">Type:</span></span>

*   <span data-ttu-id="b877b-320">String</span><span class="sxs-lookup"><span data-stu-id="b877b-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-321">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-321">Requirements</span></span>

|<span data-ttu-id="b877b-322">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-322">Requirement</span></span>| <span data-ttu-id="b877b-323">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-324">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-325">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-325">1.0</span></span>|
|[<span data-ttu-id="b877b-326">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-327">ReadItem</span></span>|
|[<span data-ttu-id="b877b-328">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-329">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-330">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-330">Example</span></span>

<span data-ttu-id="b877b-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="b877b-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b877b-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b877b-334">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="b877b-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b877b-335">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="b877b-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-336">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-336">Type:</span></span>

*   [<span data-ttu-id="b877b-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b877b-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b877b-338">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-338">Requirements</span></span>

|<span data-ttu-id="b877b-339">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-339">Requirement</span></span>| <span data-ttu-id="b877b-340">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-341">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-342">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-342">1.0</span></span>|
|[<span data-ttu-id="b877b-343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-344">ReadItem</span></span>|
|[<span data-ttu-id="b877b-345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-346">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-347">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="b877b-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="b877b-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="b877b-349">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="b877b-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b877b-350">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b877b-350">Read mode</span></span>

<span data-ttu-id="b877b-351">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="b877b-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b877b-352">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b877b-352">Compose mode</span></span>

<span data-ttu-id="b877b-353">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="b877b-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-354">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-354">Type:</span></span>

*   <span data-ttu-id="b877b-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="b877b-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-356">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-356">Requirements</span></span>

|<span data-ttu-id="b877b-357">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-357">Requirement</span></span>| <span data-ttu-id="b877b-358">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-359">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-360">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-360">1.0</span></span>|
|[<span data-ttu-id="b877b-361">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-362">ReadItem</span></span>|
|[<span data-ttu-id="b877b-363">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-364">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-365">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b877b-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b877b-366">normalizedSubject :String</span></span>

<span data-ttu-id="b877b-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b877b-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="b877b-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-371">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-371">Type:</span></span>

*   <span data-ttu-id="b877b-372">String</span><span class="sxs-lookup"><span data-stu-id="b877b-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-373">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-373">Requirements</span></span>

|<span data-ttu-id="b877b-374">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-374">Requirement</span></span>| <span data-ttu-id="b877b-375">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-376">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-377">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-377">1.0</span></span>|
|[<span data-ttu-id="b877b-378">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-379">ReadItem</span></span>|
|[<span data-ttu-id="b877b-380">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-381">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-382">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="b877b-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b877b-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="b877b-384">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="b877b-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b877b-385">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b877b-386">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b877b-386">Read mode</span></span>

<span data-ttu-id="b877b-387">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="b877b-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b877b-388">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b877b-388">Compose mode</span></span>

<span data-ttu-id="b877b-389">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="b877b-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-390">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-390">Type:</span></span>

*   <span data-ttu-id="b877b-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b877b-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-392">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-392">Requirements</span></span>

|<span data-ttu-id="b877b-393">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-393">Requirement</span></span>| <span data-ttu-id="b877b-394">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-395">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-396">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-396">1.0</span></span>|
|[<span data-ttu-id="b877b-397">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-398">ReadItem</span></span>|
|[<span data-ttu-id="b877b-399">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-400">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-401">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="b877b-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b877b-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="b877b-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-405">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-405">Type:</span></span>

*   [<span data-ttu-id="b877b-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b877b-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b877b-407">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-407">Requirements</span></span>

|<span data-ttu-id="b877b-408">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-408">Requirement</span></span>| <span data-ttu-id="b877b-409">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-410">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-411">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-411">1.0</span></span>|
|[<span data-ttu-id="b877b-412">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-413">ReadItem</span></span>|
|[<span data-ttu-id="b877b-414">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-415">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-416">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="b877b-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b877b-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="b877b-418">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="b877b-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b877b-419">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b877b-420">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b877b-420">Read mode</span></span>

<span data-ttu-id="b877b-421">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="b877b-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b877b-422">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b877b-422">Compose mode</span></span>

<span data-ttu-id="b877b-423">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="b877b-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-424">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-424">Type:</span></span>

*   <span data-ttu-id="b877b-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b877b-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-426">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-426">Requirements</span></span>

|<span data-ttu-id="b877b-427">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-427">Requirement</span></span>| <span data-ttu-id="b877b-428">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-429">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-430">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-430">1.0</span></span>|
|[<span data-ttu-id="b877b-431">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-432">ReadItem</span></span>|
|[<span data-ttu-id="b877b-433">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-434">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-435">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="b877b-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b877b-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="b877b-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b877b-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b877b-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="b877b-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-441">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b877b-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-442">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-442">Type:</span></span>

*   [<span data-ttu-id="b877b-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b877b-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b877b-444">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-444">Requirements</span></span>

|<span data-ttu-id="b877b-445">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-445">Requirement</span></span>| <span data-ttu-id="b877b-446">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-447">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-448">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-448">1.0</span></span>|
|[<span data-ttu-id="b877b-449">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-450">ReadItem</span></span>|
|[<span data-ttu-id="b877b-451">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-452">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-453">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="b877b-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="b877b-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="b877b-455">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="b877b-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b877b-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="b877b-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b877b-458">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b877b-458">Read mode</span></span>

<span data-ttu-id="b877b-459">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="b877b-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b877b-460">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b877b-460">Compose mode</span></span>

<span data-ttu-id="b877b-461">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="b877b-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b877b-462">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="b877b-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-463">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-463">Type:</span></span>

*   <span data-ttu-id="b877b-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="b877b-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-465">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-465">Requirements</span></span>

|<span data-ttu-id="b877b-466">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-466">Requirement</span></span>| <span data-ttu-id="b877b-467">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-468">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-469">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-469">1.0</span></span>|
|[<span data-ttu-id="b877b-470">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-471">ReadItem</span></span>|
|[<span data-ttu-id="b877b-472">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-473">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-474">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-474">Example</span></span>

<span data-ttu-id="b877b-475">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b877b-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="b877b-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b877b-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="b877b-477">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b877b-478">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="b877b-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b877b-479">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b877b-479">Read mode</span></span>

<span data-ttu-id="b877b-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="b877b-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="b877b-482">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b877b-482">Compose mode</span></span>

<span data-ttu-id="b877b-483">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="b877b-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b877b-484">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-484">Type:</span></span>

*   <span data-ttu-id="b877b-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b877b-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-486">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-486">Requirements</span></span>

|<span data-ttu-id="b877b-487">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-487">Requirement</span></span>| <span data-ttu-id="b877b-488">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-489">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-490">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-490">1.0</span></span>|
|[<span data-ttu-id="b877b-491">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-492">ReadItem</span></span>|
|[<span data-ttu-id="b877b-493">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-494">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="b877b-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b877b-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="b877b-496">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b877b-497">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b877b-498">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b877b-498">Read mode</span></span>

<span data-ttu-id="b877b-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="b877b-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b877b-501">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b877b-501">Compose mode</span></span>

<span data-ttu-id="b877b-502">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b877b-503">Тип:</span><span class="sxs-lookup"><span data-stu-id="b877b-503">Type:</span></span>

*   <span data-ttu-id="b877b-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b877b-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-505">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-505">Requirements</span></span>

|<span data-ttu-id="b877b-506">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-506">Requirement</span></span>| <span data-ttu-id="b877b-507">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-508">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-509">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-509">1.0</span></span>|
|[<span data-ttu-id="b877b-510">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-511">ReadItem</span></span>|
|[<span data-ttu-id="b877b-512">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-513">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-514">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="b877b-515">Методы</span><span class="sxs-lookup"><span data-stu-id="b877b-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b877b-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b877b-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b877b-517">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="b877b-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b877b-518">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="b877b-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b877b-519">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="b877b-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b877b-520">Параметры</span><span class="sxs-lookup"><span data-stu-id="b877b-520">Parameters:</span></span>

|<span data-ttu-id="b877b-521">Имя</span><span class="sxs-lookup"><span data-stu-id="b877b-521">Name</span></span>| <span data-ttu-id="b877b-522">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-522">Type</span></span>| <span data-ttu-id="b877b-523">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b877b-523">Attributes</span></span>| <span data-ttu-id="b877b-524">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="b877b-525">String</span><span class="sxs-lookup"><span data-stu-id="b877b-525">String</span></span>||<span data-ttu-id="b877b-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="b877b-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b877b-528">String</span><span class="sxs-lookup"><span data-stu-id="b877b-528">String</span></span>||<span data-ttu-id="b877b-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b877b-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b877b-531">Object</span><span class="sxs-lookup"><span data-stu-id="b877b-531">Object</span></span>| <span data-ttu-id="b877b-532">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-532">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-533">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b877b-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b877b-534">Object</span><span class="sxs-lookup"><span data-stu-id="b877b-534">Object</span></span>| <span data-ttu-id="b877b-535">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-535">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-536">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b877b-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b877b-537">функция</span><span class="sxs-lookup"><span data-stu-id="b877b-537">function</span></span>| <span data-ttu-id="b877b-538">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-538">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-539">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b877b-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b877b-540">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b877b-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b877b-541">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="b877b-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b877b-542">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b877b-542">Errors</span></span>

| <span data-ttu-id="b877b-543">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b877b-543">Error code</span></span> | <span data-ttu-id="b877b-544">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="b877b-545">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="b877b-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="b877b-546">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="b877b-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b877b-547">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="b877b-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b877b-548">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-548">Requirements</span></span>

|<span data-ttu-id="b877b-549">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-549">Requirement</span></span>| <span data-ttu-id="b877b-550">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-551">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-552">1.1</span><span class="sxs-lookup"><span data-stu-id="b877b-552">1.1</span></span>|
|[<span data-ttu-id="b877b-553">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b877b-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="b877b-555">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-556">Создание</span><span class="sxs-lookup"><span data-stu-id="b877b-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-557">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b877b-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b877b-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b877b-559">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="b877b-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b877b-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b877b-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b877b-563">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="b877b-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b877b-564">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="b877b-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b877b-565">Параметры</span><span class="sxs-lookup"><span data-stu-id="b877b-565">Parameters:</span></span>

|<span data-ttu-id="b877b-566">Имя</span><span class="sxs-lookup"><span data-stu-id="b877b-566">Name</span></span>| <span data-ttu-id="b877b-567">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-567">Type</span></span>| <span data-ttu-id="b877b-568">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b877b-568">Attributes</span></span>| <span data-ttu-id="b877b-569">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="b877b-570">String</span><span class="sxs-lookup"><span data-stu-id="b877b-570">String</span></span>||<span data-ttu-id="b877b-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="b877b-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b877b-573">String</span><span class="sxs-lookup"><span data-stu-id="b877b-573">String</span></span>||<span data-ttu-id="b877b-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b877b-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b877b-576">Object</span><span class="sxs-lookup"><span data-stu-id="b877b-576">Object</span></span>| <span data-ttu-id="b877b-577">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-577">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-578">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b877b-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b877b-579">Object</span><span class="sxs-lookup"><span data-stu-id="b877b-579">Object</span></span>| <span data-ttu-id="b877b-580">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-580">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-581">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b877b-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b877b-582">функция</span><span class="sxs-lookup"><span data-stu-id="b877b-582">function</span></span>| <span data-ttu-id="b877b-583">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-583">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-584">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b877b-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b877b-585">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b877b-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b877b-586">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="b877b-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b877b-587">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b877b-587">Errors</span></span>

| <span data-ttu-id="b877b-588">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b877b-588">Error code</span></span> | <span data-ttu-id="b877b-589">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b877b-590">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="b877b-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b877b-591">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-591">Requirements</span></span>

|<span data-ttu-id="b877b-592">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-592">Requirement</span></span>| <span data-ttu-id="b877b-593">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-594">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-595">1.1</span><span class="sxs-lookup"><span data-stu-id="b877b-595">1.1</span></span>|
|[<span data-ttu-id="b877b-596">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b877b-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="b877b-598">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-599">Создание</span><span class="sxs-lookup"><span data-stu-id="b877b-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-600">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-600">Example</span></span>

<span data-ttu-id="b877b-601">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="b877b-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="b877b-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b877b-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="b877b-603">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="b877b-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-604">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b877b-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b877b-605">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="b877b-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b877b-606">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="b877b-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-607">Набор обязательных элементов 1.1 не поддерживает возможность включения вложений при вызове `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="b877b-607">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="b877b-608">Поддержка вложений была добавлена для `displayReplyAllForm` в наборе обязательных элементов 1.2 и более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="b877b-608">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b877b-609">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b877b-609">Parameters:</span></span>

|<span data-ttu-id="b877b-610">Имя</span><span class="sxs-lookup"><span data-stu-id="b877b-610">Name</span></span>| <span data-ttu-id="b877b-611">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-611">Type</span></span>| <span data-ttu-id="b877b-612">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b877b-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b877b-613">String &#124; Object</span></span>| |<span data-ttu-id="b877b-p138">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b877b-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b877b-616">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="b877b-616">**OR**</span></span><br/><span data-ttu-id="b877b-p139">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="b877b-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b877b-619">String</span><span class="sxs-lookup"><span data-stu-id="b877b-619">String</span></span> | <span data-ttu-id="b877b-620">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-620">&lt;optional&gt;</span></span> | <span data-ttu-id="b877b-p140">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="b877b-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="b877b-623">функция</span><span class="sxs-lookup"><span data-stu-id="b877b-623">function</span></span> | <span data-ttu-id="b877b-624">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-624">&lt;optional&gt;</span></span> | <span data-ttu-id="b877b-625">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b877b-625">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b877b-626">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-626">Requirements</span></span>

|<span data-ttu-id="b877b-627">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-627">Requirement</span></span>| <span data-ttu-id="b877b-628">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-629">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-630">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-630">1.0</span></span>|
|[<span data-ttu-id="b877b-631">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-632">ReadItem</span></span>|
|[<span data-ttu-id="b877b-633">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-634">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-634">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b877b-635">Примеры</span><span class="sxs-lookup"><span data-stu-id="b877b-635">Examples</span></span>

<span data-ttu-id="b877b-636">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="b877b-636">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b877b-637">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-637">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b877b-638">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-638">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b877b-639">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="b877b-639">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="b877b-640">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b877b-640">displayReplyForm(formData)</span></span>

<span data-ttu-id="b877b-641">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="b877b-641">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-642">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b877b-642">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b877b-643">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="b877b-643">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b877b-644">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="b877b-644">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-645">Набор обязательных элементов 1.1 не поддерживает возможность включения вложений при вызове `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="b877b-645">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="b877b-646">Поддержка вложений была добавлена для `displayReplyForm` в наборе обязательных элементов 1.2 и более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="b877b-646">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b877b-647">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b877b-647">Parameters:</span></span>

|<span data-ttu-id="b877b-648">Имя</span><span class="sxs-lookup"><span data-stu-id="b877b-648">Name</span></span>| <span data-ttu-id="b877b-649">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-649">Type</span></span>| <span data-ttu-id="b877b-650">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-650">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b877b-651">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b877b-651">String &#124; Object</span></span>| | <span data-ttu-id="b877b-p142">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b877b-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b877b-654">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="b877b-654">**OR**</span></span><br/><span data-ttu-id="b877b-p143">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="b877b-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b877b-657">String</span><span class="sxs-lookup"><span data-stu-id="b877b-657">String</span></span> | <span data-ttu-id="b877b-658">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-658">&lt;optional&gt;</span></span> | <span data-ttu-id="b877b-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="b877b-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="b877b-661">функция</span><span class="sxs-lookup"><span data-stu-id="b877b-661">function</span></span> | <span data-ttu-id="b877b-662">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-662">&lt;optional&gt;</span></span> | <span data-ttu-id="b877b-663">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b877b-663">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b877b-664">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-664">Requirements</span></span>

|<span data-ttu-id="b877b-665">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-665">Requirement</span></span>| <span data-ttu-id="b877b-666">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-667">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-667">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-668">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-668">1.0</span></span>|
|[<span data-ttu-id="b877b-669">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-670">ReadItem</span></span>|
|[<span data-ttu-id="b877b-671">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-672">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-672">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b877b-673">Примеры</span><span class="sxs-lookup"><span data-stu-id="b877b-673">Examples</span></span>

<span data-ttu-id="b877b-674">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="b877b-674">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b877b-675">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-675">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b877b-676">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b877b-676">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b877b-677">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="b877b-677">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="b877b-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b877b-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="b877b-679">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-679">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-680">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b877b-680">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-681">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-681">Requirements</span></span>

|<span data-ttu-id="b877b-682">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-682">Requirement</span></span>| <span data-ttu-id="b877b-683">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-683">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-684">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-684">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-685">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-685">1.0</span></span>|
|[<span data-ttu-id="b877b-686">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-686">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-687">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-687">ReadItem</span></span>|
|[<span data-ttu-id="b877b-688">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-688">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-689">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-689">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b877b-690">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b877b-690">Returns:</span></span>

<span data-ttu-id="b877b-691">Тип: [Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b877b-691">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b877b-692">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-692">Example</span></span>

<span data-ttu-id="b877b-693">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-693">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="b877b-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b877b-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b877b-695">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-695">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-696">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b877b-696">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b877b-697">Параметры</span><span class="sxs-lookup"><span data-stu-id="b877b-697">Parameters:</span></span>

|<span data-ttu-id="b877b-698">Имя</span><span class="sxs-lookup"><span data-stu-id="b877b-698">Name</span></span>| <span data-ttu-id="b877b-699">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-699">Type</span></span>| <span data-ttu-id="b877b-700">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-700">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="b877b-701">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b877b-701">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="b877b-702">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="b877b-702">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b877b-703">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-703">Requirements</span></span>

|<span data-ttu-id="b877b-704">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-704">Requirement</span></span>| <span data-ttu-id="b877b-705">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-705">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-706">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-706">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-707">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-707">1.0</span></span>|
|[<span data-ttu-id="b877b-708">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-708">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-709">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="b877b-709">Restricted</span></span>|
|[<span data-ttu-id="b877b-710">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-710">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-711">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-711">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b877b-712">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b877b-712">Returns:</span></span>

<span data-ttu-id="b877b-713">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="b877b-713">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b877b-714">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="b877b-714">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="b877b-715">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="b877b-715">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b877b-716">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="b877b-716">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="b877b-717">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="b877b-717">Value of `entityType`</span></span> | <span data-ttu-id="b877b-718">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="b877b-718">Type of objects in returned array</span></span> | <span data-ttu-id="b877b-719">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-719">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="b877b-720">String</span><span class="sxs-lookup"><span data-stu-id="b877b-720">String</span></span> | <span data-ttu-id="b877b-721">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b877b-721">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="b877b-722">Contact</span><span class="sxs-lookup"><span data-stu-id="b877b-722">Contact</span></span> | <span data-ttu-id="b877b-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b877b-723">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="b877b-724">String</span><span class="sxs-lookup"><span data-stu-id="b877b-724">String</span></span> | <span data-ttu-id="b877b-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b877b-725">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="b877b-726">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b877b-726">MeetingSuggestion</span></span> | <span data-ttu-id="b877b-727">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b877b-727">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="b877b-728">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b877b-728">PhoneNumber</span></span> | <span data-ttu-id="b877b-729">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b877b-729">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="b877b-730">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b877b-730">TaskSuggestion</span></span> | <span data-ttu-id="b877b-731">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b877b-731">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="b877b-732">String</span><span class="sxs-lookup"><span data-stu-id="b877b-732">String</span></span> | <span data-ttu-id="b877b-733">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b877b-733">**Restricted**</span></span> |

<span data-ttu-id="b877b-734">Тип:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b877b-734">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="b877b-735">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-735">Example</span></span>

<span data-ttu-id="b877b-736">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-736">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="b877b-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b877b-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b877b-738">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b877b-738">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-739">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b877b-739">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b877b-740">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="b877b-740">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b877b-741">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b877b-741">Parameters:</span></span>

|<span data-ttu-id="b877b-742">Имя</span><span class="sxs-lookup"><span data-stu-id="b877b-742">Name</span></span>| <span data-ttu-id="b877b-743">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-743">Type</span></span>| <span data-ttu-id="b877b-744">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-744">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b877b-745">String</span><span class="sxs-lookup"><span data-stu-id="b877b-745">String</span></span>|<span data-ttu-id="b877b-746">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="b877b-746">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b877b-747">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-747">Requirements</span></span>

|<span data-ttu-id="b877b-748">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-748">Requirement</span></span>| <span data-ttu-id="b877b-749">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-750">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-751">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-751">1.0</span></span>|
|[<span data-ttu-id="b877b-752">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-752">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-753">ReadItem</span></span>|
|[<span data-ttu-id="b877b-754">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-754">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-755">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-755">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b877b-756">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b877b-756">Returns:</span></span>

<span data-ttu-id="b877b-p146">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="b877b-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="b877b-759">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b877b-759">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="b877b-760">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b877b-760">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b877b-761">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b877b-761">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-762">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b877b-762">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b877b-p147">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="b877b-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b877b-766">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="b877b-766">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b877b-767">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="b877b-767">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="b877b-p148">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="b877b-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b877b-770">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-770">Requirements</span></span>

|<span data-ttu-id="b877b-771">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-771">Requirement</span></span>| <span data-ttu-id="b877b-772">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-773">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-774">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-774">1.0</span></span>|
|[<span data-ttu-id="b877b-775">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-775">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-776">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-776">ReadItem</span></span>|
|[<span data-ttu-id="b877b-777">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-777">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-778">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b877b-779">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b877b-779">Returns:</span></span>

<span data-ttu-id="b877b-p149">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="b877b-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b877b-782">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="b877b-782">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b877b-783">Object</span><span class="sxs-lookup"><span data-stu-id="b877b-783">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b877b-784">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-784">Example</span></span>

<span data-ttu-id="b877b-785">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="b877b-785">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b877b-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="b877b-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b877b-787">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b877b-787">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b877b-788">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b877b-788">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b877b-789">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="b877b-789">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b877b-p150">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="b877b-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b877b-792">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b877b-792">Parameters:</span></span>

|<span data-ttu-id="b877b-793">Имя</span><span class="sxs-lookup"><span data-stu-id="b877b-793">Name</span></span>| <span data-ttu-id="b877b-794">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-794">Type</span></span>| <span data-ttu-id="b877b-795">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-795">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b877b-796">String</span><span class="sxs-lookup"><span data-stu-id="b877b-796">String</span></span>|<span data-ttu-id="b877b-797">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="b877b-797">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b877b-798">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-798">Requirements</span></span>

|<span data-ttu-id="b877b-799">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-799">Requirement</span></span>| <span data-ttu-id="b877b-800">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-800">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-801">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-801">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-802">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-802">1.0</span></span>|
|[<span data-ttu-id="b877b-803">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-803">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-804">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-804">ReadItem</span></span>|
|[<span data-ttu-id="b877b-805">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-805">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-806">Чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-806">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b877b-807">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b877b-807">Returns:</span></span>

<span data-ttu-id="b877b-808">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b877b-808">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b877b-809">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="b877b-809">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b877b-810">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="b877b-810">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b877b-811">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-811">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b877b-812">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b877b-812">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b877b-813">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="b877b-813">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b877b-p151">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="b877b-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b877b-817">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b877b-817">Parameters:</span></span>

|<span data-ttu-id="b877b-818">Имя</span><span class="sxs-lookup"><span data-stu-id="b877b-818">Name</span></span>| <span data-ttu-id="b877b-819">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-819">Type</span></span>| <span data-ttu-id="b877b-820">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b877b-820">Attributes</span></span>| <span data-ttu-id="b877b-821">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-821">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b877b-822">функция</span><span class="sxs-lookup"><span data-stu-id="b877b-822">function</span></span>||<span data-ttu-id="b877b-823">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b877b-823">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b877b-824">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b877b-824">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b877b-825">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="b877b-825">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="b877b-826">Объект</span><span class="sxs-lookup"><span data-stu-id="b877b-826">Object</span></span>| <span data-ttu-id="b877b-827">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-827">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-828">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b877b-828">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="b877b-829">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b877b-829">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b877b-830">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-830">Requirements</span></span>

|<span data-ttu-id="b877b-831">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-831">Requirement</span></span>| <span data-ttu-id="b877b-832">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-832">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-833">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b877b-833">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-834">1.0</span><span class="sxs-lookup"><span data-stu-id="b877b-834">1.0</span></span>|
|[<span data-ttu-id="b877b-835">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-835">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-836">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b877b-836">ReadItem</span></span>|
|[<span data-ttu-id="b877b-837">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-837">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-838">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b877b-838">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-839">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-839">Example</span></span>

<span data-ttu-id="b877b-p154">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="b877b-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b877b-843">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b877b-843">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b877b-844">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="b877b-844">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b877b-p155">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="b877b-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b877b-849">Параметры</span><span class="sxs-lookup"><span data-stu-id="b877b-849">Parameters:</span></span>

|<span data-ttu-id="b877b-850">Имя</span><span class="sxs-lookup"><span data-stu-id="b877b-850">Name</span></span>| <span data-ttu-id="b877b-851">Тип</span><span class="sxs-lookup"><span data-stu-id="b877b-851">Type</span></span>| <span data-ttu-id="b877b-852">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b877b-852">Attributes</span></span>| <span data-ttu-id="b877b-853">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-853">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="b877b-854">String</span><span class="sxs-lookup"><span data-stu-id="b877b-854">String</span></span>||<span data-ttu-id="b877b-855">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="b877b-855">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="b877b-856">Object</span><span class="sxs-lookup"><span data-stu-id="b877b-856">Object</span></span>| <span data-ttu-id="b877b-857">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-857">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-858">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b877b-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b877b-859">Object</span><span class="sxs-lookup"><span data-stu-id="b877b-859">Object</span></span>| <span data-ttu-id="b877b-860">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-860">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-861">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b877b-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b877b-862">функция</span><span class="sxs-lookup"><span data-stu-id="b877b-862">function</span></span>| <span data-ttu-id="b877b-863">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b877b-863">&lt;optional&gt;</span></span>|<span data-ttu-id="b877b-864">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b877b-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b877b-865">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="b877b-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b877b-866">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b877b-866">Errors</span></span>

| <span data-ttu-id="b877b-867">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b877b-867">Error code</span></span> | <span data-ttu-id="b877b-868">Описание</span><span class="sxs-lookup"><span data-stu-id="b877b-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="b877b-869">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="b877b-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b877b-870">Требования</span><span class="sxs-lookup"><span data-stu-id="b877b-870">Requirements</span></span>

|<span data-ttu-id="b877b-871">Требование</span><span class="sxs-lookup"><span data-stu-id="b877b-871">Requirement</span></span>| <span data-ttu-id="b877b-872">Значение</span><span class="sxs-lookup"><span data-stu-id="b877b-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="b877b-873">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b877b-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b877b-874">1.1</span><span class="sxs-lookup"><span data-stu-id="b877b-874">1.1</span></span>|
|[<span data-ttu-id="b877b-875">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b877b-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b877b-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b877b-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="b877b-877">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b877b-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b877b-878">Создание</span><span class="sxs-lookup"><span data-stu-id="b877b-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b877b-879">Пример</span><span class="sxs-lookup"><span data-stu-id="b877b-879">Example</span></span>

<span data-ttu-id="b877b-880">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="b877b-880">The following code removes an attachment with an identifier of '0'.</span></span>

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
