---
title: Office. Context. Mailbox. Item — набор требований 1,1
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5cbf942ea9b1351e0f945a9ca5534a9ba090b79b
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001616"
---
# <a name="item"></a><span data-ttu-id="14c88-102">item</span><span class="sxs-lookup"><span data-stu-id="14c88-102">item</span></span>

### <span data-ttu-id="14c88-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="14c88-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="14c88-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="14c88-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-107">Requirements</span></span>

|<span data-ttu-id="14c88-108">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-108">Requirement</span></span>| <span data-ttu-id="14c88-109">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-111">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-111">1.0</span></span>|
|[<span data-ttu-id="14c88-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="14c88-113">Restricted</span></span>|
|[<span data-ttu-id="14c88-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="14c88-116">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="14c88-116">Members and methods</span></span>

| <span data-ttu-id="14c88-117">Элемент	</span><span class="sxs-lookup"><span data-stu-id="14c88-117">Member</span></span> | <span data-ttu-id="14c88-118">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="14c88-119">attachments</span><span class="sxs-lookup"><span data-stu-id="14c88-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="14c88-120">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-120">Member</span></span> |
| [<span data-ttu-id="14c88-121">bcc</span><span class="sxs-lookup"><span data-stu-id="14c88-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="14c88-122">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-122">Member</span></span> |
| [<span data-ttu-id="14c88-123">body</span><span class="sxs-lookup"><span data-stu-id="14c88-123">body</span></span>](#body-body) | <span data-ttu-id="14c88-124">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-124">Member</span></span> |
| [<span data-ttu-id="14c88-125">cc</span><span class="sxs-lookup"><span data-stu-id="14c88-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="14c88-126">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-126">Member</span></span> |
| [<span data-ttu-id="14c88-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="14c88-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="14c88-128">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-128">Member</span></span> |
| [<span data-ttu-id="14c88-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="14c88-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="14c88-130">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-130">Member</span></span> |
| [<span data-ttu-id="14c88-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="14c88-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="14c88-132">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-132">Member</span></span> |
| [<span data-ttu-id="14c88-133">end</span><span class="sxs-lookup"><span data-stu-id="14c88-133">end</span></span>](#end-datetime) | <span data-ttu-id="14c88-134">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-134">Member</span></span> |
| [<span data-ttu-id="14c88-135">from</span><span class="sxs-lookup"><span data-stu-id="14c88-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="14c88-136">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-136">Member</span></span> |
| [<span data-ttu-id="14c88-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="14c88-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="14c88-138">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-138">Member</span></span> |
| [<span data-ttu-id="14c88-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="14c88-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="14c88-140">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-140">Member</span></span> |
| [<span data-ttu-id="14c88-141">itemId</span><span class="sxs-lookup"><span data-stu-id="14c88-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="14c88-142">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-142">Member</span></span> |
| [<span data-ttu-id="14c88-143">itemType</span><span class="sxs-lookup"><span data-stu-id="14c88-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="14c88-144">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-144">Member</span></span> |
| [<span data-ttu-id="14c88-145">location</span><span class="sxs-lookup"><span data-stu-id="14c88-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="14c88-146">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-146">Member</span></span> |
| [<span data-ttu-id="14c88-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="14c88-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="14c88-148">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-148">Member</span></span> |
| [<span data-ttu-id="14c88-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="14c88-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="14c88-150">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-150">Member</span></span> |
| [<span data-ttu-id="14c88-151">organizer</span><span class="sxs-lookup"><span data-stu-id="14c88-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="14c88-152">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-152">Member</span></span> |
| [<span data-ttu-id="14c88-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="14c88-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="14c88-154">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-154">Member</span></span> |
| [<span data-ttu-id="14c88-155">sender</span><span class="sxs-lookup"><span data-stu-id="14c88-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="14c88-156">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-156">Member</span></span> |
| [<span data-ttu-id="14c88-157">start</span><span class="sxs-lookup"><span data-stu-id="14c88-157">start</span></span>](#start-datetime) | <span data-ttu-id="14c88-158">Member</span><span class="sxs-lookup"><span data-stu-id="14c88-158">Member</span></span> |
| [<span data-ttu-id="14c88-159">subject</span><span class="sxs-lookup"><span data-stu-id="14c88-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="14c88-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="14c88-160">Member</span></span> |
| [<span data-ttu-id="14c88-161">to</span><span class="sxs-lookup"><span data-stu-id="14c88-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="14c88-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="14c88-162">Member</span></span> |
| [<span data-ttu-id="14c88-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="14c88-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="14c88-164">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-164">Method</span></span> |
| [<span data-ttu-id="14c88-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="14c88-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="14c88-166">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-166">Method</span></span> |
| [<span data-ttu-id="14c88-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="14c88-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="14c88-168">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-168">Method</span></span> |
| [<span data-ttu-id="14c88-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="14c88-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="14c88-170">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-170">Method</span></span> |
| [<span data-ttu-id="14c88-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="14c88-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="14c88-172">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-172">Method</span></span> |
| [<span data-ttu-id="14c88-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="14c88-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="14c88-174">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-174">Method</span></span> |
| [<span data-ttu-id="14c88-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="14c88-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="14c88-176">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-176">Method</span></span> |
| [<span data-ttu-id="14c88-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="14c88-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="14c88-178">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-178">Method</span></span> |
| [<span data-ttu-id="14c88-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="14c88-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="14c88-180">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-180">Method</span></span> |
| [<span data-ttu-id="14c88-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="14c88-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="14c88-182">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-182">Method</span></span> |
| [<span data-ttu-id="14c88-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="14c88-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="14c88-184">Метод</span><span class="sxs-lookup"><span data-stu-id="14c88-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="14c88-185">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-185">Example</span></span>

<span data-ttu-id="14c88-186">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="14c88-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="14c88-187">Members</span><span class="sxs-lookup"><span data-stu-id="14c88-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="14c88-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="14c88-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="14c88-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-191">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="14c88-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="14c88-192">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="14c88-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-193">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-193">Type</span></span>

*   <span data-ttu-id="14c88-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="14c88-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-195">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-195">Requirements</span></span>

|<span data-ttu-id="14c88-196">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-196">Requirement</span></span>| <span data-ttu-id="14c88-197">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-198">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-199">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-199">1.0</span></span>|
|[<span data-ttu-id="14c88-200">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-201">ReadItem</span></span>|
|[<span data-ttu-id="14c88-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-203">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-204">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-204">Example</span></span>

<span data-ttu-id="14c88-205">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14c88-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-207">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="14c88-208">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="14c88-208">Compose mode only.</span></span>

<span data-ttu-id="14c88-209">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="14c88-209">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14c88-210">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="14c88-210">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14c88-211">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-211">Get 500 members maximum.</span></span>
- <span data-ttu-id="14c88-212">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-212">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-213">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-213">Type</span></span>

*   [<span data-ttu-id="14c88-214">Получатели</span><span class="sxs-lookup"><span data-stu-id="14c88-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14c88-215">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-215">Requirements</span></span>

|<span data-ttu-id="14c88-216">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-216">Requirement</span></span>| <span data-ttu-id="14c88-217">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-218">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-219">1.1</span><span class="sxs-lookup"><span data-stu-id="14c88-219">1.1</span></span>|
|[<span data-ttu-id="14c88-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-221">ReadItem</span></span>|
|[<span data-ttu-id="14c88-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-223">Создание</span><span class="sxs-lookup"><span data-stu-id="14c88-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-224">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-224">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="14c88-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-226">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-227">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-227">Type</span></span>

*   [<span data-ttu-id="14c88-228">Body</span><span class="sxs-lookup"><span data-stu-id="14c88-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14c88-229">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-229">Requirements</span></span>

|<span data-ttu-id="14c88-230">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-230">Requirement</span></span>| <span data-ttu-id="14c88-231">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-232">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-233">1.1</span><span class="sxs-lookup"><span data-stu-id="14c88-233">1.1</span></span>|
|[<span data-ttu-id="14c88-234">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-235">ReadItem</span></span>|
|[<span data-ttu-id="14c88-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-238">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-238">Example</span></span>

<span data-ttu-id="14c88-239">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="14c88-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="14c88-240">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="14c88-240">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14c88-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-242">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="14c88-243">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14c88-244">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="14c88-244">Read mode</span></span>

<span data-ttu-id="14c88-245">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-245">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="14c88-246">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="14c88-246">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14c88-247">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="14c88-247">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="14c88-248">Режим создания</span><span class="sxs-lookup"><span data-stu-id="14c88-248">Compose mode</span></span>

<span data-ttu-id="14c88-249">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="14c88-250">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="14c88-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14c88-251">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="14c88-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14c88-252">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="14c88-253">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="14c88-254">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-254">Type</span></span>

*   <span data-ttu-id="14c88-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-256">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-256">Requirements</span></span>

|<span data-ttu-id="14c88-257">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-257">Requirement</span></span>| <span data-ttu-id="14c88-258">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-259">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-260">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-260">1.0</span></span>|
|[<span data-ttu-id="14c88-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-262">ReadItem</span></span>|
|[<span data-ttu-id="14c88-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="14c88-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="14c88-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="14c88-266">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="14c88-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="14c88-p110">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="14c88-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="14c88-p111">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="14c88-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-271">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-271">Type</span></span>

*   <span data-ttu-id="14c88-272">String</span><span class="sxs-lookup"><span data-stu-id="14c88-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-273">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-273">Requirements</span></span>

|<span data-ttu-id="14c88-274">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-274">Requirement</span></span>| <span data-ttu-id="14c88-275">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-276">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-277">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-277">1.0</span></span>|
|[<span data-ttu-id="14c88-278">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-279">ReadItem</span></span>|
|[<span data-ttu-id="14c88-280">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-281">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-282">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="14c88-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="14c88-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="14c88-p112">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-286">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-286">Type</span></span>

*   <span data-ttu-id="14c88-287">Дата</span><span class="sxs-lookup"><span data-stu-id="14c88-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-288">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-288">Requirements</span></span>

|<span data-ttu-id="14c88-289">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-289">Requirement</span></span>| <span data-ttu-id="14c88-290">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-291">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-292">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-292">1.0</span></span>|
|[<span data-ttu-id="14c88-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-294">ReadItem</span></span>|
|[<span data-ttu-id="14c88-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-296">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-297">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="14c88-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="14c88-298">dateTimeModified: Date</span></span>

<span data-ttu-id="14c88-p113">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-301">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14c88-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-302">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-302">Type</span></span>

*   <span data-ttu-id="14c88-303">Дата</span><span class="sxs-lookup"><span data-stu-id="14c88-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-304">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-304">Requirements</span></span>

|<span data-ttu-id="14c88-305">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-305">Requirement</span></span>| <span data-ttu-id="14c88-306">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-307">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-308">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-308">1.0</span></span>|
|[<span data-ttu-id="14c88-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-310">ReadItem</span></span>|
|[<span data-ttu-id="14c88-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-313">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="14c88-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-315">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="14c88-p114">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="14c88-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14c88-318">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="14c88-318">Read mode</span></span>

<span data-ttu-id="14c88-319">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="14c88-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="14c88-320">Режим создания</span><span class="sxs-lookup"><span data-stu-id="14c88-320">Compose mode</span></span>

<span data-ttu-id="14c88-321">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="14c88-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="14c88-322">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="14c88-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="14c88-323">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="14c88-324">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-324">Type</span></span>

*   <span data-ttu-id="14c88-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-326">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-326">Requirements</span></span>

|<span data-ttu-id="14c88-327">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-327">Requirement</span></span>| <span data-ttu-id="14c88-328">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-329">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-330">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-330">1.0</span></span>|
|[<span data-ttu-id="14c88-331">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-332">ReadItem</span></span>|
|[<span data-ttu-id="14c88-333">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-334">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="14c88-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-p115">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="14c88-p116">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="14c88-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-340">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="14c88-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-341">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-341">Type</span></span>

*   [<span data-ttu-id="14c88-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="14c88-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14c88-343">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-343">Requirements</span></span>

|<span data-ttu-id="14c88-344">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-344">Requirement</span></span>| <span data-ttu-id="14c88-345">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-346">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-347">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-347">1.0</span></span>|
|[<span data-ttu-id="14c88-348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-349">ReadItem</span></span>|
|[<span data-ttu-id="14c88-350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-351">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-352">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="14c88-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="14c88-353">internetMessageId: String</span></span>

<span data-ttu-id="14c88-p117">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-356">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-356">Type</span></span>

*   <span data-ttu-id="14c88-357">String</span><span class="sxs-lookup"><span data-stu-id="14c88-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-358">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-358">Requirements</span></span>

|<span data-ttu-id="14c88-359">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-359">Requirement</span></span>| <span data-ttu-id="14c88-360">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-361">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-362">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-362">1.0</span></span>|
|[<span data-ttu-id="14c88-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-364">ReadItem</span></span>|
|[<span data-ttu-id="14c88-365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-366">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-367">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="14c88-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="14c88-368">itemClass: String</span></span>

<span data-ttu-id="14c88-p118">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="14c88-p119">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="14c88-373">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-373">Type</span></span> | <span data-ttu-id="14c88-374">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-374">Description</span></span> | <span data-ttu-id="14c88-375">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="14c88-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="14c88-376">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="14c88-376">Appointment items</span></span> | <span data-ttu-id="14c88-377">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="14c88-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="14c88-378">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="14c88-378">Message items</span></span> | <span data-ttu-id="14c88-379">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="14c88-380">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="14c88-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-381">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-381">Type</span></span>

*   <span data-ttu-id="14c88-382">String</span><span class="sxs-lookup"><span data-stu-id="14c88-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-383">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-383">Requirements</span></span>

|<span data-ttu-id="14c88-384">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-384">Requirement</span></span>| <span data-ttu-id="14c88-385">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-386">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-387">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-387">1.0</span></span>|
|[<span data-ttu-id="14c88-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-389">ReadItem</span></span>|
|[<span data-ttu-id="14c88-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-392">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="14c88-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="14c88-393">(nullable) itemId: String</span></span>

<span data-ttu-id="14c88-394">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-394">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="14c88-395">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-395">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-396">Идентификатор, возвращаемый `itemId` свойством, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="14c88-396">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="14c88-397">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="14c88-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="14c88-398">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="14c88-398">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="14c88-399">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="14c88-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-400">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-400">Type</span></span>

*   <span data-ttu-id="14c88-401">String</span><span class="sxs-lookup"><span data-stu-id="14c88-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-402">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-402">Requirements</span></span>

|<span data-ttu-id="14c88-403">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-403">Requirement</span></span>| <span data-ttu-id="14c88-404">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-405">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-405">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-406">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-406">1.0</span></span>|
|[<span data-ttu-id="14c88-407">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-407">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-408">ReadItem</span></span>|
|[<span data-ttu-id="14c88-409">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-409">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-410">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-411">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-411">Example</span></span>

<span data-ttu-id="14c88-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="14c88-414">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-414">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-415">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="14c88-415">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="14c88-416">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="14c88-416">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-417">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-417">Type</span></span>

*   [<span data-ttu-id="14c88-418">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="14c88-418">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14c88-419">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-419">Requirements</span></span>

|<span data-ttu-id="14c88-420">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-420">Requirement</span></span>| <span data-ttu-id="14c88-421">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-421">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-422">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-423">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-423">1.0</span></span>|
|[<span data-ttu-id="14c88-424">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-424">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-425">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-425">ReadItem</span></span>|
|[<span data-ttu-id="14c88-426">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-426">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-427">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-427">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-428">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-428">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="14c88-429">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-429">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-430">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-430">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14c88-431">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="14c88-431">Read mode</span></span>

<span data-ttu-id="14c88-432">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-432">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="14c88-433">Режим создания</span><span class="sxs-lookup"><span data-stu-id="14c88-433">Compose mode</span></span>

<span data-ttu-id="14c88-434">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-434">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="14c88-435">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-435">Type</span></span>

*   <span data-ttu-id="14c88-436">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-436">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-437">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-437">Requirements</span></span>

|<span data-ttu-id="14c88-438">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-438">Requirement</span></span>| <span data-ttu-id="14c88-439">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-440">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-441">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-441">1.0</span></span>|
|[<span data-ttu-id="14c88-442">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-443">ReadItem</span></span>|
|[<span data-ttu-id="14c88-444">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-445">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-445">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="14c88-446">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="14c88-446">normalizedSubject: String</span></span>

<span data-ttu-id="14c88-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="14c88-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="14c88-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-451">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-451">Type</span></span>

*   <span data-ttu-id="14c88-452">String</span><span class="sxs-lookup"><span data-stu-id="14c88-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-453">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-453">Requirements</span></span>

|<span data-ttu-id="14c88-454">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-454">Requirement</span></span>| <span data-ttu-id="14c88-455">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-456">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-457">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-457">1.0</span></span>|
|[<span data-ttu-id="14c88-458">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-459">ReadItem</span></span>|
|[<span data-ttu-id="14c88-460">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-461">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-462">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14c88-463">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-463">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-464">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="14c88-464">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="14c88-465">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-465">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14c88-466">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="14c88-466">Read mode</span></span>

<span data-ttu-id="14c88-467">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="14c88-467">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="14c88-468">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="14c88-468">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14c88-469">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="14c88-469">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="14c88-470">Режим создания</span><span class="sxs-lookup"><span data-stu-id="14c88-470">Compose mode</span></span>

<span data-ttu-id="14c88-471">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="14c88-471">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="14c88-472">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="14c88-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14c88-473">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="14c88-473">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14c88-474">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-474">Get 500 members maximum.</span></span>
- <span data-ttu-id="14c88-475">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-475">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="14c88-476">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-476">Type</span></span>

*   <span data-ttu-id="14c88-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-478">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-478">Requirements</span></span>

|<span data-ttu-id="14c88-479">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-479">Requirement</span></span>| <span data-ttu-id="14c88-480">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-481">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-482">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-482">1.0</span></span>|
|[<span data-ttu-id="14c88-483">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-484">ReadItem</span></span>|
|[<span data-ttu-id="14c88-485">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-486">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-486">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="14c88-487">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-487">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-p128">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-490">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-490">Type</span></span>

*   [<span data-ttu-id="14c88-491">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="14c88-491">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14c88-492">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-492">Requirements</span></span>

|<span data-ttu-id="14c88-493">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-493">Requirement</span></span>| <span data-ttu-id="14c88-494">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-495">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-496">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-496">1.0</span></span>|
|[<span data-ttu-id="14c88-497">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-498">ReadItem</span></span>|
|[<span data-ttu-id="14c88-499">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-500">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-500">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-501">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-501">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14c88-502">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-502">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-503">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="14c88-503">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="14c88-504">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-504">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14c88-505">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="14c88-505">Read mode</span></span>

<span data-ttu-id="14c88-506">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="14c88-506">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="14c88-507">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="14c88-507">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14c88-508">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="14c88-508">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="14c88-509">Режим создания</span><span class="sxs-lookup"><span data-stu-id="14c88-509">Compose mode</span></span>

<span data-ttu-id="14c88-510">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="14c88-510">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="14c88-511">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="14c88-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14c88-512">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="14c88-512">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14c88-513">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-513">Get 500 members maximum.</span></span>
- <span data-ttu-id="14c88-514">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-514">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="14c88-515">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-515">Type</span></span>

*   <span data-ttu-id="14c88-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-517">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-517">Requirements</span></span>

|<span data-ttu-id="14c88-518">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-518">Requirement</span></span>| <span data-ttu-id="14c88-519">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-520">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-521">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-521">1.0</span></span>|
|[<span data-ttu-id="14c88-522">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-523">ReadItem</span></span>|
|[<span data-ttu-id="14c88-524">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-525">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-525">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="14c88-526">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-526">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-p132">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14c88-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="14c88-p133">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="14c88-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-531">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="14c88-531">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="14c88-532">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-532">Type</span></span>

*   [<span data-ttu-id="14c88-533">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="14c88-533">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="14c88-534">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-534">Requirements</span></span>

|<span data-ttu-id="14c88-535">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-535">Requirement</span></span>| <span data-ttu-id="14c88-536">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-537">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-538">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-538">1.0</span></span>|
|[<span data-ttu-id="14c88-539">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-540">ReadItem</span></span>|
|[<span data-ttu-id="14c88-541">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-542">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-542">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-543">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-543">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="14c88-544">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-544">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-545">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-545">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="14c88-p134">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="14c88-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14c88-548">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="14c88-548">Read mode</span></span>

<span data-ttu-id="14c88-549">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="14c88-549">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="14c88-550">Режим создания</span><span class="sxs-lookup"><span data-stu-id="14c88-550">Compose mode</span></span>

<span data-ttu-id="14c88-551">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="14c88-551">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="14c88-552">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="14c88-552">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="14c88-553">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="14c88-553">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="14c88-554">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-554">Type</span></span>

*   <span data-ttu-id="14c88-555">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-555">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-556">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-556">Requirements</span></span>

|<span data-ttu-id="14c88-557">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-557">Requirement</span></span>| <span data-ttu-id="14c88-558">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-558">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-559">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-559">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-560">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-560">1.0</span></span>|
|[<span data-ttu-id="14c88-561">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-561">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-562">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-562">ReadItem</span></span>|
|[<span data-ttu-id="14c88-563">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-563">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-564">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-564">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="14c88-565">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-565">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-566">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-566">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="14c88-567">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="14c88-567">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14c88-568">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="14c88-568">Read mode</span></span>

<span data-ttu-id="14c88-p135">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="14c88-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="14c88-571">Режим создания</span><span class="sxs-lookup"><span data-stu-id="14c88-571">Compose mode</span></span>

<span data-ttu-id="14c88-572">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="14c88-572">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="14c88-573">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-573">Type</span></span>

*   <span data-ttu-id="14c88-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-575">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-575">Requirements</span></span>

|<span data-ttu-id="14c88-576">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-576">Requirement</span></span>| <span data-ttu-id="14c88-577">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-577">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-578">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-578">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-579">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-579">1.0</span></span>|
|[<span data-ttu-id="14c88-580">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-580">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-581">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-581">ReadItem</span></span>|
|[<span data-ttu-id="14c88-582">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-582">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-583">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-583">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="14c88-584">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-584">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="14c88-585">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-585">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="14c88-586">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-586">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="14c88-587">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="14c88-587">Read mode</span></span>

<span data-ttu-id="14c88-588">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-588">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="14c88-589">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="14c88-589">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14c88-590">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="14c88-590">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="14c88-591">Режим создания</span><span class="sxs-lookup"><span data-stu-id="14c88-591">Compose mode</span></span>

<span data-ttu-id="14c88-592">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-592">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="14c88-593">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="14c88-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="14c88-594">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="14c88-594">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="14c88-595">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-595">Get 500 members maximum.</span></span>
- <span data-ttu-id="14c88-596">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="14c88-596">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="14c88-597">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-597">Type</span></span>

*   <span data-ttu-id="14c88-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-599">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-599">Requirements</span></span>

|<span data-ttu-id="14c88-600">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-600">Requirement</span></span>| <span data-ttu-id="14c88-601">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-602">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-603">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-603">1.0</span></span>|
|[<span data-ttu-id="14c88-604">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-604">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-605">ReadItem</span></span>|
|[<span data-ttu-id="14c88-606">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-606">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-607">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-607">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="14c88-608">Методы</span><span class="sxs-lookup"><span data-stu-id="14c88-608">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="14c88-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14c88-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="14c88-610">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="14c88-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="14c88-611">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="14c88-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="14c88-612">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="14c88-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14c88-613">Параметры</span><span class="sxs-lookup"><span data-stu-id="14c88-613">Parameters</span></span>

|<span data-ttu-id="14c88-614">Имя</span><span class="sxs-lookup"><span data-stu-id="14c88-614">Name</span></span>| <span data-ttu-id="14c88-615">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-615">Type</span></span>| <span data-ttu-id="14c88-616">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14c88-616">Attributes</span></span>| <span data-ttu-id="14c88-617">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="14c88-618">String</span><span class="sxs-lookup"><span data-stu-id="14c88-618">String</span></span>||<span data-ttu-id="14c88-p139">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="14c88-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="14c88-621">String</span><span class="sxs-lookup"><span data-stu-id="14c88-621">String</span></span>||<span data-ttu-id="14c88-p140">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="14c88-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="14c88-624">Object</span><span class="sxs-lookup"><span data-stu-id="14c88-624">Object</span></span>| <span data-ttu-id="14c88-625">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-625">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-626">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="14c88-626">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14c88-627">Object</span><span class="sxs-lookup"><span data-stu-id="14c88-627">Object</span></span>| <span data-ttu-id="14c88-628">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-628">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-629">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="14c88-629">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14c88-630">функция</span><span class="sxs-lookup"><span data-stu-id="14c88-630">function</span></span>| <span data-ttu-id="14c88-631">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-631">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-632">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14c88-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="14c88-633">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14c88-633">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="14c88-634">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="14c88-634">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14c88-635">Ошибки</span><span class="sxs-lookup"><span data-stu-id="14c88-635">Errors</span></span>

| <span data-ttu-id="14c88-636">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="14c88-636">Error code</span></span> | <span data-ttu-id="14c88-637">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-637">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="14c88-638">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="14c88-638">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="14c88-639">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="14c88-639">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="14c88-640">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="14c88-640">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14c88-641">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-641">Requirements</span></span>

|<span data-ttu-id="14c88-642">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-642">Requirement</span></span>| <span data-ttu-id="14c88-643">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-644">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-645">1.1</span><span class="sxs-lookup"><span data-stu-id="14c88-645">1.1</span></span>|
|[<span data-ttu-id="14c88-646">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-646">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-647">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14c88-647">ReadWriteItem</span></span>|
|[<span data-ttu-id="14c88-648">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-648">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-649">Создание</span><span class="sxs-lookup"><span data-stu-id="14c88-649">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-650">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-650">Example</span></span>

```js
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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="14c88-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14c88-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="14c88-652">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="14c88-652">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="14c88-p141">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="14c88-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="14c88-656">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="14c88-656">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="14c88-657">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="14c88-657">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14c88-658">Параметры</span><span class="sxs-lookup"><span data-stu-id="14c88-658">Parameters</span></span>

|<span data-ttu-id="14c88-659">Имя</span><span class="sxs-lookup"><span data-stu-id="14c88-659">Name</span></span>| <span data-ttu-id="14c88-660">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-660">Type</span></span>| <span data-ttu-id="14c88-661">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14c88-661">Attributes</span></span>| <span data-ttu-id="14c88-662">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-662">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="14c88-663">String</span><span class="sxs-lookup"><span data-stu-id="14c88-663">String</span></span>||<span data-ttu-id="14c88-p142">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="14c88-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="14c88-666">String</span><span class="sxs-lookup"><span data-stu-id="14c88-666">String</span></span>||<span data-ttu-id="14c88-667">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-667">The subject of the item to be attached.</span></span> <span data-ttu-id="14c88-668">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="14c88-668">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="14c88-669">Object</span><span class="sxs-lookup"><span data-stu-id="14c88-669">Object</span></span>| <span data-ttu-id="14c88-670">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-670">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-671">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="14c88-671">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14c88-672">Object</span><span class="sxs-lookup"><span data-stu-id="14c88-672">Object</span></span>| <span data-ttu-id="14c88-673">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-673">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-674">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="14c88-674">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14c88-675">функция</span><span class="sxs-lookup"><span data-stu-id="14c88-675">function</span></span>| <span data-ttu-id="14c88-676">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-676">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-677">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14c88-677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="14c88-678">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14c88-678">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="14c88-679">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="14c88-679">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14c88-680">Ошибки</span><span class="sxs-lookup"><span data-stu-id="14c88-680">Errors</span></span>

| <span data-ttu-id="14c88-681">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="14c88-681">Error code</span></span> | <span data-ttu-id="14c88-682">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-682">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="14c88-683">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="14c88-683">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14c88-684">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-684">Requirements</span></span>

|<span data-ttu-id="14c88-685">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-685">Requirement</span></span>| <span data-ttu-id="14c88-686">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-686">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-687">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-687">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-688">1.1</span><span class="sxs-lookup"><span data-stu-id="14c88-688">1.1</span></span>|
|[<span data-ttu-id="14c88-689">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-689">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-690">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14c88-690">ReadWriteItem</span></span>|
|[<span data-ttu-id="14c88-691">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-691">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-692">Создание</span><span class="sxs-lookup"><span data-stu-id="14c88-692">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-693">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-693">Example</span></span>

<span data-ttu-id="14c88-694">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="14c88-694">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="14c88-695">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="14c88-695">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="14c88-696">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-696">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-697">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14c88-697">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14c88-698">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="14c88-698">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="14c88-699">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="14c88-699">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-700">Возможность включать вложения в вызове `displayReplyAllForm` не поддерживается в наборе требований 1,1.</span><span class="sxs-lookup"><span data-stu-id="14c88-700">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="14c88-701">Добавлена поддержка вложений `displayReplyAllForm` в наборе требований 1,2 и выше.</span><span class="sxs-lookup"><span data-stu-id="14c88-701">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14c88-702">Параметры</span><span class="sxs-lookup"><span data-stu-id="14c88-702">Parameters</span></span>

|<span data-ttu-id="14c88-703">Имя</span><span class="sxs-lookup"><span data-stu-id="14c88-703">Name</span></span>| <span data-ttu-id="14c88-704">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-704">Type</span></span>| <span data-ttu-id="14c88-705">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-705">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="14c88-706">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="14c88-706">String &#124; Object</span></span>| |<span data-ttu-id="14c88-p145">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="14c88-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="14c88-709">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="14c88-709">**OR**</span></span><br/><span data-ttu-id="14c88-p146">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="14c88-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="14c88-712">String</span><span class="sxs-lookup"><span data-stu-id="14c88-712">String</span></span> | <span data-ttu-id="14c88-713">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-713">&lt;optional&gt;</span></span> | <span data-ttu-id="14c88-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="14c88-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="14c88-716">функция</span><span class="sxs-lookup"><span data-stu-id="14c88-716">function</span></span> | <span data-ttu-id="14c88-717">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-717">&lt;optional&gt;</span></span> | <span data-ttu-id="14c88-718">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14c88-718">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14c88-719">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-719">Requirements</span></span>

|<span data-ttu-id="14c88-720">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-720">Requirement</span></span>| <span data-ttu-id="14c88-721">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-721">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-722">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-722">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-723">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-723">1.0</span></span>|
|[<span data-ttu-id="14c88-724">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-724">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-725">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-725">ReadItem</span></span>|
|[<span data-ttu-id="14c88-726">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-726">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-727">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-727">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="14c88-728">Примеры</span><span class="sxs-lookup"><span data-stu-id="14c88-728">Examples</span></span>

<span data-ttu-id="14c88-729">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="14c88-729">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="14c88-730">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-730">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="14c88-731">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-731">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="14c88-732">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="14c88-732">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="14c88-733">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="14c88-733">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="14c88-734">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-734">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-735">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14c88-735">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14c88-736">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="14c88-736">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="14c88-737">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="14c88-737">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-738">Возможность включать вложения в вызове `displayReplyForm` не поддерживается в наборе требований 1,1.</span><span class="sxs-lookup"><span data-stu-id="14c88-738">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="14c88-739">Добавлена поддержка вложений `displayReplyForm` в наборе требований 1,2 и выше.</span><span class="sxs-lookup"><span data-stu-id="14c88-739">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14c88-740">Параметры</span><span class="sxs-lookup"><span data-stu-id="14c88-740">Parameters</span></span>

|<span data-ttu-id="14c88-741">Имя</span><span class="sxs-lookup"><span data-stu-id="14c88-741">Name</span></span>| <span data-ttu-id="14c88-742">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-742">Type</span></span>| <span data-ttu-id="14c88-743">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-743">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="14c88-744">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="14c88-744">String &#124; Object</span></span>| | <span data-ttu-id="14c88-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="14c88-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="14c88-747">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="14c88-747">**OR**</span></span><br/><span data-ttu-id="14c88-p150">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="14c88-p150">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="14c88-750">String</span><span class="sxs-lookup"><span data-stu-id="14c88-750">String</span></span> | <span data-ttu-id="14c88-751">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-751">&lt;optional&gt;</span></span> | <span data-ttu-id="14c88-p151">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="14c88-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="14c88-754">функция</span><span class="sxs-lookup"><span data-stu-id="14c88-754">function</span></span> | <span data-ttu-id="14c88-755">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-755">&lt;optional&gt;</span></span> | <span data-ttu-id="14c88-756">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14c88-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14c88-757">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-757">Requirements</span></span>

|<span data-ttu-id="14c88-758">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-758">Requirement</span></span>| <span data-ttu-id="14c88-759">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-760">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-761">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-761">1.0</span></span>|
|[<span data-ttu-id="14c88-762">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-763">ReadItem</span></span>|
|[<span data-ttu-id="14c88-764">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-765">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="14c88-766">Примеры</span><span class="sxs-lookup"><span data-stu-id="14c88-766">Examples</span></span>

<span data-ttu-id="14c88-767">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="14c88-767">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="14c88-768">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="14c88-769">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="14c88-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="14c88-770">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="14c88-770">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="14c88-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="14c88-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="14c88-772">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-772">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-773">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14c88-773">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-774">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-774">Requirements</span></span>

|<span data-ttu-id="14c88-775">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-775">Requirement</span></span>| <span data-ttu-id="14c88-776">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-776">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-777">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-777">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-778">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-778">1.0</span></span>|
|[<span data-ttu-id="14c88-779">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-779">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-780">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-780">ReadItem</span></span>|
|[<span data-ttu-id="14c88-781">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-781">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-782">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-782">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14c88-783">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="14c88-783">Returns:</span></span>

<span data-ttu-id="14c88-784">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="14c88-784">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="14c88-785">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-785">Example</span></span>

<span data-ttu-id="14c88-786">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-786">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="14c88-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="14c88-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="14c88-788">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-788">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-789">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14c88-789">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14c88-790">Параметры</span><span class="sxs-lookup"><span data-stu-id="14c88-790">Parameters</span></span>

|<span data-ttu-id="14c88-791">Имя</span><span class="sxs-lookup"><span data-stu-id="14c88-791">Name</span></span>| <span data-ttu-id="14c88-792">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-792">Type</span></span>| <span data-ttu-id="14c88-793">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-793">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="14c88-794">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="14c88-794">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="14c88-795">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="14c88-795">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14c88-796">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-796">Requirements</span></span>

|<span data-ttu-id="14c88-797">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-797">Requirement</span></span>| <span data-ttu-id="14c88-798">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-799">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-800">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-800">1.0</span></span>|
|[<span data-ttu-id="14c88-801">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-801">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-802">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="14c88-802">Restricted</span></span>|
|[<span data-ttu-id="14c88-803">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-803">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-804">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14c88-805">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="14c88-805">Returns:</span></span>

<span data-ttu-id="14c88-806">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="14c88-806">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="14c88-807">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="14c88-807">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="14c88-808">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="14c88-808">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="14c88-809">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="14c88-809">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="14c88-810">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="14c88-810">Value of `entityType`</span></span> | <span data-ttu-id="14c88-811">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="14c88-811">Type of objects in returned array</span></span> | <span data-ttu-id="14c88-812">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-812">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="14c88-813">String</span><span class="sxs-lookup"><span data-stu-id="14c88-813">String</span></span> | <span data-ttu-id="14c88-814">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="14c88-814">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="14c88-815">Contact</span><span class="sxs-lookup"><span data-stu-id="14c88-815">Contact</span></span> | <span data-ttu-id="14c88-816">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14c88-816">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="14c88-817">String</span><span class="sxs-lookup"><span data-stu-id="14c88-817">String</span></span> | <span data-ttu-id="14c88-818">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14c88-818">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="14c88-819">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="14c88-819">MeetingSuggestion</span></span> | <span data-ttu-id="14c88-820">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14c88-820">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="14c88-821">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="14c88-821">PhoneNumber</span></span> | <span data-ttu-id="14c88-822">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="14c88-822">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="14c88-823">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="14c88-823">TaskSuggestion</span></span> | <span data-ttu-id="14c88-824">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="14c88-824">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="14c88-825">String</span><span class="sxs-lookup"><span data-stu-id="14c88-825">String</span></span> | <span data-ttu-id="14c88-826">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="14c88-826">**Restricted**</span></span> |

<span data-ttu-id="14c88-827">Тип:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="14c88-827">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="14c88-828">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-828">Example</span></span>

<span data-ttu-id="14c88-829">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-829">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="14c88-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="14c88-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="14c88-831">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="14c88-831">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-832">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14c88-832">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14c88-833">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="14c88-833">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14c88-834">Параметры</span><span class="sxs-lookup"><span data-stu-id="14c88-834">Parameters</span></span>

|<span data-ttu-id="14c88-835">Имя</span><span class="sxs-lookup"><span data-stu-id="14c88-835">Name</span></span>| <span data-ttu-id="14c88-836">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-836">Type</span></span>| <span data-ttu-id="14c88-837">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-837">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="14c88-838">String</span><span class="sxs-lookup"><span data-stu-id="14c88-838">String</span></span>|<span data-ttu-id="14c88-839">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="14c88-839">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14c88-840">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-840">Requirements</span></span>

|<span data-ttu-id="14c88-841">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-841">Requirement</span></span>| <span data-ttu-id="14c88-842">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-843">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-844">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-844">1.0</span></span>|
|[<span data-ttu-id="14c88-845">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-846">ReadItem</span></span>|
|[<span data-ttu-id="14c88-847">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-848">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14c88-849">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="14c88-849">Returns:</span></span>

<span data-ttu-id="14c88-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="14c88-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="14c88-852">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="14c88-852">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="14c88-853">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="14c88-853">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="14c88-854">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="14c88-854">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-855">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14c88-855">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14c88-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="14c88-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="14c88-859">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="14c88-859">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="14c88-860">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="14c88-860">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="14c88-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="14c88-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14c88-863">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-863">Requirements</span></span>

|<span data-ttu-id="14c88-864">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-864">Requirement</span></span>| <span data-ttu-id="14c88-865">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-866">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-867">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-867">1.0</span></span>|
|[<span data-ttu-id="14c88-868">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-869">ReadItem</span></span>|
|[<span data-ttu-id="14c88-870">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-871">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14c88-872">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="14c88-872">Returns:</span></span>

<span data-ttu-id="14c88-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="14c88-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="14c88-875">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="14c88-875">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="14c88-876">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-876">Example</span></span>

<span data-ttu-id="14c88-877">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="14c88-877">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="14c88-878">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="14c88-878">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="14c88-879">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="14c88-879">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="14c88-880">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14c88-880">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="14c88-881">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="14c88-881">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="14c88-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="14c88-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14c88-884">Параметры</span><span class="sxs-lookup"><span data-stu-id="14c88-884">Parameters</span></span>

|<span data-ttu-id="14c88-885">Имя</span><span class="sxs-lookup"><span data-stu-id="14c88-885">Name</span></span>| <span data-ttu-id="14c88-886">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-886">Type</span></span>| <span data-ttu-id="14c88-887">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="14c88-888">String</span><span class="sxs-lookup"><span data-stu-id="14c88-888">String</span></span>|<span data-ttu-id="14c88-889">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="14c88-889">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14c88-890">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-890">Requirements</span></span>

|<span data-ttu-id="14c88-891">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-891">Requirement</span></span>| <span data-ttu-id="14c88-892">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-893">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-894">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-894">1.0</span></span>|
|[<span data-ttu-id="14c88-895">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-896">ReadItem</span></span>|
|[<span data-ttu-id="14c88-897">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-898">Чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14c88-899">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="14c88-899">Returns:</span></span>

<span data-ttu-id="14c88-900">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="14c88-900">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="14c88-901">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="14c88-901">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="14c88-902">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-902">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="14c88-903">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="14c88-903">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="14c88-904">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="14c88-904">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="14c88-p158">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="14c88-p158">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14c88-908">Параметры</span><span class="sxs-lookup"><span data-stu-id="14c88-908">Parameters</span></span>

|<span data-ttu-id="14c88-909">Имя</span><span class="sxs-lookup"><span data-stu-id="14c88-909">Name</span></span>| <span data-ttu-id="14c88-910">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-910">Type</span></span>| <span data-ttu-id="14c88-911">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14c88-911">Attributes</span></span>| <span data-ttu-id="14c88-912">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-912">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="14c88-913">function</span><span class="sxs-lookup"><span data-stu-id="14c88-913">function</span></span>||<span data-ttu-id="14c88-914">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14c88-914">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="14c88-915">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14c88-915">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="14c88-916">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="14c88-916">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="14c88-917">Объект</span><span class="sxs-lookup"><span data-stu-id="14c88-917">Object</span></span>| <span data-ttu-id="14c88-918">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-918">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-919">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="14c88-919">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="14c88-920">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="14c88-920">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14c88-921">Requirements</span><span class="sxs-lookup"><span data-stu-id="14c88-921">Requirements</span></span>

|<span data-ttu-id="14c88-922">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-922">Requirement</span></span>| <span data-ttu-id="14c88-923">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-923">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-924">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14c88-924">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-925">1.0</span><span class="sxs-lookup"><span data-stu-id="14c88-925">1.0</span></span>|
|[<span data-ttu-id="14c88-926">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-926">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-927">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14c88-927">ReadItem</span></span>|
|[<span data-ttu-id="14c88-928">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-928">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-929">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14c88-929">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-930">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-930">Example</span></span>

<span data-ttu-id="14c88-p161">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="14c88-p161">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="14c88-934">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14c88-934">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="14c88-935">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="14c88-935">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="14c88-936">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="14c88-936">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="14c88-937">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="14c88-937">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="14c88-938">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="14c88-938">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="14c88-939">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="14c88-939">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14c88-940">Параметры</span><span class="sxs-lookup"><span data-stu-id="14c88-940">Parameters</span></span>

|<span data-ttu-id="14c88-941">Имя</span><span class="sxs-lookup"><span data-stu-id="14c88-941">Name</span></span>| <span data-ttu-id="14c88-942">Тип</span><span class="sxs-lookup"><span data-stu-id="14c88-942">Type</span></span>| <span data-ttu-id="14c88-943">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14c88-943">Attributes</span></span>| <span data-ttu-id="14c88-944">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-944">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="14c88-945">String</span><span class="sxs-lookup"><span data-stu-id="14c88-945">String</span></span>||<span data-ttu-id="14c88-946">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="14c88-946">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="14c88-947">Object</span><span class="sxs-lookup"><span data-stu-id="14c88-947">Object</span></span>| <span data-ttu-id="14c88-948">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-948">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-949">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="14c88-949">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="14c88-950">Object</span><span class="sxs-lookup"><span data-stu-id="14c88-950">Object</span></span>| <span data-ttu-id="14c88-951">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-951">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-952">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="14c88-952">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="14c88-953">функция</span><span class="sxs-lookup"><span data-stu-id="14c88-953">function</span></span>| <span data-ttu-id="14c88-954">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14c88-954">&lt;optional&gt;</span></span>|<span data-ttu-id="14c88-955">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14c88-955">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="14c88-956">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="14c88-956">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14c88-957">Ошибки</span><span class="sxs-lookup"><span data-stu-id="14c88-957">Errors</span></span>

| <span data-ttu-id="14c88-958">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="14c88-958">Error code</span></span> | <span data-ttu-id="14c88-959">Описание</span><span class="sxs-lookup"><span data-stu-id="14c88-959">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="14c88-960">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="14c88-960">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14c88-961">Требования</span><span class="sxs-lookup"><span data-stu-id="14c88-961">Requirements</span></span>

|<span data-ttu-id="14c88-962">Требование</span><span class="sxs-lookup"><span data-stu-id="14c88-962">Requirement</span></span>| <span data-ttu-id="14c88-963">Значение</span><span class="sxs-lookup"><span data-stu-id="14c88-963">Value</span></span>|
|---|---|
|[<span data-ttu-id="14c88-964">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14c88-964">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14c88-965">1.1</span><span class="sxs-lookup"><span data-stu-id="14c88-965">1.1</span></span>|
|[<span data-ttu-id="14c88-966">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14c88-966">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14c88-967">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="14c88-967">ReadWriteItem</span></span>|
|[<span data-ttu-id="14c88-968">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14c88-968">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14c88-969">Создание</span><span class="sxs-lookup"><span data-stu-id="14c88-969">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="14c88-970">Пример</span><span class="sxs-lookup"><span data-stu-id="14c88-970">Example</span></span>

<span data-ttu-id="14c88-971">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="14c88-971">The following code removes an attachment with an identifier of '0'.</span></span>

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
