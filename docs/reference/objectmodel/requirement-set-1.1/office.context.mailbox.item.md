---
title: Office. Context. Mailbox. Item — набор требований 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 20d3aaecc5e0c62f86a46ae29010a6462446bf1d
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696444"
---
# <a name="item"></a><span data-ttu-id="8d1f2-102">item</span><span class="sxs-lookup"><span data-stu-id="8d1f2-102">item</span></span>

### <span data-ttu-id="8d1f2-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="8d1f2-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="8d1f2-107">Requirements</span></span>

|<span data-ttu-id="8d1f2-108">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-108">Requirement</span></span>| <span data-ttu-id="8d1f2-109">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-111">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-111">1.0</span></span>|
|[<span data-ttu-id="8d1f2-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="8d1f2-113">Restricted</span></span>|
|[<span data-ttu-id="8d1f2-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8d1f2-116">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="8d1f2-116">Members and methods</span></span>

| <span data-ttu-id="8d1f2-117">Элемент	</span><span class="sxs-lookup"><span data-stu-id="8d1f2-117">Member</span></span> | <span data-ttu-id="8d1f2-118">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8d1f2-119">attachments</span><span class="sxs-lookup"><span data-stu-id="8d1f2-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="8d1f2-120">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-120">Member</span></span> |
| [<span data-ttu-id="8d1f2-121">bcc</span><span class="sxs-lookup"><span data-stu-id="8d1f2-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="8d1f2-122">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-122">Member</span></span> |
| [<span data-ttu-id="8d1f2-123">body</span><span class="sxs-lookup"><span data-stu-id="8d1f2-123">body</span></span>](#body-body) | <span data-ttu-id="8d1f2-124">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-124">Member</span></span> |
| [<span data-ttu-id="8d1f2-125">cc</span><span class="sxs-lookup"><span data-stu-id="8d1f2-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8d1f2-126">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-126">Member</span></span> |
| [<span data-ttu-id="8d1f2-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="8d1f2-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8d1f2-128">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-128">Member</span></span> |
| [<span data-ttu-id="8d1f2-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8d1f2-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8d1f2-130">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-130">Member</span></span> |
| [<span data-ttu-id="8d1f2-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8d1f2-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8d1f2-132">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-132">Member</span></span> |
| [<span data-ttu-id="8d1f2-133">end</span><span class="sxs-lookup"><span data-stu-id="8d1f2-133">end</span></span>](#end-datetime) | <span data-ttu-id="8d1f2-134">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-134">Member</span></span> |
| [<span data-ttu-id="8d1f2-135">from</span><span class="sxs-lookup"><span data-stu-id="8d1f2-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="8d1f2-136">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-136">Member</span></span> |
| [<span data-ttu-id="8d1f2-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8d1f2-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8d1f2-138">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-138">Member</span></span> |
| [<span data-ttu-id="8d1f2-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="8d1f2-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8d1f2-140">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-140">Member</span></span> |
| [<span data-ttu-id="8d1f2-141">itemId</span><span class="sxs-lookup"><span data-stu-id="8d1f2-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8d1f2-142">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-142">Member</span></span> |
| [<span data-ttu-id="8d1f2-143">itemType</span><span class="sxs-lookup"><span data-stu-id="8d1f2-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="8d1f2-144">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-144">Member</span></span> |
| [<span data-ttu-id="8d1f2-145">location</span><span class="sxs-lookup"><span data-stu-id="8d1f2-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="8d1f2-146">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-146">Member</span></span> |
| [<span data-ttu-id="8d1f2-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8d1f2-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8d1f2-148">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-148">Member</span></span> |
| [<span data-ttu-id="8d1f2-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8d1f2-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8d1f2-150">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-150">Member</span></span> |
| [<span data-ttu-id="8d1f2-151">organizer</span><span class="sxs-lookup"><span data-stu-id="8d1f2-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="8d1f2-152">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-152">Member</span></span> |
| [<span data-ttu-id="8d1f2-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8d1f2-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8d1f2-154">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-154">Member</span></span> |
| [<span data-ttu-id="8d1f2-155">sender</span><span class="sxs-lookup"><span data-stu-id="8d1f2-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="8d1f2-156">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-156">Member</span></span> |
| [<span data-ttu-id="8d1f2-157">start</span><span class="sxs-lookup"><span data-stu-id="8d1f2-157">start</span></span>](#start-datetime) | <span data-ttu-id="8d1f2-158">Member</span><span class="sxs-lookup"><span data-stu-id="8d1f2-158">Member</span></span> |
| [<span data-ttu-id="8d1f2-159">subject</span><span class="sxs-lookup"><span data-stu-id="8d1f2-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="8d1f2-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="8d1f2-160">Member</span></span> |
| [<span data-ttu-id="8d1f2-161">to</span><span class="sxs-lookup"><span data-stu-id="8d1f2-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8d1f2-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="8d1f2-162">Member</span></span> |
| [<span data-ttu-id="8d1f2-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8d1f2-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8d1f2-164">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-164">Method</span></span> |
| [<span data-ttu-id="8d1f2-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8d1f2-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8d1f2-166">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-166">Method</span></span> |
| [<span data-ttu-id="8d1f2-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8d1f2-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="8d1f2-168">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-168">Method</span></span> |
| [<span data-ttu-id="8d1f2-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8d1f2-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="8d1f2-170">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-170">Method</span></span> |
| [<span data-ttu-id="8d1f2-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="8d1f2-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="8d1f2-172">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-172">Method</span></span> |
| [<span data-ttu-id="8d1f2-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8d1f2-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8d1f2-174">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-174">Method</span></span> |
| [<span data-ttu-id="8d1f2-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8d1f2-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8d1f2-176">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-176">Method</span></span> |
| [<span data-ttu-id="8d1f2-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8d1f2-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8d1f2-178">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-178">Method</span></span> |
| [<span data-ttu-id="8d1f2-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8d1f2-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8d1f2-180">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-180">Method</span></span> |
| [<span data-ttu-id="8d1f2-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8d1f2-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8d1f2-182">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-182">Method</span></span> |
| [<span data-ttu-id="8d1f2-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8d1f2-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8d1f2-184">Метод</span><span class="sxs-lookup"><span data-stu-id="8d1f2-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8d1f2-185">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-185">Example</span></span>

<span data-ttu-id="8d1f2-186">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8d1f2-187">Элементы</span><span class="sxs-lookup"><span data-stu-id="8d1f2-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="8d1f2-188">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="8d1f2-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="8d1f2-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-191">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8d1f2-192">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-193">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-193">Type</span></span>

*   <span data-ttu-id="8d1f2-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="8d1f2-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-195">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-195">Requirements</span></span>

|<span data-ttu-id="8d1f2-196">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-196">Requirement</span></span>| <span data-ttu-id="8d1f2-197">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-198">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-199">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-199">1.0</span></span>|
|[<span data-ttu-id="8d1f2-200">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-201">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-203">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-204">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-204">Example</span></span>

<span data-ttu-id="8d1f2-205">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8d1f2-206">СК: [получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-207">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8d1f2-208">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-208">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-209">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-209">Type</span></span>

*   [<span data-ttu-id="8d1f2-210">Получатели</span><span class="sxs-lookup"><span data-stu-id="8d1f2-210">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8d1f2-211">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-211">Requirements</span></span>

|<span data-ttu-id="8d1f2-212">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-212">Requirement</span></span>| <span data-ttu-id="8d1f2-213">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-214">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-215">1.1</span><span class="sxs-lookup"><span data-stu-id="8d1f2-215">1.1</span></span>|
|[<span data-ttu-id="8d1f2-216">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-217">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-218">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-219">Создание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-219">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-220">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-220">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="8d1f2-221">основной текст: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-221">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-222">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-222">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-223">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-223">Type</span></span>

*   [<span data-ttu-id="8d1f2-224">Body</span><span class="sxs-lookup"><span data-stu-id="8d1f2-224">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8d1f2-225">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-225">Requirements</span></span>

|<span data-ttu-id="8d1f2-226">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-226">Requirement</span></span>| <span data-ttu-id="8d1f2-227">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-228">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-229">1.1</span><span class="sxs-lookup"><span data-stu-id="8d1f2-229">1.1</span></span>|
|[<span data-ttu-id="8d1f2-230">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-231">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-232">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-233">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-234">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-234">Example</span></span>

<span data-ttu-id="8d1f2-235">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-235">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="8d1f2-236">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-236">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8d1f2-237">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-237">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-238">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-238">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8d1f2-239">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-239">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d1f2-240">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8d1f2-240">Read mode</span></span>

<span data-ttu-id="8d1f2-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d1f2-243">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8d1f2-243">Compose mode</span></span>

<span data-ttu-id="8d1f2-244">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-244">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8d1f2-245">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-245">Type</span></span>

*   <span data-ttu-id="8d1f2-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-247">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-247">Requirements</span></span>

|<span data-ttu-id="8d1f2-248">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-248">Requirement</span></span>| <span data-ttu-id="8d1f2-249">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-250">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-251">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-251">1.0</span></span>|
|[<span data-ttu-id="8d1f2-252">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-253">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-254">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-255">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="8d1f2-256">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="8d1f2-256">(nullable) conversationId: String</span></span>

<span data-ttu-id="8d1f2-257">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-257">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8d1f2-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8d1f2-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-262">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-262">Type</span></span>

*   <span data-ttu-id="8d1f2-263">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-263">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-264">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-264">Requirements</span></span>

|<span data-ttu-id="8d1f2-265">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-265">Requirement</span></span>| <span data-ttu-id="8d1f2-266">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-267">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-268">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-268">1.0</span></span>|
|[<span data-ttu-id="8d1f2-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-270">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-273">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-273">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="8d1f2-274">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="8d1f2-274">dateTimeCreated: Date</span></span>

<span data-ttu-id="8d1f2-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-277">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-277">Type</span></span>

*   <span data-ttu-id="8d1f2-278">Дата</span><span class="sxs-lookup"><span data-stu-id="8d1f2-278">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-279">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-279">Requirements</span></span>

|<span data-ttu-id="8d1f2-280">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-280">Requirement</span></span>| <span data-ttu-id="8d1f2-281">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-282">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-283">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-283">1.0</span></span>|
|[<span data-ttu-id="8d1f2-284">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-285">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-286">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-287">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-288">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-288">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="8d1f2-289">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="8d1f2-289">dateTimeModified: Date</span></span>

<span data-ttu-id="8d1f2-290">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-290">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="8d1f2-291">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-291">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-292">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-292">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-293">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-293">Type</span></span>

*   <span data-ttu-id="8d1f2-294">Дата</span><span class="sxs-lookup"><span data-stu-id="8d1f2-294">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-295">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-295">Requirements</span></span>

|<span data-ttu-id="8d1f2-296">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-296">Requirement</span></span>| <span data-ttu-id="8d1f2-297">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-298">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-299">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-299">1.0</span></span>|
|[<span data-ttu-id="8d1f2-300">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-301">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-302">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-303">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-303">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-304">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-304">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="8d1f2-305">конец: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="8d1f2-305">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-306">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-306">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8d1f2-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d1f2-309">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8d1f2-309">Read mode</span></span>

<span data-ttu-id="8d1f2-310">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-310">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="8d1f2-311">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8d1f2-311">Compose mode</span></span>

<span data-ttu-id="8d1f2-312">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-312">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8d1f2-313">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-313">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8d1f2-314">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-314">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8d1f2-315">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-315">Type</span></span>

*   <span data-ttu-id="8d1f2-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-317">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-317">Requirements</span></span>

|<span data-ttu-id="8d1f2-318">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-318">Requirement</span></span>| <span data-ttu-id="8d1f2-319">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-320">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-321">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-321">1.0</span></span>|
|[<span data-ttu-id="8d1f2-322">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-323">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-324">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-325">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-325">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="8d1f2-326">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-326">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="8d1f2-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-331">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-331">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-332">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-332">Type</span></span>

*   [<span data-ttu-id="8d1f2-333">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8d1f2-333">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8d1f2-334">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-334">Requirements</span></span>

|<span data-ttu-id="8d1f2-335">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-335">Requirement</span></span>| <span data-ttu-id="8d1f2-336">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-337">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-338">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-338">1.0</span></span>|
|[<span data-ttu-id="8d1f2-339">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-340">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-341">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-342">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-343">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-343">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="8d1f2-344">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="8d1f2-344">internetMessageId: String</span></span>

<span data-ttu-id="8d1f2-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-347">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-347">Type</span></span>

*   <span data-ttu-id="8d1f2-348">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-348">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-349">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-349">Requirements</span></span>

|<span data-ttu-id="8d1f2-350">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-350">Requirement</span></span>| <span data-ttu-id="8d1f2-351">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-351">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-352">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-353">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-353">1.0</span></span>|
|[<span data-ttu-id="8d1f2-354">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-355">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-356">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-357">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-357">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-358">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-358">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="8d1f2-359">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="8d1f2-359">itemClass: String</span></span>

<span data-ttu-id="8d1f2-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8d1f2-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="8d1f2-364">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-364">Type</span></span> | <span data-ttu-id="8d1f2-365">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-365">Description</span></span> | <span data-ttu-id="8d1f2-366">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="8d1f2-366">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="8d1f2-367">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="8d1f2-367">Appointment items</span></span> | <span data-ttu-id="8d1f2-368">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-368">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="8d1f2-369">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="8d1f2-369">Message items</span></span> | <span data-ttu-id="8d1f2-370">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-370">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="8d1f2-371">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-371">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-372">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-372">Type</span></span>

*   <span data-ttu-id="8d1f2-373">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-374">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-374">Requirements</span></span>

|<span data-ttu-id="8d1f2-375">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-375">Requirement</span></span>| <span data-ttu-id="8d1f2-376">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-377">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-378">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-378">1.0</span></span>|
|[<span data-ttu-id="8d1f2-379">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-380">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-381">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-382">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-383">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-383">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8d1f2-384">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="8d1f2-384">(nullable) itemId: String</span></span>

<span data-ttu-id="8d1f2-385">Получает идентификатор элемента веб-служб Exchange для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-385">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="8d1f2-386">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-386">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-387">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-387">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8d1f2-388">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-388">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8d1f2-389">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-389">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="8d1f2-390">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-390">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-391">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-391">Type</span></span>

*   <span data-ttu-id="8d1f2-392">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-392">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-393">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-393">Requirements</span></span>

|<span data-ttu-id="8d1f2-394">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-394">Requirement</span></span>| <span data-ttu-id="8d1f2-395">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-395">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-396">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-397">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-397">1.0</span></span>|
|[<span data-ttu-id="8d1f2-398">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-398">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-399">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-399">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-400">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-400">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-401">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-401">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-402">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-402">Example</span></span>

<span data-ttu-id="8d1f2-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="8d1f2-405">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-405">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-406">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-406">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8d1f2-407">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-407">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-408">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-408">Type</span></span>

*   [<span data-ttu-id="8d1f2-409">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8d1f2-409">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8d1f2-410">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-410">Requirements</span></span>

|<span data-ttu-id="8d1f2-411">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-411">Requirement</span></span>| <span data-ttu-id="8d1f2-412">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-413">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-414">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-414">1.0</span></span>|
|[<span data-ttu-id="8d1f2-415">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-416">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-417">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-418">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-418">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-419">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-419">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="8d1f2-420">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="8d1f2-420">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-421">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-421">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d1f2-422">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8d1f2-422">Read mode</span></span>

<span data-ttu-id="8d1f2-423">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-423">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="8d1f2-424">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8d1f2-424">Compose mode</span></span>

<span data-ttu-id="8d1f2-425">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-425">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8d1f2-426">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-426">Type</span></span>

*   <span data-ttu-id="8d1f2-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-428">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-428">Requirements</span></span>

|<span data-ttu-id="8d1f2-429">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-429">Requirement</span></span>| <span data-ttu-id="8d1f2-430">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-430">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-431">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-431">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-432">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-432">1.0</span></span>|
|[<span data-ttu-id="8d1f2-433">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-433">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-434">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-434">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-435">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-435">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-436">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-436">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8d1f2-437">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="8d1f2-437">normalizedSubject: String</span></span>

<span data-ttu-id="8d1f2-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8d1f2-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-442">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-442">Type</span></span>

*   <span data-ttu-id="8d1f2-443">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-443">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-444">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-444">Requirements</span></span>

|<span data-ttu-id="8d1f2-445">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-445">Requirement</span></span>| <span data-ttu-id="8d1f2-446">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-447">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-448">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-448">1.0</span></span>|
|[<span data-ttu-id="8d1f2-449">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-450">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-451">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-452">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-453">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-453">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8d1f2-454">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-454">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-455">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-455">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8d1f2-456">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-456">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d1f2-457">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8d1f2-457">Read mode</span></span>

<span data-ttu-id="8d1f2-458">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-458">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d1f2-459">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8d1f2-459">Compose mode</span></span>

<span data-ttu-id="8d1f2-460">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-460">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8d1f2-461">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-461">Type</span></span>

*   <span data-ttu-id="8d1f2-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-463">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-463">Requirements</span></span>

|<span data-ttu-id="8d1f2-464">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-464">Requirement</span></span>| <span data-ttu-id="8d1f2-465">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-466">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-467">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-467">1.0</span></span>|
|[<span data-ttu-id="8d1f2-468">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-469">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-470">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-471">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-471">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="8d1f2-472">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-472">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-475">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-475">Type</span></span>

*   [<span data-ttu-id="8d1f2-476">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8d1f2-476">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8d1f2-477">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-477">Requirements</span></span>

|<span data-ttu-id="8d1f2-478">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-478">Requirement</span></span>| <span data-ttu-id="8d1f2-479">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-480">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-481">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-481">1.0</span></span>|
|[<span data-ttu-id="8d1f2-482">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-482">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-483">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-484">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-484">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-485">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-485">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-486">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-486">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8d1f2-487">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-487">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-488">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-488">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8d1f2-489">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-489">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d1f2-490">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8d1f2-490">Read mode</span></span>

<span data-ttu-id="8d1f2-491">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-491">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d1f2-492">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8d1f2-492">Compose mode</span></span>

<span data-ttu-id="8d1f2-493">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-493">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="8d1f2-494">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-494">Type</span></span>

*   <span data-ttu-id="8d1f2-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-496">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-496">Requirements</span></span>

|<span data-ttu-id="8d1f2-497">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-497">Requirement</span></span>| <span data-ttu-id="8d1f2-498">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-499">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-500">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-500">1.0</span></span>|
|[<span data-ttu-id="8d1f2-501">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-502">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-503">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-504">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-504">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="8d1f2-505">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-505">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8d1f2-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-510">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-510">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8d1f2-511">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-511">Type</span></span>

*   [<span data-ttu-id="8d1f2-512">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8d1f2-512">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8d1f2-513">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-513">Requirements</span></span>

|<span data-ttu-id="8d1f2-514">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-514">Requirement</span></span>| <span data-ttu-id="8d1f2-515">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-515">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-516">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-516">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-517">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-517">1.0</span></span>|
|[<span data-ttu-id="8d1f2-518">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-518">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-519">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-519">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-520">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-520">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-521">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-521">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-522">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-522">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="8d1f2-523">Начало: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="8d1f2-523">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-524">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-524">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8d1f2-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d1f2-527">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8d1f2-527">Read mode</span></span>

<span data-ttu-id="8d1f2-528">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-528">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d1f2-529">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8d1f2-529">Compose mode</span></span>

<span data-ttu-id="8d1f2-530">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-530">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8d1f2-531">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-531">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8d1f2-532">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-532">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8d1f2-533">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-533">Type</span></span>

*   <span data-ttu-id="8d1f2-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-535">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-535">Requirements</span></span>

|<span data-ttu-id="8d1f2-536">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-536">Requirement</span></span>| <span data-ttu-id="8d1f2-537">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-538">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-539">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-539">1.0</span></span>|
|[<span data-ttu-id="8d1f2-540">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-541">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-542">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-543">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-543">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="8d1f2-544">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="8d1f2-544">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-545">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-545">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8d1f2-546">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-546">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d1f2-547">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8d1f2-547">Read mode</span></span>

<span data-ttu-id="8d1f2-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="8d1f2-550">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8d1f2-550">Compose mode</span></span>

<span data-ttu-id="8d1f2-551">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-551">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="8d1f2-552">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-552">Type</span></span>

*   <span data-ttu-id="8d1f2-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-554">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-554">Requirements</span></span>

|<span data-ttu-id="8d1f2-555">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-555">Requirement</span></span>| <span data-ttu-id="8d1f2-556">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-557">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-558">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-558">1.0</span></span>|
|[<span data-ttu-id="8d1f2-559">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-560">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-561">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-562">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-562">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8d1f2-563">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-563">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8d1f2-564">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-564">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8d1f2-565">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-565">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8d1f2-566">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8d1f2-566">Read mode</span></span>

<span data-ttu-id="8d1f2-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="8d1f2-569">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8d1f2-569">Compose mode</span></span>

<span data-ttu-id="8d1f2-570">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-570">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8d1f2-571">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-571">Type</span></span>

*   <span data-ttu-id="8d1f2-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-573">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-573">Requirements</span></span>

|<span data-ttu-id="8d1f2-574">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-574">Requirement</span></span>| <span data-ttu-id="8d1f2-575">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-575">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-576">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-576">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-577">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-577">1.0</span></span>|
|[<span data-ttu-id="8d1f2-578">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-578">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-579">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-579">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-580">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-580">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-581">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-581">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="8d1f2-582">Методы</span><span class="sxs-lookup"><span data-stu-id="8d1f2-582">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8d1f2-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8d1f2-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8d1f2-584">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-584">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8d1f2-585">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-585">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8d1f2-586">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-586">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d1f2-587">Параметры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-587">Parameters</span></span>

|<span data-ttu-id="8d1f2-588">Имя</span><span class="sxs-lookup"><span data-stu-id="8d1f2-588">Name</span></span>| <span data-ttu-id="8d1f2-589">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-589">Type</span></span>| <span data-ttu-id="8d1f2-590">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8d1f2-590">Attributes</span></span>| <span data-ttu-id="8d1f2-591">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-591">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="8d1f2-592">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-592">String</span></span>||<span data-ttu-id="8d1f2-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8d1f2-595">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-595">String</span></span>||<span data-ttu-id="8d1f2-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8d1f2-598">Объект</span><span class="sxs-lookup"><span data-stu-id="8d1f2-598">Object</span></span>| <span data-ttu-id="8d1f2-599">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-599">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-600">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-600">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8d1f2-601">Объект</span><span class="sxs-lookup"><span data-stu-id="8d1f2-601">Object</span></span>| <span data-ttu-id="8d1f2-602">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-602">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-603">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-603">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8d1f2-604">функция</span><span class="sxs-lookup"><span data-stu-id="8d1f2-604">function</span></span>| <span data-ttu-id="8d1f2-605">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-605">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-606">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8d1f2-607">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-607">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8d1f2-608">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-608">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8d1f2-609">Ошибки</span><span class="sxs-lookup"><span data-stu-id="8d1f2-609">Errors</span></span>

| <span data-ttu-id="8d1f2-610">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="8d1f2-610">Error code</span></span> | <span data-ttu-id="8d1f2-611">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-611">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="8d1f2-612">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-612">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="8d1f2-613">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-613">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8d1f2-614">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-614">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d1f2-615">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-615">Requirements</span></span>

|<span data-ttu-id="8d1f2-616">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-616">Requirement</span></span>| <span data-ttu-id="8d1f2-617">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-618">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-619">1.1</span><span class="sxs-lookup"><span data-stu-id="8d1f2-619">1.1</span></span>|
|[<span data-ttu-id="8d1f2-620">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-621">ReadWriteItem</span></span>|
|[<span data-ttu-id="8d1f2-622">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-623">Создание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-624">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-624">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8d1f2-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8d1f2-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8d1f2-626">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-626">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8d1f2-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8d1f2-630">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-630">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8d1f2-631">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-631">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d1f2-632">Параметры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-632">Parameters</span></span>

|<span data-ttu-id="8d1f2-633">Имя</span><span class="sxs-lookup"><span data-stu-id="8d1f2-633">Name</span></span>| <span data-ttu-id="8d1f2-634">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-634">Type</span></span>| <span data-ttu-id="8d1f2-635">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8d1f2-635">Attributes</span></span>| <span data-ttu-id="8d1f2-636">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-636">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="8d1f2-637">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-637">String</span></span>||<span data-ttu-id="8d1f2-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8d1f2-640">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-640">String</span></span>||<span data-ttu-id="8d1f2-641">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-641">The subject of the item to be attached.</span></span> <span data-ttu-id="8d1f2-642">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-642">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8d1f2-643">Object</span><span class="sxs-lookup"><span data-stu-id="8d1f2-643">Object</span></span>| <span data-ttu-id="8d1f2-644">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-644">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-645">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-645">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8d1f2-646">Объект</span><span class="sxs-lookup"><span data-stu-id="8d1f2-646">Object</span></span>| <span data-ttu-id="8d1f2-647">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-647">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-648">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-648">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8d1f2-649">функция</span><span class="sxs-lookup"><span data-stu-id="8d1f2-649">function</span></span>| <span data-ttu-id="8d1f2-650">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-650">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-651">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-651">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8d1f2-652">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-652">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8d1f2-653">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-653">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8d1f2-654">Ошибки</span><span class="sxs-lookup"><span data-stu-id="8d1f2-654">Errors</span></span>

| <span data-ttu-id="8d1f2-655">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="8d1f2-655">Error code</span></span> | <span data-ttu-id="8d1f2-656">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-656">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8d1f2-657">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-657">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d1f2-658">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-658">Requirements</span></span>

|<span data-ttu-id="8d1f2-659">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-659">Requirement</span></span>| <span data-ttu-id="8d1f2-660">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-661">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-662">1.1</span><span class="sxs-lookup"><span data-stu-id="8d1f2-662">1.1</span></span>|
|[<span data-ttu-id="8d1f2-663">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-664">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-664">ReadWriteItem</span></span>|
|[<span data-ttu-id="8d1f2-665">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-666">Создание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-666">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-667">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-667">Example</span></span>

<span data-ttu-id="8d1f2-668">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-668">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="8d1f2-669">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8d1f2-669">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="8d1f2-670">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-670">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-671">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-671">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d1f2-672">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-672">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8d1f2-673">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-673">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-674">Возможность включать вложения в вызове `displayReplyAllForm` не поддерживается в наборе требований 1,1.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-674">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="8d1f2-675">Добавлена поддержка вложений `displayReplyAllForm` в наборе требований 1,2 и выше.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-675">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d1f2-676">Параметры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-676">Parameters</span></span>

|<span data-ttu-id="8d1f2-677">Имя</span><span class="sxs-lookup"><span data-stu-id="8d1f2-677">Name</span></span>| <span data-ttu-id="8d1f2-678">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-678">Type</span></span>| <span data-ttu-id="8d1f2-679">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-679">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8d1f2-680">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8d1f2-680">String &#124; Object</span></span>| |<span data-ttu-id="8d1f2-p138">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8d1f2-683">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="8d1f2-683">**OR**</span></span><br/><span data-ttu-id="8d1f2-p139">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8d1f2-686">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-686">String</span></span> | <span data-ttu-id="8d1f2-687">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-687">&lt;optional&gt;</span></span> | <span data-ttu-id="8d1f2-p140">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="8d1f2-690">функция</span><span class="sxs-lookup"><span data-stu-id="8d1f2-690">function</span></span> | <span data-ttu-id="8d1f2-691">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-691">&lt;optional&gt;</span></span> | <span data-ttu-id="8d1f2-692">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-692">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d1f2-693">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-693">Requirements</span></span>

|<span data-ttu-id="8d1f2-694">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-694">Requirement</span></span>| <span data-ttu-id="8d1f2-695">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-695">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-696">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-696">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-697">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-697">1.0</span></span>|
|[<span data-ttu-id="8d1f2-698">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-698">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-699">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-699">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-700">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-700">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-701">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-701">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8d1f2-702">Примеры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-702">Examples</span></span>

<span data-ttu-id="8d1f2-703">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-703">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8d1f2-704">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-704">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8d1f2-705">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-705">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8d1f2-706">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-706">Reply with a body and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="8d1f2-707">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8d1f2-707">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="8d1f2-708">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-708">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-709">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-709">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d1f2-710">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-710">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8d1f2-711">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-711">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-712">Возможность включать вложения в вызове `displayReplyForm` не поддерживается в наборе требований 1,1.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-712">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="8d1f2-713">Добавлена поддержка вложений `displayReplyForm` в наборе требований 1,2 и выше.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-713">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d1f2-714">Параметры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-714">Parameters</span></span>

|<span data-ttu-id="8d1f2-715">Имя</span><span class="sxs-lookup"><span data-stu-id="8d1f2-715">Name</span></span>| <span data-ttu-id="8d1f2-716">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-716">Type</span></span>| <span data-ttu-id="8d1f2-717">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-717">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8d1f2-718">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8d1f2-718">String &#124; Object</span></span>| | <span data-ttu-id="8d1f2-p142">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8d1f2-721">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="8d1f2-721">**OR**</span></span><br/><span data-ttu-id="8d1f2-p143">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8d1f2-724">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-724">String</span></span> | <span data-ttu-id="8d1f2-725">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-725">&lt;optional&gt;</span></span> | <span data-ttu-id="8d1f2-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="8d1f2-728">функция</span><span class="sxs-lookup"><span data-stu-id="8d1f2-728">function</span></span> | <span data-ttu-id="8d1f2-729">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-729">&lt;optional&gt;</span></span> | <span data-ttu-id="8d1f2-730">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-730">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d1f2-731">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-731">Requirements</span></span>

|<span data-ttu-id="8d1f2-732">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-732">Requirement</span></span>| <span data-ttu-id="8d1f2-733">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-734">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-735">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-735">1.0</span></span>|
|[<span data-ttu-id="8d1f2-736">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-737">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-738">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-739">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-739">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8d1f2-740">Примеры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-740">Examples</span></span>

<span data-ttu-id="8d1f2-741">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-741">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8d1f2-742">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-742">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8d1f2-743">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-743">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8d1f2-744">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-744">Reply with a body and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="8d1f2-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="8d1f2-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="8d1f2-746">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-746">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-747">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-747">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-748">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-748">Requirements</span></span>

|<span data-ttu-id="8d1f2-749">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-749">Requirement</span></span>| <span data-ttu-id="8d1f2-750">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-750">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-751">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-752">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-752">1.0</span></span>|
|[<span data-ttu-id="8d1f2-753">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-754">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-754">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-755">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-756">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-756">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d1f2-757">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8d1f2-757">Returns:</span></span>

<span data-ttu-id="8d1f2-758">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8d1f2-758">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="8d1f2-759">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-759">Example</span></span>

<span data-ttu-id="8d1f2-760">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-760">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="8d1f2-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="8d1f2-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="8d1f2-762">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-762">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-763">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-763">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d1f2-764">Параметры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-764">Parameters</span></span>

|<span data-ttu-id="8d1f2-765">Имя</span><span class="sxs-lookup"><span data-stu-id="8d1f2-765">Name</span></span>| <span data-ttu-id="8d1f2-766">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-766">Type</span></span>| <span data-ttu-id="8d1f2-767">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-767">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="8d1f2-768">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8d1f2-768">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="8d1f2-769">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-769">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d1f2-770">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-770">Requirements</span></span>

|<span data-ttu-id="8d1f2-771">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-771">Requirement</span></span>| <span data-ttu-id="8d1f2-772">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-773">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-774">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-774">1.0</span></span>|
|[<span data-ttu-id="8d1f2-775">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-775">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-776">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="8d1f2-776">Restricted</span></span>|
|[<span data-ttu-id="8d1f2-777">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-777">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-778">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d1f2-779">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8d1f2-779">Returns:</span></span>

<span data-ttu-id="8d1f2-780">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-780">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8d1f2-781">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-781">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8d1f2-782">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-782">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8d1f2-783">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-783">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="8d1f2-784">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="8d1f2-784">Value of `entityType`</span></span> | <span data-ttu-id="8d1f2-785">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="8d1f2-785">Type of objects in returned array</span></span> | <span data-ttu-id="8d1f2-786">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-786">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="8d1f2-787">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-787">String</span></span> | <span data-ttu-id="8d1f2-788">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="8d1f2-788">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="8d1f2-789">Contact</span><span class="sxs-lookup"><span data-stu-id="8d1f2-789">Contact</span></span> | <span data-ttu-id="8d1f2-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8d1f2-790">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="8d1f2-791">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-791">String</span></span> | <span data-ttu-id="8d1f2-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8d1f2-792">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="8d1f2-793">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8d1f2-793">MeetingSuggestion</span></span> | <span data-ttu-id="8d1f2-794">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8d1f2-794">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="8d1f2-795">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8d1f2-795">PhoneNumber</span></span> | <span data-ttu-id="8d1f2-796">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="8d1f2-796">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="8d1f2-797">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8d1f2-797">TaskSuggestion</span></span> | <span data-ttu-id="8d1f2-798">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8d1f2-798">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="8d1f2-799">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-799">String</span></span> | <span data-ttu-id="8d1f2-800">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="8d1f2-800">**Restricted**</span></span> |

<span data-ttu-id="8d1f2-801">Тип:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="8d1f2-801">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="8d1f2-802">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-802">Example</span></span>

<span data-ttu-id="8d1f2-803">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-803">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="8d1f2-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="8d1f2-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="8d1f2-805">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-805">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-806">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-806">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d1f2-807">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-807">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d1f2-808">Параметры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-808">Parameters</span></span>

|<span data-ttu-id="8d1f2-809">Имя</span><span class="sxs-lookup"><span data-stu-id="8d1f2-809">Name</span></span>| <span data-ttu-id="8d1f2-810">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-810">Type</span></span>| <span data-ttu-id="8d1f2-811">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-811">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8d1f2-812">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-812">String</span></span>|<span data-ttu-id="8d1f2-813">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-813">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d1f2-814">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-814">Requirements</span></span>

|<span data-ttu-id="8d1f2-815">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-815">Requirement</span></span>| <span data-ttu-id="8d1f2-816">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-817">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-818">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-818">1.0</span></span>|
|[<span data-ttu-id="8d1f2-819">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-820">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-821">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-822">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d1f2-823">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8d1f2-823">Returns:</span></span>

<span data-ttu-id="8d1f2-p146">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="8d1f2-826">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="8d1f2-826">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="8d1f2-827">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8d1f2-827">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8d1f2-828">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-828">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-829">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-829">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d1f2-p147">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8d1f2-833">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-833">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8d1f2-834">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-834">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="8d1f2-p148">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d1f2-837">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-837">Requirements</span></span>

|<span data-ttu-id="8d1f2-838">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-838">Requirement</span></span>| <span data-ttu-id="8d1f2-839">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-840">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-841">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-841">1.0</span></span>|
|[<span data-ttu-id="8d1f2-842">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-842">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-843">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-844">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-844">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-845">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d1f2-846">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8d1f2-846">Returns:</span></span>

<span data-ttu-id="8d1f2-p149">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="8d1f2-849">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="8d1f2-849">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="8d1f2-850">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-850">Example</span></span>

<span data-ttu-id="8d1f2-851">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-851">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8d1f2-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8d1f2-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8d1f2-853">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-853">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8d1f2-854">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-854">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8d1f2-855">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-855">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8d1f2-p150">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d1f2-858">Параметры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-858">Parameters</span></span>

|<span data-ttu-id="8d1f2-859">Имя</span><span class="sxs-lookup"><span data-stu-id="8d1f2-859">Name</span></span>| <span data-ttu-id="8d1f2-860">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-860">Type</span></span>| <span data-ttu-id="8d1f2-861">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-861">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8d1f2-862">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-862">String</span></span>|<span data-ttu-id="8d1f2-863">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-863">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d1f2-864">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-864">Requirements</span></span>

|<span data-ttu-id="8d1f2-865">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-865">Requirement</span></span>| <span data-ttu-id="8d1f2-866">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-866">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-867">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-867">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-868">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-868">1.0</span></span>|
|[<span data-ttu-id="8d1f2-869">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-869">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-870">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-870">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-871">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-871">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-872">Чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-872">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8d1f2-873">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8d1f2-873">Returns:</span></span>

<span data-ttu-id="8d1f2-874">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-874">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="8d1f2-875">Тип: Array. < String ></span><span class="sxs-lookup"><span data-stu-id="8d1f2-875">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="8d1f2-876">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-876">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8d1f2-877">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8d1f2-877">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8d1f2-878">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-878">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8d1f2-p151">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d1f2-882">Параметры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-882">Parameters</span></span>

|<span data-ttu-id="8d1f2-883">Имя</span><span class="sxs-lookup"><span data-stu-id="8d1f2-883">Name</span></span>| <span data-ttu-id="8d1f2-884">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-884">Type</span></span>| <span data-ttu-id="8d1f2-885">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8d1f2-885">Attributes</span></span>| <span data-ttu-id="8d1f2-886">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-886">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8d1f2-887">function</span><span class="sxs-lookup"><span data-stu-id="8d1f2-887">function</span></span>||<span data-ttu-id="8d1f2-888">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-888">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8d1f2-889">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-889">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8d1f2-890">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-890">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="8d1f2-891">Объект</span><span class="sxs-lookup"><span data-stu-id="8d1f2-891">Object</span></span>| <span data-ttu-id="8d1f2-892">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-892">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-893">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-893">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8d1f2-894">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-894">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d1f2-895">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-895">Requirements</span></span>

|<span data-ttu-id="8d1f2-896">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-896">Requirement</span></span>| <span data-ttu-id="8d1f2-897">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-897">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-898">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8d1f2-898">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-899">1.0</span><span class="sxs-lookup"><span data-stu-id="8d1f2-899">1.0</span></span>|
|[<span data-ttu-id="8d1f2-900">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-900">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-901">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-901">ReadItem</span></span>|
|[<span data-ttu-id="8d1f2-902">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-902">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-903">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-903">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-904">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-904">Example</span></span>

<span data-ttu-id="8d1f2-p154">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8d1f2-908">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8d1f2-908">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8d1f2-909">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-909">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8d1f2-910">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-910">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="8d1f2-911">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-911">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="8d1f2-912">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-912">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="8d1f2-913">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-913">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8d1f2-914">Параметры</span><span class="sxs-lookup"><span data-stu-id="8d1f2-914">Parameters</span></span>

|<span data-ttu-id="8d1f2-915">Имя</span><span class="sxs-lookup"><span data-stu-id="8d1f2-915">Name</span></span>| <span data-ttu-id="8d1f2-916">Тип</span><span class="sxs-lookup"><span data-stu-id="8d1f2-916">Type</span></span>| <span data-ttu-id="8d1f2-917">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8d1f2-917">Attributes</span></span>| <span data-ttu-id="8d1f2-918">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-918">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="8d1f2-919">String</span><span class="sxs-lookup"><span data-stu-id="8d1f2-919">String</span></span>||<span data-ttu-id="8d1f2-920">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-920">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="8d1f2-921">Object</span><span class="sxs-lookup"><span data-stu-id="8d1f2-921">Object</span></span>| <span data-ttu-id="8d1f2-922">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-922">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-923">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-923">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8d1f2-924">Объект</span><span class="sxs-lookup"><span data-stu-id="8d1f2-924">Object</span></span>| <span data-ttu-id="8d1f2-925">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-925">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-926">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-926">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8d1f2-927">функция</span><span class="sxs-lookup"><span data-stu-id="8d1f2-927">function</span></span>| <span data-ttu-id="8d1f2-928">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8d1f2-928">&lt;optional&gt;</span></span>|<span data-ttu-id="8d1f2-929">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8d1f2-929">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8d1f2-930">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-930">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8d1f2-931">Ошибки</span><span class="sxs-lookup"><span data-stu-id="8d1f2-931">Errors</span></span>

| <span data-ttu-id="8d1f2-932">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="8d1f2-932">Error code</span></span> | <span data-ttu-id="8d1f2-933">Описание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-933">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="8d1f2-934">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="8d1f2-934">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d1f2-935">Требования</span><span class="sxs-lookup"><span data-stu-id="8d1f2-935">Requirements</span></span>

|<span data-ttu-id="8d1f2-936">Требование</span><span class="sxs-lookup"><span data-stu-id="8d1f2-936">Requirement</span></span>| <span data-ttu-id="8d1f2-937">Значение</span><span class="sxs-lookup"><span data-stu-id="8d1f2-937">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d1f2-938">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8d1f2-938">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8d1f2-939">1.1</span><span class="sxs-lookup"><span data-stu-id="8d1f2-939">1.1</span></span>|
|[<span data-ttu-id="8d1f2-940">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8d1f2-940">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8d1f2-941">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8d1f2-941">ReadWriteItem</span></span>|
|[<span data-ttu-id="8d1f2-942">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8d1f2-942">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d1f2-943">Создание</span><span class="sxs-lookup"><span data-stu-id="8d1f2-943">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8d1f2-944">Пример</span><span class="sxs-lookup"><span data-stu-id="8d1f2-944">Example</span></span>

<span data-ttu-id="8d1f2-945">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="8d1f2-945">The following code removes an attachment with an identifier of '0'.</span></span>

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
