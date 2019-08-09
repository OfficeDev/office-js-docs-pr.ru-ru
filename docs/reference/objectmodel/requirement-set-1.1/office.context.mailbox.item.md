---
title: Office. Context. Mailbox. Item — набор требований 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: d3242f2bdabf464c262fdb8e6efd8695dc7ee330
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268504"
---
# <a name="item"></a><span data-ttu-id="55c70-102">item</span><span class="sxs-lookup"><span data-stu-id="55c70-102">item</span></span>

### <span data-ttu-id="55c70-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="55c70-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="55c70-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="55c70-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="55c70-107">Requirements</span></span>

|<span data-ttu-id="55c70-108">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-108">Requirement</span></span>| <span data-ttu-id="55c70-109">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-111">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-111">1.0</span></span>|
|[<span data-ttu-id="55c70-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="55c70-113">Restricted</span></span>|
|[<span data-ttu-id="55c70-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="55c70-116">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="55c70-116">Members and methods</span></span>

| <span data-ttu-id="55c70-117">Элемент	</span><span class="sxs-lookup"><span data-stu-id="55c70-117">Member</span></span> | <span data-ttu-id="55c70-118">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="55c70-119">attachments</span><span class="sxs-lookup"><span data-stu-id="55c70-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="55c70-120">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-120">Member</span></span> |
| [<span data-ttu-id="55c70-121">bcc</span><span class="sxs-lookup"><span data-stu-id="55c70-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="55c70-122">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-122">Member</span></span> |
| [<span data-ttu-id="55c70-123">body</span><span class="sxs-lookup"><span data-stu-id="55c70-123">body</span></span>](#body-body) | <span data-ttu-id="55c70-124">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-124">Member</span></span> |
| [<span data-ttu-id="55c70-125">cc</span><span class="sxs-lookup"><span data-stu-id="55c70-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="55c70-126">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-126">Member</span></span> |
| [<span data-ttu-id="55c70-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="55c70-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="55c70-128">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-128">Member</span></span> |
| [<span data-ttu-id="55c70-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="55c70-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="55c70-130">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-130">Member</span></span> |
| [<span data-ttu-id="55c70-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="55c70-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="55c70-132">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-132">Member</span></span> |
| [<span data-ttu-id="55c70-133">end</span><span class="sxs-lookup"><span data-stu-id="55c70-133">end</span></span>](#end-datetime) | <span data-ttu-id="55c70-134">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-134">Member</span></span> |
| [<span data-ttu-id="55c70-135">from</span><span class="sxs-lookup"><span data-stu-id="55c70-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="55c70-136">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-136">Member</span></span> |
| [<span data-ttu-id="55c70-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="55c70-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="55c70-138">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-138">Member</span></span> |
| [<span data-ttu-id="55c70-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="55c70-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="55c70-140">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-140">Member</span></span> |
| [<span data-ttu-id="55c70-141">itemId</span><span class="sxs-lookup"><span data-stu-id="55c70-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="55c70-142">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-142">Member</span></span> |
| [<span data-ttu-id="55c70-143">itemType</span><span class="sxs-lookup"><span data-stu-id="55c70-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="55c70-144">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-144">Member</span></span> |
| [<span data-ttu-id="55c70-145">location</span><span class="sxs-lookup"><span data-stu-id="55c70-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="55c70-146">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-146">Member</span></span> |
| [<span data-ttu-id="55c70-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="55c70-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="55c70-148">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-148">Member</span></span> |
| [<span data-ttu-id="55c70-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="55c70-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="55c70-150">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-150">Member</span></span> |
| [<span data-ttu-id="55c70-151">organizer</span><span class="sxs-lookup"><span data-stu-id="55c70-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="55c70-152">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-152">Member</span></span> |
| [<span data-ttu-id="55c70-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="55c70-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="55c70-154">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-154">Member</span></span> |
| [<span data-ttu-id="55c70-155">sender</span><span class="sxs-lookup"><span data-stu-id="55c70-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="55c70-156">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-156">Member</span></span> |
| [<span data-ttu-id="55c70-157">start</span><span class="sxs-lookup"><span data-stu-id="55c70-157">start</span></span>](#start-datetime) | <span data-ttu-id="55c70-158">Member</span><span class="sxs-lookup"><span data-stu-id="55c70-158">Member</span></span> |
| [<span data-ttu-id="55c70-159">subject</span><span class="sxs-lookup"><span data-stu-id="55c70-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="55c70-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="55c70-160">Member</span></span> |
| [<span data-ttu-id="55c70-161">to</span><span class="sxs-lookup"><span data-stu-id="55c70-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="55c70-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="55c70-162">Member</span></span> |
| [<span data-ttu-id="55c70-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="55c70-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="55c70-164">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-164">Method</span></span> |
| [<span data-ttu-id="55c70-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="55c70-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="55c70-166">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-166">Method</span></span> |
| [<span data-ttu-id="55c70-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="55c70-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="55c70-168">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-168">Method</span></span> |
| [<span data-ttu-id="55c70-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="55c70-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="55c70-170">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-170">Method</span></span> |
| [<span data-ttu-id="55c70-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="55c70-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="55c70-172">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-172">Method</span></span> |
| [<span data-ttu-id="55c70-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="55c70-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="55c70-174">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-174">Method</span></span> |
| [<span data-ttu-id="55c70-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="55c70-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="55c70-176">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-176">Method</span></span> |
| [<span data-ttu-id="55c70-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="55c70-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="55c70-178">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-178">Method</span></span> |
| [<span data-ttu-id="55c70-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="55c70-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="55c70-180">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-180">Method</span></span> |
| [<span data-ttu-id="55c70-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="55c70-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="55c70-182">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-182">Method</span></span> |
| [<span data-ttu-id="55c70-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="55c70-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="55c70-184">Метод</span><span class="sxs-lookup"><span data-stu-id="55c70-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="55c70-185">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-185">Example</span></span>

<span data-ttu-id="55c70-186">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="55c70-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="55c70-187">Элементы</span><span class="sxs-lookup"><span data-stu-id="55c70-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="55c70-188">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="55c70-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="55c70-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-191">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="55c70-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="55c70-192">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="55c70-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-193">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-193">Type</span></span>

*   <span data-ttu-id="55c70-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="55c70-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-195">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-195">Requirements</span></span>

|<span data-ttu-id="55c70-196">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-196">Requirement</span></span>| <span data-ttu-id="55c70-197">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-198">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-199">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-199">1.0</span></span>|
|[<span data-ttu-id="55c70-200">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-201">ReadItem</span></span>|
|[<span data-ttu-id="55c70-202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-203">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-204">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-204">Example</span></span>

<span data-ttu-id="55c70-205">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="55c70-206">СК: [получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-207">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="55c70-208">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="55c70-208">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-209">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-209">Type</span></span>

*   [<span data-ttu-id="55c70-210">Получатели</span><span class="sxs-lookup"><span data-stu-id="55c70-210">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="55c70-211">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-211">Requirements</span></span>

|<span data-ttu-id="55c70-212">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-212">Requirement</span></span>| <span data-ttu-id="55c70-213">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-214">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-215">1.1</span><span class="sxs-lookup"><span data-stu-id="55c70-215">1.1</span></span>|
|[<span data-ttu-id="55c70-216">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-217">ReadItem</span></span>|
|[<span data-ttu-id="55c70-218">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-219">Создание</span><span class="sxs-lookup"><span data-stu-id="55c70-219">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-220">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-220">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="55c70-221">основной текст: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-221">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-222">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-222">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-223">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-223">Type</span></span>

*   [<span data-ttu-id="55c70-224">Body</span><span class="sxs-lookup"><span data-stu-id="55c70-224">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="55c70-225">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-225">Requirements</span></span>

|<span data-ttu-id="55c70-226">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-226">Requirement</span></span>| <span data-ttu-id="55c70-227">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-228">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-229">1.1</span><span class="sxs-lookup"><span data-stu-id="55c70-229">1.1</span></span>|
|[<span data-ttu-id="55c70-230">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-231">ReadItem</span></span>|
|[<span data-ttu-id="55c70-232">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-233">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-234">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-234">Example</span></span>

<span data-ttu-id="55c70-235">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="55c70-235">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="55c70-236">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="55c70-236">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="55c70-237">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-237">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-238">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-238">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="55c70-239">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-239">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="55c70-240">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="55c70-240">Read mode</span></span>

<span data-ttu-id="55c70-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="55c70-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="55c70-243">Режим создания</span><span class="sxs-lookup"><span data-stu-id="55c70-243">Compose mode</span></span>

<span data-ttu-id="55c70-244">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-244">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="55c70-245">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-245">Type</span></span>

*   <span data-ttu-id="55c70-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-247">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-247">Requirements</span></span>

|<span data-ttu-id="55c70-248">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-248">Requirement</span></span>| <span data-ttu-id="55c70-249">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-250">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-251">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-251">1.0</span></span>|
|[<span data-ttu-id="55c70-252">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-253">ReadItem</span></span>|
|[<span data-ttu-id="55c70-254">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-255">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-255">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="55c70-256">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="55c70-256">(nullable) conversationId: String</span></span>

<span data-ttu-id="55c70-257">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="55c70-257">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="55c70-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="55c70-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="55c70-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="55c70-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-262">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-262">Type</span></span>

*   <span data-ttu-id="55c70-263">String</span><span class="sxs-lookup"><span data-stu-id="55c70-263">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-264">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-264">Requirements</span></span>

|<span data-ttu-id="55c70-265">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-265">Requirement</span></span>| <span data-ttu-id="55c70-266">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-267">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-268">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-268">1.0</span></span>|
|[<span data-ttu-id="55c70-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-270">ReadItem</span></span>|
|[<span data-ttu-id="55c70-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-273">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-273">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="55c70-274">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="55c70-274">dateTimeCreated: Date</span></span>

<span data-ttu-id="55c70-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-277">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-277">Type</span></span>

*   <span data-ttu-id="55c70-278">Дата</span><span class="sxs-lookup"><span data-stu-id="55c70-278">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-279">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-279">Requirements</span></span>

|<span data-ttu-id="55c70-280">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-280">Requirement</span></span>| <span data-ttu-id="55c70-281">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-282">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-283">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-283">1.0</span></span>|
|[<span data-ttu-id="55c70-284">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-285">ReadItem</span></span>|
|[<span data-ttu-id="55c70-286">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-287">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-288">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-288">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="55c70-289">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="55c70-289">dateTimeModified: Date</span></span>

<span data-ttu-id="55c70-290">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-290">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="55c70-291">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-291">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-292">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="55c70-292">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-293">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-293">Type</span></span>

*   <span data-ttu-id="55c70-294">Дата</span><span class="sxs-lookup"><span data-stu-id="55c70-294">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-295">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-295">Requirements</span></span>

|<span data-ttu-id="55c70-296">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-296">Requirement</span></span>| <span data-ttu-id="55c70-297">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-298">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-299">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-299">1.0</span></span>|
|[<span data-ttu-id="55c70-300">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-301">ReadItem</span></span>|
|[<span data-ttu-id="55c70-302">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-303">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-303">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-304">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-304">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="55c70-305">конец: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="55c70-305">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-306">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-306">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="55c70-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="55c70-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="55c70-309">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="55c70-309">Read mode</span></span>

<span data-ttu-id="55c70-310">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="55c70-310">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="55c70-311">Режим создания</span><span class="sxs-lookup"><span data-stu-id="55c70-311">Compose mode</span></span>

<span data-ttu-id="55c70-312">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="55c70-312">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="55c70-313">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="55c70-313">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="55c70-314">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-314">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="55c70-315">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-315">Type</span></span>

*   <span data-ttu-id="55c70-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-317">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-317">Requirements</span></span>

|<span data-ttu-id="55c70-318">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-318">Requirement</span></span>| <span data-ttu-id="55c70-319">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-320">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-321">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-321">1.0</span></span>|
|[<span data-ttu-id="55c70-322">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-323">ReadItem</span></span>|
|[<span data-ttu-id="55c70-324">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-325">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-325">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="55c70-326">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-326">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="55c70-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="55c70-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-331">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="55c70-331">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-332">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-332">Type</span></span>

*   [<span data-ttu-id="55c70-333">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="55c70-333">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="55c70-334">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-334">Requirements</span></span>

|<span data-ttu-id="55c70-335">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-335">Requirement</span></span>| <span data-ttu-id="55c70-336">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-337">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-338">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-338">1.0</span></span>|
|[<span data-ttu-id="55c70-339">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-340">ReadItem</span></span>|
|[<span data-ttu-id="55c70-341">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-342">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-343">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-343">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="55c70-344">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="55c70-344">internetMessageId: String</span></span>

<span data-ttu-id="55c70-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-347">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-347">Type</span></span>

*   <span data-ttu-id="55c70-348">String</span><span class="sxs-lookup"><span data-stu-id="55c70-348">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-349">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-349">Requirements</span></span>

|<span data-ttu-id="55c70-350">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-350">Requirement</span></span>| <span data-ttu-id="55c70-351">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-351">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-352">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-353">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-353">1.0</span></span>|
|[<span data-ttu-id="55c70-354">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-355">ReadItem</span></span>|
|[<span data-ttu-id="55c70-356">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-357">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-357">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-358">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-358">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="55c70-359">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="55c70-359">itemClass: String</span></span>

<span data-ttu-id="55c70-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="55c70-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="55c70-364">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-364">Type</span></span> | <span data-ttu-id="55c70-365">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-365">Description</span></span> | <span data-ttu-id="55c70-366">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="55c70-366">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="55c70-367">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="55c70-367">Appointment items</span></span> | <span data-ttu-id="55c70-368">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="55c70-368">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="55c70-369">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="55c70-369">Message items</span></span> | <span data-ttu-id="55c70-370">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-370">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="55c70-371">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="55c70-371">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-372">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-372">Type</span></span>

*   <span data-ttu-id="55c70-373">String</span><span class="sxs-lookup"><span data-stu-id="55c70-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-374">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-374">Requirements</span></span>

|<span data-ttu-id="55c70-375">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-375">Requirement</span></span>| <span data-ttu-id="55c70-376">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-377">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-378">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-378">1.0</span></span>|
|[<span data-ttu-id="55c70-379">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-380">ReadItem</span></span>|
|[<span data-ttu-id="55c70-381">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-382">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-383">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-383">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="55c70-384">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="55c70-384">(nullable) itemId: String</span></span>

<span data-ttu-id="55c70-385">Получает идентификатор элемента веб-служб Exchange для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-385">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="55c70-386">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-386">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-387">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="55c70-387">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="55c70-388">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="55c70-388">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="55c70-389">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="55c70-389">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="55c70-390">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="55c70-390">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-391">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-391">Type</span></span>

*   <span data-ttu-id="55c70-392">String</span><span class="sxs-lookup"><span data-stu-id="55c70-392">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-393">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-393">Requirements</span></span>

|<span data-ttu-id="55c70-394">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-394">Requirement</span></span>| <span data-ttu-id="55c70-395">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-395">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-396">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-397">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-397">1.0</span></span>|
|[<span data-ttu-id="55c70-398">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-398">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-399">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-399">ReadItem</span></span>|
|[<span data-ttu-id="55c70-400">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-400">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-401">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-401">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-402">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-402">Example</span></span>

<span data-ttu-id="55c70-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="55c70-405">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-405">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-406">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="55c70-406">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="55c70-407">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="55c70-407">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-408">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-408">Type</span></span>

*   [<span data-ttu-id="55c70-409">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="55c70-409">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="55c70-410">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-410">Requirements</span></span>

|<span data-ttu-id="55c70-411">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-411">Requirement</span></span>| <span data-ttu-id="55c70-412">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-413">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-414">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-414">1.0</span></span>|
|[<span data-ttu-id="55c70-415">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-416">ReadItem</span></span>|
|[<span data-ttu-id="55c70-417">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-418">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-418">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-419">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-419">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="55c70-420">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="55c70-420">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-421">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-421">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="55c70-422">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="55c70-422">Read mode</span></span>

<span data-ttu-id="55c70-423">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-423">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="55c70-424">Режим создания</span><span class="sxs-lookup"><span data-stu-id="55c70-424">Compose mode</span></span>

<span data-ttu-id="55c70-425">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-425">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="55c70-426">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-426">Type</span></span>

*   <span data-ttu-id="55c70-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-428">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-428">Requirements</span></span>

|<span data-ttu-id="55c70-429">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-429">Requirement</span></span>| <span data-ttu-id="55c70-430">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-430">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-431">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-431">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-432">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-432">1.0</span></span>|
|[<span data-ttu-id="55c70-433">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-433">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-434">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-434">ReadItem</span></span>|
|[<span data-ttu-id="55c70-435">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-435">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-436">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-436">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="55c70-437">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="55c70-437">normalizedSubject: String</span></span>

<span data-ttu-id="55c70-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="55c70-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="55c70-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-442">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-442">Type</span></span>

*   <span data-ttu-id="55c70-443">String</span><span class="sxs-lookup"><span data-stu-id="55c70-443">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-444">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-444">Requirements</span></span>

|<span data-ttu-id="55c70-445">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-445">Requirement</span></span>| <span data-ttu-id="55c70-446">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-447">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-448">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-448">1.0</span></span>|
|[<span data-ttu-id="55c70-449">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-450">ReadItem</span></span>|
|[<span data-ttu-id="55c70-451">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-452">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-453">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-453">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="55c70-454">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-454">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-455">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="55c70-455">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="55c70-456">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-456">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="55c70-457">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="55c70-457">Read mode</span></span>

<span data-ttu-id="55c70-458">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="55c70-458">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="55c70-459">Режим создания</span><span class="sxs-lookup"><span data-stu-id="55c70-459">Compose mode</span></span>

<span data-ttu-id="55c70-460">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="55c70-460">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="55c70-461">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-461">Type</span></span>

*   <span data-ttu-id="55c70-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-463">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-463">Requirements</span></span>

|<span data-ttu-id="55c70-464">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-464">Requirement</span></span>| <span data-ttu-id="55c70-465">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-466">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-467">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-467">1.0</span></span>|
|[<span data-ttu-id="55c70-468">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-469">ReadItem</span></span>|
|[<span data-ttu-id="55c70-470">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-471">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-471">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="55c70-472">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-472">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-475">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-475">Type</span></span>

*   [<span data-ttu-id="55c70-476">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="55c70-476">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="55c70-477">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-477">Requirements</span></span>

|<span data-ttu-id="55c70-478">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-478">Requirement</span></span>| <span data-ttu-id="55c70-479">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-480">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-481">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-481">1.0</span></span>|
|[<span data-ttu-id="55c70-482">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-482">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-483">ReadItem</span></span>|
|[<span data-ttu-id="55c70-484">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-484">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-485">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-485">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-486">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-486">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="55c70-487">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-487">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-488">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="55c70-488">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="55c70-489">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-489">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="55c70-490">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="55c70-490">Read mode</span></span>

<span data-ttu-id="55c70-491">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="55c70-491">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="55c70-492">Режим создания</span><span class="sxs-lookup"><span data-stu-id="55c70-492">Compose mode</span></span>

<span data-ttu-id="55c70-493">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="55c70-493">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="55c70-494">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-494">Type</span></span>

*   <span data-ttu-id="55c70-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-496">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-496">Requirements</span></span>

|<span data-ttu-id="55c70-497">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-497">Requirement</span></span>| <span data-ttu-id="55c70-498">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-499">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-500">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-500">1.0</span></span>|
|[<span data-ttu-id="55c70-501">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-502">ReadItem</span></span>|
|[<span data-ttu-id="55c70-503">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-504">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-504">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="55c70-505">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-505">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="55c70-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="55c70-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="55c70-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-510">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="55c70-510">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="55c70-511">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-511">Type</span></span>

*   [<span data-ttu-id="55c70-512">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="55c70-512">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="55c70-513">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-513">Requirements</span></span>

|<span data-ttu-id="55c70-514">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-514">Requirement</span></span>| <span data-ttu-id="55c70-515">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-515">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-516">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-516">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-517">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-517">1.0</span></span>|
|[<span data-ttu-id="55c70-518">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-518">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-519">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-519">ReadItem</span></span>|
|[<span data-ttu-id="55c70-520">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-520">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-521">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-521">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-522">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-522">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="55c70-523">Начало: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="55c70-523">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-524">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-524">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="55c70-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="55c70-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="55c70-527">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="55c70-527">Read mode</span></span>

<span data-ttu-id="55c70-528">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="55c70-528">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="55c70-529">Режим создания</span><span class="sxs-lookup"><span data-stu-id="55c70-529">Compose mode</span></span>

<span data-ttu-id="55c70-530">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="55c70-530">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="55c70-531">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="55c70-531">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="55c70-532">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="55c70-532">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="55c70-533">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-533">Type</span></span>

*   <span data-ttu-id="55c70-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-535">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-535">Requirements</span></span>

|<span data-ttu-id="55c70-536">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-536">Requirement</span></span>| <span data-ttu-id="55c70-537">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-538">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-539">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-539">1.0</span></span>|
|[<span data-ttu-id="55c70-540">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-541">ReadItem</span></span>|
|[<span data-ttu-id="55c70-542">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-543">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-543">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="55c70-544">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject?view=outlook-js-1.1) )</span><span class="sxs-lookup"><span data-stu-id="55c70-544">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-545">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-545">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="55c70-546">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="55c70-546">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="55c70-547">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="55c70-547">Read mode</span></span>

<span data-ttu-id="55c70-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="55c70-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="55c70-550">Режим создания</span><span class="sxs-lookup"><span data-stu-id="55c70-550">Compose mode</span></span>

<span data-ttu-id="55c70-551">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="55c70-551">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="55c70-552">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-552">Type</span></span>

*   <span data-ttu-id="55c70-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-554">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-554">Requirements</span></span>

|<span data-ttu-id="55c70-555">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-555">Requirement</span></span>| <span data-ttu-id="55c70-556">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-557">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-558">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-558">1.0</span></span>|
|[<span data-ttu-id="55c70-559">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-560">ReadItem</span></span>|
|[<span data-ttu-id="55c70-561">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-562">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-562">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="55c70-563">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-563">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="55c70-564">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-564">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="55c70-565">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-565">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="55c70-566">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="55c70-566">Read mode</span></span>

<span data-ttu-id="55c70-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="55c70-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="55c70-569">Режим создания</span><span class="sxs-lookup"><span data-stu-id="55c70-569">Compose mode</span></span>

<span data-ttu-id="55c70-570">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-570">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="55c70-571">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-571">Type</span></span>

*   <span data-ttu-id="55c70-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-573">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-573">Requirements</span></span>

|<span data-ttu-id="55c70-574">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-574">Requirement</span></span>| <span data-ttu-id="55c70-575">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-575">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-576">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-576">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-577">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-577">1.0</span></span>|
|[<span data-ttu-id="55c70-578">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-578">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-579">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-579">ReadItem</span></span>|
|[<span data-ttu-id="55c70-580">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-580">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-581">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-581">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="55c70-582">Методы</span><span class="sxs-lookup"><span data-stu-id="55c70-582">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="55c70-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="55c70-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="55c70-584">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="55c70-584">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="55c70-585">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="55c70-585">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="55c70-586">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="55c70-586">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="55c70-587">Параметры</span><span class="sxs-lookup"><span data-stu-id="55c70-587">Parameters</span></span>

|<span data-ttu-id="55c70-588">Имя</span><span class="sxs-lookup"><span data-stu-id="55c70-588">Name</span></span>| <span data-ttu-id="55c70-589">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-589">Type</span></span>| <span data-ttu-id="55c70-590">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="55c70-590">Attributes</span></span>| <span data-ttu-id="55c70-591">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-591">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="55c70-592">String</span><span class="sxs-lookup"><span data-stu-id="55c70-592">String</span></span>||<span data-ttu-id="55c70-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="55c70-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="55c70-595">String</span><span class="sxs-lookup"><span data-stu-id="55c70-595">String</span></span>||<span data-ttu-id="55c70-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="55c70-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="55c70-598">Объект</span><span class="sxs-lookup"><span data-stu-id="55c70-598">Object</span></span>| <span data-ttu-id="55c70-599">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-599">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-600">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="55c70-600">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="55c70-601">Объект</span><span class="sxs-lookup"><span data-stu-id="55c70-601">Object</span></span>| <span data-ttu-id="55c70-602">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-602">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-603">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="55c70-603">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="55c70-604">функция</span><span class="sxs-lookup"><span data-stu-id="55c70-604">function</span></span>| <span data-ttu-id="55c70-605">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-605">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-606">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="55c70-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="55c70-607">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="55c70-607">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="55c70-608">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="55c70-608">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="55c70-609">Ошибки</span><span class="sxs-lookup"><span data-stu-id="55c70-609">Errors</span></span>

| <span data-ttu-id="55c70-610">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="55c70-610">Error code</span></span> | <span data-ttu-id="55c70-611">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-611">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="55c70-612">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="55c70-612">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="55c70-613">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="55c70-613">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="55c70-614">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="55c70-614">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="55c70-615">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-615">Requirements</span></span>

|<span data-ttu-id="55c70-616">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-616">Requirement</span></span>| <span data-ttu-id="55c70-617">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-618">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-619">1.1</span><span class="sxs-lookup"><span data-stu-id="55c70-619">1.1</span></span>|
|[<span data-ttu-id="55c70-620">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="55c70-621">ReadWriteItem</span></span>|
|[<span data-ttu-id="55c70-622">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-623">Создание</span><span class="sxs-lookup"><span data-stu-id="55c70-623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-624">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-624">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="55c70-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="55c70-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="55c70-626">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="55c70-626">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="55c70-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="55c70-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="55c70-630">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="55c70-630">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="55c70-631">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="55c70-631">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="55c70-632">Параметры</span><span class="sxs-lookup"><span data-stu-id="55c70-632">Parameters</span></span>

|<span data-ttu-id="55c70-633">Имя</span><span class="sxs-lookup"><span data-stu-id="55c70-633">Name</span></span>| <span data-ttu-id="55c70-634">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-634">Type</span></span>| <span data-ttu-id="55c70-635">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="55c70-635">Attributes</span></span>| <span data-ttu-id="55c70-636">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-636">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="55c70-637">String</span><span class="sxs-lookup"><span data-stu-id="55c70-637">String</span></span>||<span data-ttu-id="55c70-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="55c70-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="55c70-640">String</span><span class="sxs-lookup"><span data-stu-id="55c70-640">String</span></span>||<span data-ttu-id="55c70-641">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-641">The subject of the item to be attached.</span></span> <span data-ttu-id="55c70-642">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="55c70-642">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="55c70-643">Object</span><span class="sxs-lookup"><span data-stu-id="55c70-643">Object</span></span>| <span data-ttu-id="55c70-644">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-644">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-645">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="55c70-645">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="55c70-646">Объект</span><span class="sxs-lookup"><span data-stu-id="55c70-646">Object</span></span>| <span data-ttu-id="55c70-647">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-647">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-648">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="55c70-648">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="55c70-649">функция</span><span class="sxs-lookup"><span data-stu-id="55c70-649">function</span></span>| <span data-ttu-id="55c70-650">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-650">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-651">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="55c70-651">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="55c70-652">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="55c70-652">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="55c70-653">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="55c70-653">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="55c70-654">Ошибки</span><span class="sxs-lookup"><span data-stu-id="55c70-654">Errors</span></span>

| <span data-ttu-id="55c70-655">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="55c70-655">Error code</span></span> | <span data-ttu-id="55c70-656">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-656">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="55c70-657">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="55c70-657">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="55c70-658">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-658">Requirements</span></span>

|<span data-ttu-id="55c70-659">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-659">Requirement</span></span>| <span data-ttu-id="55c70-660">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-661">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-662">1.1</span><span class="sxs-lookup"><span data-stu-id="55c70-662">1.1</span></span>|
|[<span data-ttu-id="55c70-663">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-664">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="55c70-664">ReadWriteItem</span></span>|
|[<span data-ttu-id="55c70-665">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-666">Создание</span><span class="sxs-lookup"><span data-stu-id="55c70-666">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-667">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-667">Example</span></span>

<span data-ttu-id="55c70-668">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="55c70-668">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="55c70-669">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="55c70-669">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="55c70-670">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-670">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-671">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="55c70-671">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="55c70-672">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="55c70-672">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="55c70-673">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="55c70-673">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-674">Возможность включать вложения в вызове `displayReplyAllForm` не поддерживается в наборе требований 1,1.</span><span class="sxs-lookup"><span data-stu-id="55c70-674">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="55c70-675">Добавлена поддержка вложений `displayReplyAllForm` в наборе требований 1,2 и выше.</span><span class="sxs-lookup"><span data-stu-id="55c70-675">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="55c70-676">Параметры</span><span class="sxs-lookup"><span data-stu-id="55c70-676">Parameters</span></span>

|<span data-ttu-id="55c70-677">Имя</span><span class="sxs-lookup"><span data-stu-id="55c70-677">Name</span></span>| <span data-ttu-id="55c70-678">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-678">Type</span></span>| <span data-ttu-id="55c70-679">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-679">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="55c70-680">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="55c70-680">String &#124; Object</span></span>| |<span data-ttu-id="55c70-p138">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="55c70-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="55c70-683">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="55c70-683">**OR**</span></span><br/><span data-ttu-id="55c70-p139">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="55c70-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="55c70-686">String</span><span class="sxs-lookup"><span data-stu-id="55c70-686">String</span></span> | <span data-ttu-id="55c70-687">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-687">&lt;optional&gt;</span></span> | <span data-ttu-id="55c70-p140">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="55c70-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="55c70-690">функция</span><span class="sxs-lookup"><span data-stu-id="55c70-690">function</span></span> | <span data-ttu-id="55c70-691">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-691">&lt;optional&gt;</span></span> | <span data-ttu-id="55c70-692">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="55c70-692">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="55c70-693">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-693">Requirements</span></span>

|<span data-ttu-id="55c70-694">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-694">Requirement</span></span>| <span data-ttu-id="55c70-695">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-695">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-696">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-696">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-697">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-697">1.0</span></span>|
|[<span data-ttu-id="55c70-698">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-698">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-699">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-699">ReadItem</span></span>|
|[<span data-ttu-id="55c70-700">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-700">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-701">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-701">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="55c70-702">Примеры</span><span class="sxs-lookup"><span data-stu-id="55c70-702">Examples</span></span>

<span data-ttu-id="55c70-703">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="55c70-703">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="55c70-704">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-704">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="55c70-705">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-705">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="55c70-706">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="55c70-706">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="55c70-707">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="55c70-707">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="55c70-708">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-708">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-709">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="55c70-709">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="55c70-710">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="55c70-710">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="55c70-711">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="55c70-711">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-712">Возможность включать вложения в вызове `displayReplyForm` не поддерживается в наборе требований 1,1.</span><span class="sxs-lookup"><span data-stu-id="55c70-712">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="55c70-713">Добавлена поддержка вложений `displayReplyForm` в наборе требований 1,2 и выше.</span><span class="sxs-lookup"><span data-stu-id="55c70-713">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="55c70-714">Параметры</span><span class="sxs-lookup"><span data-stu-id="55c70-714">Parameters</span></span>

|<span data-ttu-id="55c70-715">Имя</span><span class="sxs-lookup"><span data-stu-id="55c70-715">Name</span></span>| <span data-ttu-id="55c70-716">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-716">Type</span></span>| <span data-ttu-id="55c70-717">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-717">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="55c70-718">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="55c70-718">String &#124; Object</span></span>| | <span data-ttu-id="55c70-p142">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="55c70-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="55c70-721">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="55c70-721">**OR**</span></span><br/><span data-ttu-id="55c70-p143">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="55c70-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="55c70-724">String</span><span class="sxs-lookup"><span data-stu-id="55c70-724">String</span></span> | <span data-ttu-id="55c70-725">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-725">&lt;optional&gt;</span></span> | <span data-ttu-id="55c70-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="55c70-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="55c70-728">функция</span><span class="sxs-lookup"><span data-stu-id="55c70-728">function</span></span> | <span data-ttu-id="55c70-729">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-729">&lt;optional&gt;</span></span> | <span data-ttu-id="55c70-730">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="55c70-730">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="55c70-731">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-731">Requirements</span></span>

|<span data-ttu-id="55c70-732">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-732">Requirement</span></span>| <span data-ttu-id="55c70-733">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-734">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-735">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-735">1.0</span></span>|
|[<span data-ttu-id="55c70-736">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-737">ReadItem</span></span>|
|[<span data-ttu-id="55c70-738">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-739">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-739">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="55c70-740">Примеры</span><span class="sxs-lookup"><span data-stu-id="55c70-740">Examples</span></span>

<span data-ttu-id="55c70-741">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="55c70-741">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="55c70-742">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-742">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="55c70-743">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="55c70-743">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="55c70-744">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="55c70-744">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="55c70-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="55c70-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="55c70-746">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-746">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-747">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="55c70-747">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-748">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-748">Requirements</span></span>

|<span data-ttu-id="55c70-749">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-749">Requirement</span></span>| <span data-ttu-id="55c70-750">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-750">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-751">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-752">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-752">1.0</span></span>|
|[<span data-ttu-id="55c70-753">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-754">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-754">ReadItem</span></span>|
|[<span data-ttu-id="55c70-755">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-756">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-756">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="55c70-757">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="55c70-757">Returns:</span></span>

<span data-ttu-id="55c70-758">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="55c70-758">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="55c70-759">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-759">Example</span></span>

<span data-ttu-id="55c70-760">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-760">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="55c70-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="55c70-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="55c70-762">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-762">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-763">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="55c70-763">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="55c70-764">Параметры</span><span class="sxs-lookup"><span data-stu-id="55c70-764">Parameters</span></span>

|<span data-ttu-id="55c70-765">Имя</span><span class="sxs-lookup"><span data-stu-id="55c70-765">Name</span></span>| <span data-ttu-id="55c70-766">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-766">Type</span></span>| <span data-ttu-id="55c70-767">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-767">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="55c70-768">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="55c70-768">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="55c70-769">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="55c70-769">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55c70-770">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-770">Requirements</span></span>

|<span data-ttu-id="55c70-771">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-771">Requirement</span></span>| <span data-ttu-id="55c70-772">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-773">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-774">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-774">1.0</span></span>|
|[<span data-ttu-id="55c70-775">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-775">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-776">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="55c70-776">Restricted</span></span>|
|[<span data-ttu-id="55c70-777">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-777">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-778">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="55c70-779">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="55c70-779">Returns:</span></span>

<span data-ttu-id="55c70-780">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="55c70-780">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="55c70-781">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="55c70-781">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="55c70-782">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="55c70-782">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="55c70-783">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="55c70-783">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="55c70-784">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="55c70-784">Value of `entityType`</span></span> | <span data-ttu-id="55c70-785">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="55c70-785">Type of objects in returned array</span></span> | <span data-ttu-id="55c70-786">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-786">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="55c70-787">String</span><span class="sxs-lookup"><span data-stu-id="55c70-787">String</span></span> | <span data-ttu-id="55c70-788">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="55c70-788">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="55c70-789">Contact</span><span class="sxs-lookup"><span data-stu-id="55c70-789">Contact</span></span> | <span data-ttu-id="55c70-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="55c70-790">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="55c70-791">String</span><span class="sxs-lookup"><span data-stu-id="55c70-791">String</span></span> | <span data-ttu-id="55c70-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="55c70-792">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="55c70-793">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="55c70-793">MeetingSuggestion</span></span> | <span data-ttu-id="55c70-794">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="55c70-794">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="55c70-795">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="55c70-795">PhoneNumber</span></span> | <span data-ttu-id="55c70-796">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="55c70-796">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="55c70-797">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="55c70-797">TaskSuggestion</span></span> | <span data-ttu-id="55c70-798">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="55c70-798">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="55c70-799">String</span><span class="sxs-lookup"><span data-stu-id="55c70-799">String</span></span> | <span data-ttu-id="55c70-800">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="55c70-800">**Restricted**</span></span> |

<span data-ttu-id="55c70-801">Тип:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="55c70-801">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="55c70-802">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-802">Example</span></span>

<span data-ttu-id="55c70-803">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-803">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="55c70-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="55c70-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="55c70-805">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="55c70-805">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-806">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="55c70-806">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="55c70-807">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="55c70-807">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="55c70-808">Параметры</span><span class="sxs-lookup"><span data-stu-id="55c70-808">Parameters</span></span>

|<span data-ttu-id="55c70-809">Имя</span><span class="sxs-lookup"><span data-stu-id="55c70-809">Name</span></span>| <span data-ttu-id="55c70-810">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-810">Type</span></span>| <span data-ttu-id="55c70-811">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-811">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="55c70-812">String</span><span class="sxs-lookup"><span data-stu-id="55c70-812">String</span></span>|<span data-ttu-id="55c70-813">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="55c70-813">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55c70-814">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-814">Requirements</span></span>

|<span data-ttu-id="55c70-815">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-815">Requirement</span></span>| <span data-ttu-id="55c70-816">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-817">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-818">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-818">1.0</span></span>|
|[<span data-ttu-id="55c70-819">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-820">ReadItem</span></span>|
|[<span data-ttu-id="55c70-821">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-822">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="55c70-823">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="55c70-823">Returns:</span></span>

<span data-ttu-id="55c70-p146">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="55c70-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="55c70-826">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="55c70-826">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="55c70-827">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="55c70-827">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="55c70-828">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="55c70-828">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-829">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="55c70-829">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="55c70-p147">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="55c70-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="55c70-833">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="55c70-833">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="55c70-834">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="55c70-834">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="55c70-p148">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="55c70-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="55c70-837">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-837">Requirements</span></span>

|<span data-ttu-id="55c70-838">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-838">Requirement</span></span>| <span data-ttu-id="55c70-839">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-840">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-841">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-841">1.0</span></span>|
|[<span data-ttu-id="55c70-842">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-842">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-843">ReadItem</span></span>|
|[<span data-ttu-id="55c70-844">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-844">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-845">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="55c70-846">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="55c70-846">Returns:</span></span>

<span data-ttu-id="55c70-p149">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="55c70-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="55c70-849">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="55c70-849">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="55c70-850">Object</span><span class="sxs-lookup"><span data-stu-id="55c70-850">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="55c70-851">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-851">Example</span></span>

<span data-ttu-id="55c70-852">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="55c70-852">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="55c70-853">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="55c70-853">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="55c70-854">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="55c70-854">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="55c70-855">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="55c70-855">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="55c70-856">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="55c70-856">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="55c70-p150">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="55c70-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="55c70-859">Параметры</span><span class="sxs-lookup"><span data-stu-id="55c70-859">Parameters</span></span>

|<span data-ttu-id="55c70-860">Имя</span><span class="sxs-lookup"><span data-stu-id="55c70-860">Name</span></span>| <span data-ttu-id="55c70-861">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-861">Type</span></span>| <span data-ttu-id="55c70-862">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-862">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="55c70-863">String</span><span class="sxs-lookup"><span data-stu-id="55c70-863">String</span></span>|<span data-ttu-id="55c70-864">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="55c70-864">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55c70-865">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-865">Requirements</span></span>

|<span data-ttu-id="55c70-866">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-866">Requirement</span></span>| <span data-ttu-id="55c70-867">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-867">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-868">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-868">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-869">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-869">1.0</span></span>|
|[<span data-ttu-id="55c70-870">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-870">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-871">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-871">ReadItem</span></span>|
|[<span data-ttu-id="55c70-872">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-872">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-873">Чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-873">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="55c70-874">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="55c70-874">Returns:</span></span>

<span data-ttu-id="55c70-875">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="55c70-875">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="55c70-876">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="55c70-876">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="55c70-877">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="55c70-877">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="55c70-878">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-878">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="55c70-879">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="55c70-879">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="55c70-880">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="55c70-880">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="55c70-p151">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="55c70-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="55c70-884">Параметры</span><span class="sxs-lookup"><span data-stu-id="55c70-884">Parameters</span></span>

|<span data-ttu-id="55c70-885">Имя</span><span class="sxs-lookup"><span data-stu-id="55c70-885">Name</span></span>| <span data-ttu-id="55c70-886">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-886">Type</span></span>| <span data-ttu-id="55c70-887">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="55c70-887">Attributes</span></span>| <span data-ttu-id="55c70-888">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-888">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="55c70-889">function</span><span class="sxs-lookup"><span data-stu-id="55c70-889">function</span></span>||<span data-ttu-id="55c70-890">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="55c70-890">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="55c70-891">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="55c70-891">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="55c70-892">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="55c70-892">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="55c70-893">Объект</span><span class="sxs-lookup"><span data-stu-id="55c70-893">Object</span></span>| <span data-ttu-id="55c70-894">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-894">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-895">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="55c70-895">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="55c70-896">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="55c70-896">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55c70-897">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-897">Requirements</span></span>

|<span data-ttu-id="55c70-898">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-898">Requirement</span></span>| <span data-ttu-id="55c70-899">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-899">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-900">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="55c70-900">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-901">1.0</span><span class="sxs-lookup"><span data-stu-id="55c70-901">1.0</span></span>|
|[<span data-ttu-id="55c70-902">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-902">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-903">ReadItem</span><span class="sxs-lookup"><span data-stu-id="55c70-903">ReadItem</span></span>|
|[<span data-ttu-id="55c70-904">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-904">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-905">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="55c70-905">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-906">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-906">Example</span></span>

<span data-ttu-id="55c70-p154">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="55c70-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="55c70-910">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="55c70-910">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="55c70-911">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="55c70-911">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="55c70-912">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="55c70-912">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="55c70-913">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="55c70-913">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="55c70-914">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="55c70-914">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="55c70-915">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="55c70-915">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="55c70-916">Параметры</span><span class="sxs-lookup"><span data-stu-id="55c70-916">Parameters</span></span>

|<span data-ttu-id="55c70-917">Имя</span><span class="sxs-lookup"><span data-stu-id="55c70-917">Name</span></span>| <span data-ttu-id="55c70-918">Тип</span><span class="sxs-lookup"><span data-stu-id="55c70-918">Type</span></span>| <span data-ttu-id="55c70-919">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="55c70-919">Attributes</span></span>| <span data-ttu-id="55c70-920">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-920">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="55c70-921">String</span><span class="sxs-lookup"><span data-stu-id="55c70-921">String</span></span>||<span data-ttu-id="55c70-922">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="55c70-922">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="55c70-923">Object</span><span class="sxs-lookup"><span data-stu-id="55c70-923">Object</span></span>| <span data-ttu-id="55c70-924">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-924">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-925">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="55c70-925">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="55c70-926">Объект</span><span class="sxs-lookup"><span data-stu-id="55c70-926">Object</span></span>| <span data-ttu-id="55c70-927">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-927">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-928">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="55c70-928">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="55c70-929">функция</span><span class="sxs-lookup"><span data-stu-id="55c70-929">function</span></span>| <span data-ttu-id="55c70-930">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="55c70-930">&lt;optional&gt;</span></span>|<span data-ttu-id="55c70-931">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="55c70-931">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="55c70-932">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="55c70-932">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="55c70-933">Ошибки</span><span class="sxs-lookup"><span data-stu-id="55c70-933">Errors</span></span>

| <span data-ttu-id="55c70-934">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="55c70-934">Error code</span></span> | <span data-ttu-id="55c70-935">Описание</span><span class="sxs-lookup"><span data-stu-id="55c70-935">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="55c70-936">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="55c70-936">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="55c70-937">Требования</span><span class="sxs-lookup"><span data-stu-id="55c70-937">Requirements</span></span>

|<span data-ttu-id="55c70-938">Требование</span><span class="sxs-lookup"><span data-stu-id="55c70-938">Requirement</span></span>| <span data-ttu-id="55c70-939">Значение</span><span class="sxs-lookup"><span data-stu-id="55c70-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="55c70-940">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="55c70-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55c70-941">1.1</span><span class="sxs-lookup"><span data-stu-id="55c70-941">1.1</span></span>|
|[<span data-ttu-id="55c70-942">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="55c70-942">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="55c70-943">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="55c70-943">ReadWriteItem</span></span>|
|[<span data-ttu-id="55c70-944">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="55c70-944">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55c70-945">Создание</span><span class="sxs-lookup"><span data-stu-id="55c70-945">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="55c70-946">Пример</span><span class="sxs-lookup"><span data-stu-id="55c70-946">Example</span></span>

<span data-ttu-id="55c70-947">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="55c70-947">The following code removes an attachment with an identifier of '0'.</span></span>

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
