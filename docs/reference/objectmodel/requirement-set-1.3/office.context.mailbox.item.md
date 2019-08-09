---
title: Office. Context. Mailbox. Item — набор требований 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 5f9ef8b8018dc97dfba7d8e1509bd510dc2b920b
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268413"
---
# <a name="item"></a><span data-ttu-id="898ce-102">item</span><span class="sxs-lookup"><span data-stu-id="898ce-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="898ce-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="898ce-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="898ce-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="898ce-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="898ce-106">Requirements</span></span>

|<span data-ttu-id="898ce-107">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-107">Requirement</span></span>| <span data-ttu-id="898ce-108">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-110">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-110">1.0</span></span>|
|[<span data-ttu-id="898ce-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="898ce-112">Restricted</span></span>|
|[<span data-ttu-id="898ce-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="898ce-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="898ce-115">Members and methods</span></span>

| <span data-ttu-id="898ce-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-116">Member</span></span> | <span data-ttu-id="898ce-117">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="898ce-118">attachments</span><span class="sxs-lookup"><span data-stu-id="898ce-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="898ce-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-119">Member</span></span> |
| [<span data-ttu-id="898ce-120">bcc</span><span class="sxs-lookup"><span data-stu-id="898ce-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="898ce-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-121">Member</span></span> |
| [<span data-ttu-id="898ce-122">body</span><span class="sxs-lookup"><span data-stu-id="898ce-122">body</span></span>](#body-body) | <span data-ttu-id="898ce-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-123">Member</span></span> |
| [<span data-ttu-id="898ce-124">cc</span><span class="sxs-lookup"><span data-stu-id="898ce-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="898ce-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-125">Member</span></span> |
| [<span data-ttu-id="898ce-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="898ce-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="898ce-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-127">Member</span></span> |
| [<span data-ttu-id="898ce-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="898ce-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="898ce-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-129">Member</span></span> |
| [<span data-ttu-id="898ce-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="898ce-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="898ce-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-131">Member</span></span> |
| [<span data-ttu-id="898ce-132">end</span><span class="sxs-lookup"><span data-stu-id="898ce-132">end</span></span>](#end-datetime) | <span data-ttu-id="898ce-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-133">Member</span></span> |
| [<span data-ttu-id="898ce-134">from</span><span class="sxs-lookup"><span data-stu-id="898ce-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="898ce-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-135">Member</span></span> |
| [<span data-ttu-id="898ce-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="898ce-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="898ce-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-137">Member</span></span> |
| [<span data-ttu-id="898ce-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="898ce-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="898ce-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-139">Member</span></span> |
| [<span data-ttu-id="898ce-140">itemId</span><span class="sxs-lookup"><span data-stu-id="898ce-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="898ce-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-141">Member</span></span> |
| [<span data-ttu-id="898ce-142">itemType</span><span class="sxs-lookup"><span data-stu-id="898ce-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="898ce-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-143">Member</span></span> |
| [<span data-ttu-id="898ce-144">location</span><span class="sxs-lookup"><span data-stu-id="898ce-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="898ce-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-145">Member</span></span> |
| [<span data-ttu-id="898ce-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="898ce-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="898ce-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-147">Member</span></span> |
| [<span data-ttu-id="898ce-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="898ce-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="898ce-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-149">Member</span></span> |
| [<span data-ttu-id="898ce-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="898ce-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="898ce-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-151">Member</span></span> |
| [<span data-ttu-id="898ce-152">organizer</span><span class="sxs-lookup"><span data-stu-id="898ce-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="898ce-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-153">Member</span></span> |
| [<span data-ttu-id="898ce-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="898ce-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="898ce-155">Member</span><span class="sxs-lookup"><span data-stu-id="898ce-155">Member</span></span> |
| [<span data-ttu-id="898ce-156">sender</span><span class="sxs-lookup"><span data-stu-id="898ce-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="898ce-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-157">Member</span></span> |
| [<span data-ttu-id="898ce-158">start</span><span class="sxs-lookup"><span data-stu-id="898ce-158">start</span></span>](#start-datetime) | <span data-ttu-id="898ce-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-159">Member</span></span> |
| [<span data-ttu-id="898ce-160">subject</span><span class="sxs-lookup"><span data-stu-id="898ce-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="898ce-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-161">Member</span></span> |
| [<span data-ttu-id="898ce-162">to</span><span class="sxs-lookup"><span data-stu-id="898ce-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="898ce-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="898ce-163">Member</span></span> |
| [<span data-ttu-id="898ce-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="898ce-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="898ce-165">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-165">Method</span></span> |
| [<span data-ttu-id="898ce-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="898ce-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="898ce-167">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-167">Method</span></span> |
| [<span data-ttu-id="898ce-168">close</span><span class="sxs-lookup"><span data-stu-id="898ce-168">close</span></span>](#close) | <span data-ttu-id="898ce-169">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-169">Method</span></span> |
| [<span data-ttu-id="898ce-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="898ce-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="898ce-171">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-171">Method</span></span> |
| [<span data-ttu-id="898ce-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="898ce-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="898ce-173">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-173">Method</span></span> |
| [<span data-ttu-id="898ce-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="898ce-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="898ce-175">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-175">Method</span></span> |
| [<span data-ttu-id="898ce-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="898ce-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="898ce-177">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-177">Method</span></span> |
| [<span data-ttu-id="898ce-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="898ce-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="898ce-179">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-179">Method</span></span> |
| [<span data-ttu-id="898ce-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="898ce-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="898ce-181">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-181">Method</span></span> |
| [<span data-ttu-id="898ce-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="898ce-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="898ce-183">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-183">Method</span></span> |
| [<span data-ttu-id="898ce-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="898ce-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="898ce-185">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-185">Method</span></span> |
| [<span data-ttu-id="898ce-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="898ce-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="898ce-187">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-187">Method</span></span> |
| [<span data-ttu-id="898ce-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="898ce-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="898ce-189">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-189">Method</span></span> |
| [<span data-ttu-id="898ce-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="898ce-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="898ce-191">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-191">Method</span></span> |
| [<span data-ttu-id="898ce-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="898ce-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="898ce-193">Метод</span><span class="sxs-lookup"><span data-stu-id="898ce-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="898ce-194">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-194">Example</span></span>

<span data-ttu-id="898ce-195">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="898ce-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="898ce-196">Элементы</span><span class="sxs-lookup"><span data-stu-id="898ce-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-13"></a><span data-ttu-id="898ce-197">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="898ce-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

<span data-ttu-id="898ce-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-200">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="898ce-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="898ce-201">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="898ce-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-202">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-202">Type</span></span>

*   <span data-ttu-id="898ce-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="898ce-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-204">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-204">Requirements</span></span>

|<span data-ttu-id="898ce-205">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-205">Requirement</span></span>| <span data-ttu-id="898ce-206">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-208">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-208">1.0</span></span>|
|[<span data-ttu-id="898ce-209">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-210">ReadItem</span></span>|
|[<span data-ttu-id="898ce-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-212">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-213">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-213">Example</span></span>

<span data-ttu-id="898ce-214">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="898ce-215">СК: [получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-216">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="898ce-217">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="898ce-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-218">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-218">Type</span></span>

*   [<span data-ttu-id="898ce-219">Получатели</span><span class="sxs-lookup"><span data-stu-id="898ce-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="898ce-220">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-220">Requirements</span></span>

|<span data-ttu-id="898ce-221">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-221">Requirement</span></span>| <span data-ttu-id="898ce-222">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-224">1.1</span><span class="sxs-lookup"><span data-stu-id="898ce-224">1.1</span></span>|
|[<span data-ttu-id="898ce-225">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-226">ReadItem</span></span>|
|[<span data-ttu-id="898ce-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-228">Создание</span><span class="sxs-lookup"><span data-stu-id="898ce-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-229">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-13"></a><span data-ttu-id="898ce-230">основной текст: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-231">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-232">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-232">Type</span></span>

*   [<span data-ttu-id="898ce-233">Body</span><span class="sxs-lookup"><span data-stu-id="898ce-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="898ce-234">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-234">Requirements</span></span>

|<span data-ttu-id="898ce-235">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-235">Requirement</span></span>| <span data-ttu-id="898ce-236">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-238">1.1</span><span class="sxs-lookup"><span data-stu-id="898ce-238">1.1</span></span>|
|[<span data-ttu-id="898ce-239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-240">ReadItem</span></span>|
|[<span data-ttu-id="898ce-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-242">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-243">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-243">Example</span></span>

<span data-ttu-id="898ce-244">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="898ce-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="898ce-245">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="898ce-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="898ce-246">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-247">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="898ce-248">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="898ce-249">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="898ce-249">Read mode</span></span>

<span data-ttu-id="898ce-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="898ce-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="898ce-252">Режим создания</span><span class="sxs-lookup"><span data-stu-id="898ce-252">Compose mode</span></span>

<span data-ttu-id="898ce-253">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="898ce-254">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-254">Type</span></span>

*   <span data-ttu-id="898ce-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-256">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-256">Requirements</span></span>

|<span data-ttu-id="898ce-257">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-257">Requirement</span></span>| <span data-ttu-id="898ce-258">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-259">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-260">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-260">1.0</span></span>|
|[<span data-ttu-id="898ce-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-262">ReadItem</span></span>|
|[<span data-ttu-id="898ce-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-264">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="898ce-265">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="898ce-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="898ce-266">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="898ce-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="898ce-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="898ce-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="898ce-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="898ce-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-271">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-271">Type</span></span>

*   <span data-ttu-id="898ce-272">String</span><span class="sxs-lookup"><span data-stu-id="898ce-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-273">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-273">Requirements</span></span>

|<span data-ttu-id="898ce-274">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-274">Requirement</span></span>| <span data-ttu-id="898ce-275">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-276">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-277">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-277">1.0</span></span>|
|[<span data-ttu-id="898ce-278">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-279">ReadItem</span></span>|
|[<span data-ttu-id="898ce-280">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-281">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-282">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="898ce-283">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="898ce-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="898ce-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-286">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-286">Type</span></span>

*   <span data-ttu-id="898ce-287">Дата</span><span class="sxs-lookup"><span data-stu-id="898ce-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-288">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-288">Requirements</span></span>

|<span data-ttu-id="898ce-289">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-289">Requirement</span></span>| <span data-ttu-id="898ce-290">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-291">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-292">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-292">1.0</span></span>|
|[<span data-ttu-id="898ce-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-294">ReadItem</span></span>|
|[<span data-ttu-id="898ce-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-296">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-297">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="898ce-298">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="898ce-298">dateTimeModified: Date</span></span>

<span data-ttu-id="898ce-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-301">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="898ce-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-302">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-302">Type</span></span>

*   <span data-ttu-id="898ce-303">Дата</span><span class="sxs-lookup"><span data-stu-id="898ce-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-304">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-304">Requirements</span></span>

|<span data-ttu-id="898ce-305">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-305">Requirement</span></span>| <span data-ttu-id="898ce-306">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-307">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-308">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-308">1.0</span></span>|
|[<span data-ttu-id="898ce-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-310">ReadItem</span></span>|
|[<span data-ttu-id="898ce-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-313">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="898ce-314">конец: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.3) )</span><span class="sxs-lookup"><span data-stu-id="898ce-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-315">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="898ce-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="898ce-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="898ce-318">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="898ce-318">Read mode</span></span>

<span data-ttu-id="898ce-319">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="898ce-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="898ce-320">Режим создания</span><span class="sxs-lookup"><span data-stu-id="898ce-320">Compose mode</span></span>

<span data-ttu-id="898ce-321">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="898ce-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="898ce-322">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="898ce-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="898ce-323">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="898ce-324">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-324">Type</span></span>

*   <span data-ttu-id="898ce-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-326">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-326">Requirements</span></span>

|<span data-ttu-id="898ce-327">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-327">Requirement</span></span>| <span data-ttu-id="898ce-328">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-329">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-330">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-330">1.0</span></span>|
|[<span data-ttu-id="898ce-331">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-332">ReadItem</span></span>|
|[<span data-ttu-id="898ce-333">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-334">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-334">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="898ce-335">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="898ce-p113">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="898ce-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-340">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="898ce-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-341">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-341">Type</span></span>

*   [<span data-ttu-id="898ce-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="898ce-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="898ce-343">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-343">Requirements</span></span>

|<span data-ttu-id="898ce-344">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-344">Requirement</span></span>| <span data-ttu-id="898ce-345">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-346">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-347">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-347">1.0</span></span>|
|[<span data-ttu-id="898ce-348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-349">ReadItem</span></span>|
|[<span data-ttu-id="898ce-350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-351">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-352">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="898ce-353">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="898ce-353">internetMessageId: String</span></span>

<span data-ttu-id="898ce-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-356">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-356">Type</span></span>

*   <span data-ttu-id="898ce-357">String</span><span class="sxs-lookup"><span data-stu-id="898ce-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-358">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-358">Requirements</span></span>

|<span data-ttu-id="898ce-359">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-359">Requirement</span></span>| <span data-ttu-id="898ce-360">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-361">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-362">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-362">1.0</span></span>|
|[<span data-ttu-id="898ce-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-364">ReadItem</span></span>|
|[<span data-ttu-id="898ce-365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-366">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-367">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="898ce-368">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="898ce-368">itemClass: String</span></span>

<span data-ttu-id="898ce-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="898ce-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="898ce-373">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-373">Type</span></span> | <span data-ttu-id="898ce-374">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-374">Description</span></span> | <span data-ttu-id="898ce-375">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="898ce-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="898ce-376">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="898ce-376">Appointment items</span></span> | <span data-ttu-id="898ce-377">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="898ce-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="898ce-378">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="898ce-378">Message items</span></span> | <span data-ttu-id="898ce-379">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="898ce-380">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="898ce-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-381">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-381">Type</span></span>

*   <span data-ttu-id="898ce-382">String</span><span class="sxs-lookup"><span data-stu-id="898ce-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-383">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-383">Requirements</span></span>

|<span data-ttu-id="898ce-384">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-384">Requirement</span></span>| <span data-ttu-id="898ce-385">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-386">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-387">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-387">1.0</span></span>|
|[<span data-ttu-id="898ce-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-389">ReadItem</span></span>|
|[<span data-ttu-id="898ce-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-392">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="898ce-393">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="898ce-393">(nullable) itemId: String</span></span>

<span data-ttu-id="898ce-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-396">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="898ce-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="898ce-397">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="898ce-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="898ce-398">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="898ce-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="898ce-399">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="898ce-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="898ce-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-402">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-402">Type</span></span>

*   <span data-ttu-id="898ce-403">String</span><span class="sxs-lookup"><span data-stu-id="898ce-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-404">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-404">Requirements</span></span>

|<span data-ttu-id="898ce-405">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-405">Requirement</span></span>| <span data-ttu-id="898ce-406">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-407">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-408">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-408">1.0</span></span>|
|[<span data-ttu-id="898ce-409">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-410">ReadItem</span></span>|
|[<span data-ttu-id="898ce-411">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-412">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-413">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-413">Example</span></span>

<span data-ttu-id="898ce-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-13"></a><span data-ttu-id="898ce-416">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-417">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="898ce-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="898ce-418">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="898ce-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-419">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-419">Type</span></span>

*   [<span data-ttu-id="898ce-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="898ce-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="898ce-421">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-421">Requirements</span></span>

|<span data-ttu-id="898ce-422">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-422">Requirement</span></span>| <span data-ttu-id="898ce-423">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-424">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-425">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-425">1.0</span></span>|
|[<span data-ttu-id="898ce-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-427">ReadItem</span></span>|
|[<span data-ttu-id="898ce-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-429">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-430">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-13"></a><span data-ttu-id="898ce-431">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location?view=outlook-js-1.3) )</span><span class="sxs-lookup"><span data-stu-id="898ce-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-432">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="898ce-433">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="898ce-433">Read mode</span></span>

<span data-ttu-id="898ce-434">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="898ce-435">Режим создания</span><span class="sxs-lookup"><span data-stu-id="898ce-435">Compose mode</span></span>

<span data-ttu-id="898ce-436">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="898ce-437">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-437">Type</span></span>

*   <span data-ttu-id="898ce-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-439">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-439">Requirements</span></span>

|<span data-ttu-id="898ce-440">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-440">Requirement</span></span>| <span data-ttu-id="898ce-441">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-442">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-443">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-443">1.0</span></span>|
|[<span data-ttu-id="898ce-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-445">ReadItem</span></span>|
|[<span data-ttu-id="898ce-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-447">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-447">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="898ce-448">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="898ce-448">normalizedSubject: String</span></span>

<span data-ttu-id="898ce-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="898ce-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="898ce-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-453">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-453">Type</span></span>

*   <span data-ttu-id="898ce-454">String</span><span class="sxs-lookup"><span data-stu-id="898ce-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-455">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-455">Requirements</span></span>

|<span data-ttu-id="898ce-456">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-456">Requirement</span></span>| <span data-ttu-id="898ce-457">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-458">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-459">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-459">1.0</span></span>|
|[<span data-ttu-id="898ce-460">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-461">ReadItem</span></span>|
|[<span data-ttu-id="898ce-462">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-463">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-464">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-13"></a><span data-ttu-id="898ce-465">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-466">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-467">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-467">Type</span></span>

*   [<span data-ttu-id="898ce-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="898ce-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="898ce-469">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-469">Requirements</span></span>

|<span data-ttu-id="898ce-470">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-470">Requirement</span></span>| <span data-ttu-id="898ce-471">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-472">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="898ce-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-473">1.3</span><span class="sxs-lookup"><span data-stu-id="898ce-473">1.3</span></span>|
|[<span data-ttu-id="898ce-474">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-475">ReadItem</span></span>|
|[<span data-ttu-id="898ce-476">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-477">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-478">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="898ce-479">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-480">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="898ce-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="898ce-481">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="898ce-482">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="898ce-482">Read mode</span></span>

<span data-ttu-id="898ce-483">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="898ce-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="898ce-484">Режим создания</span><span class="sxs-lookup"><span data-stu-id="898ce-484">Compose mode</span></span>

<span data-ttu-id="898ce-485">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="898ce-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="898ce-486">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-486">Type</span></span>

*   <span data-ttu-id="898ce-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-488">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-488">Requirements</span></span>

|<span data-ttu-id="898ce-489">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-489">Requirement</span></span>| <span data-ttu-id="898ce-490">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-491">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-492">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-492">1.0</span></span>|
|[<span data-ttu-id="898ce-493">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-494">ReadItem</span></span>|
|[<span data-ttu-id="898ce-495">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-496">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-496">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="898ce-497">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-500">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-500">Type</span></span>

*   [<span data-ttu-id="898ce-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="898ce-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="898ce-502">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-502">Requirements</span></span>

|<span data-ttu-id="898ce-503">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-503">Requirement</span></span>| <span data-ttu-id="898ce-504">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-505">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-506">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-506">1.0</span></span>|
|[<span data-ttu-id="898ce-507">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-508">ReadItem</span></span>|
|[<span data-ttu-id="898ce-509">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-510">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-511">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="898ce-512">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-513">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="898ce-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="898ce-514">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="898ce-515">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="898ce-515">Read mode</span></span>

<span data-ttu-id="898ce-516">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="898ce-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="898ce-517">Режим создания</span><span class="sxs-lookup"><span data-stu-id="898ce-517">Compose mode</span></span>

<span data-ttu-id="898ce-518">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="898ce-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="898ce-519">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-519">Type</span></span>

*   <span data-ttu-id="898ce-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-521">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-521">Requirements</span></span>

|<span data-ttu-id="898ce-522">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-522">Requirement</span></span>| <span data-ttu-id="898ce-523">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-524">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-525">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-525">1.0</span></span>|
|[<span data-ttu-id="898ce-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-527">ReadItem</span></span>|
|[<span data-ttu-id="898ce-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-529">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-529">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="898ce-530">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="898ce-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="898ce-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="898ce-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-535">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="898ce-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="898ce-536">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-536">Type</span></span>

*   [<span data-ttu-id="898ce-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="898ce-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="898ce-538">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-538">Requirements</span></span>

|<span data-ttu-id="898ce-539">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-539">Requirement</span></span>| <span data-ttu-id="898ce-540">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-541">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-542">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-542">1.0</span></span>|
|[<span data-ttu-id="898ce-543">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-544">ReadItem</span></span>|
|[<span data-ttu-id="898ce-545">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-546">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-547">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="898ce-548">Начало: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.3) )</span><span class="sxs-lookup"><span data-stu-id="898ce-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-549">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="898ce-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="898ce-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="898ce-552">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="898ce-552">Read mode</span></span>

<span data-ttu-id="898ce-553">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="898ce-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="898ce-554">Режим создания</span><span class="sxs-lookup"><span data-stu-id="898ce-554">Compose mode</span></span>

<span data-ttu-id="898ce-555">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="898ce-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="898ce-556">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="898ce-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="898ce-557">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="898ce-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="898ce-558">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-558">Type</span></span>

*   <span data-ttu-id="898ce-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-560">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-560">Requirements</span></span>

|<span data-ttu-id="898ce-561">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-561">Requirement</span></span>| <span data-ttu-id="898ce-562">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-563">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-564">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-564">1.0</span></span>|
|[<span data-ttu-id="898ce-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-566">ReadItem</span></span>|
|[<span data-ttu-id="898ce-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-568">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-13"></a><span data-ttu-id="898ce-569">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject?view=outlook-js-1.3) )</span><span class="sxs-lookup"><span data-stu-id="898ce-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-570">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="898ce-571">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="898ce-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="898ce-572">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="898ce-572">Read mode</span></span>

<span data-ttu-id="898ce-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="898ce-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="898ce-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="898ce-575">Compose mode</span></span>

<span data-ttu-id="898ce-576">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="898ce-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="898ce-577">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-577">Type</span></span>

*   <span data-ttu-id="898ce-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-579">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-579">Requirements</span></span>

|<span data-ttu-id="898ce-580">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-580">Requirement</span></span>| <span data-ttu-id="898ce-581">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-582">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-583">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-583">1.0</span></span>|
|[<span data-ttu-id="898ce-584">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-585">ReadItem</span></span>|
|[<span data-ttu-id="898ce-586">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-587">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-587">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="898ce-588">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="898ce-589">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="898ce-590">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="898ce-591">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="898ce-591">Read mode</span></span>

<span data-ttu-id="898ce-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="898ce-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="898ce-594">Режим создания</span><span class="sxs-lookup"><span data-stu-id="898ce-594">Compose mode</span></span>

<span data-ttu-id="898ce-595">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="898ce-596">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-596">Type</span></span>

*   <span data-ttu-id="898ce-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-598">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-598">Requirements</span></span>

|<span data-ttu-id="898ce-599">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-599">Requirement</span></span>| <span data-ttu-id="898ce-600">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-601">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-602">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-602">1.0</span></span>|
|[<span data-ttu-id="898ce-603">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-604">ReadItem</span></span>|
|[<span data-ttu-id="898ce-605">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-606">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="898ce-607">Методы</span><span class="sxs-lookup"><span data-stu-id="898ce-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="898ce-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="898ce-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="898ce-609">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="898ce-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="898ce-610">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="898ce-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="898ce-611">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="898ce-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-612">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-612">Parameters</span></span>

|<span data-ttu-id="898ce-613">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-613">Name</span></span>| <span data-ttu-id="898ce-614">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-614">Type</span></span>| <span data-ttu-id="898ce-615">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="898ce-615">Attributes</span></span>| <span data-ttu-id="898ce-616">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="898ce-617">String</span><span class="sxs-lookup"><span data-stu-id="898ce-617">String</span></span>||<span data-ttu-id="898ce-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="898ce-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="898ce-620">String</span><span class="sxs-lookup"><span data-stu-id="898ce-620">String</span></span>||<span data-ttu-id="898ce-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="898ce-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="898ce-623">Object</span><span class="sxs-lookup"><span data-stu-id="898ce-623">Object</span></span>| <span data-ttu-id="898ce-624">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-624">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-625">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="898ce-625">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="898ce-626">Object</span><span class="sxs-lookup"><span data-stu-id="898ce-626">Object</span></span>| <span data-ttu-id="898ce-627">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-627">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-628">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="898ce-628">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="898ce-629">функция</span><span class="sxs-lookup"><span data-stu-id="898ce-629">function</span></span>| <span data-ttu-id="898ce-630">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-630">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-631">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="898ce-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="898ce-632">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="898ce-632">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="898ce-633">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="898ce-633">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="898ce-634">Ошибки</span><span class="sxs-lookup"><span data-stu-id="898ce-634">Errors</span></span>

| <span data-ttu-id="898ce-635">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="898ce-635">Error code</span></span> | <span data-ttu-id="898ce-636">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-636">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="898ce-637">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="898ce-637">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="898ce-638">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="898ce-638">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="898ce-639">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="898ce-639">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="898ce-640">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-640">Requirements</span></span>

|<span data-ttu-id="898ce-641">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-641">Requirement</span></span>| <span data-ttu-id="898ce-642">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-643">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-644">1.1</span><span class="sxs-lookup"><span data-stu-id="898ce-644">1.1</span></span>|
|[<span data-ttu-id="898ce-645">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-646">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="898ce-646">ReadWriteItem</span></span>|
|[<span data-ttu-id="898ce-647">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-648">Создание</span><span class="sxs-lookup"><span data-stu-id="898ce-648">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-649">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-649">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="898ce-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="898ce-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="898ce-651">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="898ce-651">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="898ce-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="898ce-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="898ce-655">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="898ce-655">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="898ce-656">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="898ce-656">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-657">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-657">Parameters</span></span>

|<span data-ttu-id="898ce-658">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-658">Name</span></span>| <span data-ttu-id="898ce-659">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-659">Type</span></span>| <span data-ttu-id="898ce-660">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="898ce-660">Attributes</span></span>| <span data-ttu-id="898ce-661">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-661">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="898ce-662">String</span><span class="sxs-lookup"><span data-stu-id="898ce-662">String</span></span>||<span data-ttu-id="898ce-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="898ce-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="898ce-665">String</span><span class="sxs-lookup"><span data-stu-id="898ce-665">String</span></span>||<span data-ttu-id="898ce-666">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-666">The subject of the item to be attached.</span></span> <span data-ttu-id="898ce-667">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="898ce-667">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="898ce-668">Object</span><span class="sxs-lookup"><span data-stu-id="898ce-668">Object</span></span>| <span data-ttu-id="898ce-669">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-669">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-670">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="898ce-670">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="898ce-671">Объект</span><span class="sxs-lookup"><span data-stu-id="898ce-671">Object</span></span>| <span data-ttu-id="898ce-672">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-672">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-673">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="898ce-673">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="898ce-674">функция</span><span class="sxs-lookup"><span data-stu-id="898ce-674">function</span></span>| <span data-ttu-id="898ce-675">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-675">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-676">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="898ce-676">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="898ce-677">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="898ce-677">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="898ce-678">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="898ce-678">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="898ce-679">Ошибки</span><span class="sxs-lookup"><span data-stu-id="898ce-679">Errors</span></span>

| <span data-ttu-id="898ce-680">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="898ce-680">Error code</span></span> | <span data-ttu-id="898ce-681">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-681">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="898ce-682">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="898ce-682">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="898ce-683">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-683">Requirements</span></span>

|<span data-ttu-id="898ce-684">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-684">Requirement</span></span>| <span data-ttu-id="898ce-685">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-685">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-686">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-686">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-687">1.1</span><span class="sxs-lookup"><span data-stu-id="898ce-687">1.1</span></span>|
|[<span data-ttu-id="898ce-688">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-688">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-689">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="898ce-689">ReadWriteItem</span></span>|
|[<span data-ttu-id="898ce-690">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-690">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-691">Создание</span><span class="sxs-lookup"><span data-stu-id="898ce-691">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-692">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-692">Example</span></span>

<span data-ttu-id="898ce-693">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="898ce-693">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="898ce-694">close()</span><span class="sxs-lookup"><span data-stu-id="898ce-694">close()</span></span>

<span data-ttu-id="898ce-695">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="898ce-695">Closes the current item that is being composed.</span></span>

<span data-ttu-id="898ce-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="898ce-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-698">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="898ce-698">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="898ce-699">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="898ce-699">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-700">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-700">Requirements</span></span>

|<span data-ttu-id="898ce-701">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-701">Requirement</span></span>| <span data-ttu-id="898ce-702">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-702">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-703">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="898ce-703">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-704">1.3</span><span class="sxs-lookup"><span data-stu-id="898ce-704">1.3</span></span>|
|[<span data-ttu-id="898ce-705">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-705">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-706">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="898ce-706">Restricted</span></span>|
|[<span data-ttu-id="898ce-707">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-707">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-708">Создание</span><span class="sxs-lookup"><span data-stu-id="898ce-708">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="898ce-709">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="898ce-709">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="898ce-710">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-710">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-711">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="898ce-711">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="898ce-712">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="898ce-712">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="898ce-713">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="898ce-713">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="898ce-714">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="898ce-714">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="898ce-715">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="898ce-715">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="898ce-716">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="898ce-716">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-717">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-717">Parameters</span></span>

|<span data-ttu-id="898ce-718">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-718">Name</span></span>| <span data-ttu-id="898ce-719">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-719">Type</span></span>| <span data-ttu-id="898ce-720">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-720">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="898ce-721">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="898ce-721">String &#124; Object</span></span>| |<span data-ttu-id="898ce-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="898ce-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="898ce-724">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="898ce-724">**OR**</span></span><br/><span data-ttu-id="898ce-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="898ce-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="898ce-727">String</span><span class="sxs-lookup"><span data-stu-id="898ce-727">String</span></span> | <span data-ttu-id="898ce-728">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-728">&lt;optional&gt;</span></span> | <span data-ttu-id="898ce-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="898ce-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="898ce-731">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-731">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="898ce-732">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-732">&lt;optional&gt;</span></span> | <span data-ttu-id="898ce-733">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="898ce-733">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="898ce-734">String</span><span class="sxs-lookup"><span data-stu-id="898ce-734">String</span></span> | | <span data-ttu-id="898ce-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="898ce-737">Строка</span><span class="sxs-lookup"><span data-stu-id="898ce-737">String</span></span> | | <span data-ttu-id="898ce-738">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="898ce-738">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="898ce-739">String</span><span class="sxs-lookup"><span data-stu-id="898ce-739">String</span></span> | | <span data-ttu-id="898ce-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="898ce-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="898ce-742">String</span><span class="sxs-lookup"><span data-stu-id="898ce-742">String</span></span> | | <span data-ttu-id="898ce-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="898ce-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="898ce-746">function</span><span class="sxs-lookup"><span data-stu-id="898ce-746">function</span></span> | <span data-ttu-id="898ce-747">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-747">&lt;optional&gt;</span></span> | <span data-ttu-id="898ce-748">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="898ce-748">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="898ce-749">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-749">Requirements</span></span>

|<span data-ttu-id="898ce-750">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-750">Requirement</span></span>| <span data-ttu-id="898ce-751">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-751">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-752">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-752">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-753">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-753">1.0</span></span>|
|[<span data-ttu-id="898ce-754">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-754">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-755">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-755">ReadItem</span></span>|
|[<span data-ttu-id="898ce-756">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-756">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-757">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-757">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="898ce-758">Примеры</span><span class="sxs-lookup"><span data-stu-id="898ce-758">Examples</span></span>

<span data-ttu-id="898ce-759">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="898ce-759">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="898ce-760">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-760">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="898ce-761">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-761">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="898ce-762">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="898ce-762">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="898ce-763">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="898ce-763">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="898ce-764">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="898ce-764">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="898ce-765">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="898ce-765">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="898ce-766">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-766">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-767">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="898ce-767">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="898ce-768">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="898ce-768">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="898ce-769">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="898ce-769">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="898ce-770">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="898ce-770">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="898ce-771">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="898ce-771">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="898ce-772">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="898ce-772">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-773">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-773">Parameters</span></span>

|<span data-ttu-id="898ce-774">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-774">Name</span></span>| <span data-ttu-id="898ce-775">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-775">Type</span></span>| <span data-ttu-id="898ce-776">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-776">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="898ce-777">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="898ce-777">String &#124; Object</span></span>| | <span data-ttu-id="898ce-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="898ce-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="898ce-780">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="898ce-780">**OR**</span></span><br/><span data-ttu-id="898ce-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="898ce-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="898ce-783">String</span><span class="sxs-lookup"><span data-stu-id="898ce-783">String</span></span> | <span data-ttu-id="898ce-784">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-784">&lt;optional&gt;</span></span> | <span data-ttu-id="898ce-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="898ce-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="898ce-787">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-787">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="898ce-788">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-788">&lt;optional&gt;</span></span> | <span data-ttu-id="898ce-789">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="898ce-789">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="898ce-790">String</span><span class="sxs-lookup"><span data-stu-id="898ce-790">String</span></span> | | <span data-ttu-id="898ce-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="898ce-793">Строка</span><span class="sxs-lookup"><span data-stu-id="898ce-793">String</span></span> | | <span data-ttu-id="898ce-794">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="898ce-794">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="898ce-795">String</span><span class="sxs-lookup"><span data-stu-id="898ce-795">String</span></span> | | <span data-ttu-id="898ce-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="898ce-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="898ce-798">String</span><span class="sxs-lookup"><span data-stu-id="898ce-798">String</span></span> | | <span data-ttu-id="898ce-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="898ce-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="898ce-802">function</span><span class="sxs-lookup"><span data-stu-id="898ce-802">function</span></span> | <span data-ttu-id="898ce-803">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-803">&lt;optional&gt;</span></span> | <span data-ttu-id="898ce-804">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="898ce-804">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="898ce-805">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-805">Requirements</span></span>

|<span data-ttu-id="898ce-806">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-806">Requirement</span></span>| <span data-ttu-id="898ce-807">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-808">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-809">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-809">1.0</span></span>|
|[<span data-ttu-id="898ce-810">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-811">ReadItem</span></span>|
|[<span data-ttu-id="898ce-812">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-813">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-813">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="898ce-814">Примеры</span><span class="sxs-lookup"><span data-stu-id="898ce-814">Examples</span></span>

<span data-ttu-id="898ce-815">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="898ce-815">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="898ce-816">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-816">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="898ce-817">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-817">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="898ce-818">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="898ce-818">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="898ce-819">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="898ce-819">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="898ce-820">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="898ce-820">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-13"></a><span data-ttu-id="898ce-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="898ce-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="898ce-822">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-822">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-823">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="898ce-823">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-824">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-824">Requirements</span></span>

|<span data-ttu-id="898ce-825">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-825">Requirement</span></span>| <span data-ttu-id="898ce-826">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-827">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-828">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-828">1.0</span></span>|
|[<span data-ttu-id="898ce-829">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-830">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-830">ReadItem</span></span>|
|[<span data-ttu-id="898ce-831">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-832">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-832">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="898ce-833">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="898ce-833">Returns:</span></span>

<span data-ttu-id="898ce-834">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="898ce-834">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span></span>

##### <a name="example"></a><span data-ttu-id="898ce-835">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-835">Example</span></span>

<span data-ttu-id="898ce-836">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-836">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="898ce-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="898ce-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="898ce-838">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-838">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-839">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="898ce-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-840">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-840">Parameters</span></span>

|<span data-ttu-id="898ce-841">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-841">Name</span></span>| <span data-ttu-id="898ce-842">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-842">Type</span></span>| <span data-ttu-id="898ce-843">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-843">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="898ce-844">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="898ce-844">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.3)|<span data-ttu-id="898ce-845">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="898ce-845">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="898ce-846">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-846">Requirements</span></span>

|<span data-ttu-id="898ce-847">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-847">Requirement</span></span>| <span data-ttu-id="898ce-848">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-848">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-849">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-849">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-850">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-850">1.0</span></span>|
|[<span data-ttu-id="898ce-851">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-851">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-852">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="898ce-852">Restricted</span></span>|
|[<span data-ttu-id="898ce-853">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-853">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-854">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-854">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="898ce-855">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="898ce-855">Returns:</span></span>

<span data-ttu-id="898ce-856">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="898ce-856">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="898ce-857">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="898ce-857">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="898ce-858">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="898ce-858">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="898ce-859">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="898ce-859">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="898ce-860">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="898ce-860">Value of `entityType`</span></span> | <span data-ttu-id="898ce-861">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="898ce-861">Type of objects in returned array</span></span> | <span data-ttu-id="898ce-862">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-862">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="898ce-863">String</span><span class="sxs-lookup"><span data-stu-id="898ce-863">String</span></span> | <span data-ttu-id="898ce-864">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="898ce-864">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="898ce-865">Contact</span><span class="sxs-lookup"><span data-stu-id="898ce-865">Contact</span></span> | <span data-ttu-id="898ce-866">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="898ce-866">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="898ce-867">String</span><span class="sxs-lookup"><span data-stu-id="898ce-867">String</span></span> | <span data-ttu-id="898ce-868">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="898ce-868">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="898ce-869">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="898ce-869">MeetingSuggestion</span></span> | <span data-ttu-id="898ce-870">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="898ce-870">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="898ce-871">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="898ce-871">PhoneNumber</span></span> | <span data-ttu-id="898ce-872">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="898ce-872">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="898ce-873">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="898ce-873">TaskSuggestion</span></span> | <span data-ttu-id="898ce-874">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="898ce-874">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="898ce-875">String</span><span class="sxs-lookup"><span data-stu-id="898ce-875">String</span></span> | <span data-ttu-id="898ce-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="898ce-876">**Restricted**</span></span> |

<span data-ttu-id="898ce-877">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="898ce-877">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

##### <a name="example"></a><span data-ttu-id="898ce-878">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-878">Example</span></span>

<span data-ttu-id="898ce-879">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-879">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="898ce-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="898ce-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="898ce-881">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="898ce-881">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-882">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="898ce-882">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="898ce-883">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="898ce-883">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-884">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-884">Parameters</span></span>

|<span data-ttu-id="898ce-885">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-885">Name</span></span>| <span data-ttu-id="898ce-886">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-886">Type</span></span>| <span data-ttu-id="898ce-887">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="898ce-888">String</span><span class="sxs-lookup"><span data-stu-id="898ce-888">String</span></span>|<span data-ttu-id="898ce-889">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="898ce-889">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="898ce-890">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-890">Requirements</span></span>

|<span data-ttu-id="898ce-891">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-891">Requirement</span></span>| <span data-ttu-id="898ce-892">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-893">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-894">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-894">1.0</span></span>|
|[<span data-ttu-id="898ce-895">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-896">ReadItem</span></span>|
|[<span data-ttu-id="898ce-897">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-898">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="898ce-899">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="898ce-899">Returns:</span></span>

<span data-ttu-id="898ce-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="898ce-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="898ce-902">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="898ce-902">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="898ce-903">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="898ce-903">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="898ce-904">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="898ce-904">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-905">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="898ce-905">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="898ce-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="898ce-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="898ce-909">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="898ce-909">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="898ce-910">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="898ce-910">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="898ce-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="898ce-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="898ce-914">Requirements</span><span class="sxs-lookup"><span data-stu-id="898ce-914">Requirements</span></span>

|<span data-ttu-id="898ce-915">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-915">Requirement</span></span>| <span data-ttu-id="898ce-916">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-917">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-918">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-918">1.0</span></span>|
|[<span data-ttu-id="898ce-919">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-919">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-920">ReadItem</span></span>|
|[<span data-ttu-id="898ce-921">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-921">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-922">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-922">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="898ce-923">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="898ce-923">Returns:</span></span>

<span data-ttu-id="898ce-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="898ce-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="898ce-926">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="898ce-926">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="898ce-927">Object</span><span class="sxs-lookup"><span data-stu-id="898ce-927">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="898ce-928">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-928">Example</span></span>

<span data-ttu-id="898ce-929">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="898ce-929">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="898ce-930">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="898ce-930">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="898ce-931">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="898ce-931">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-932">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="898ce-932">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="898ce-933">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="898ce-933">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="898ce-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="898ce-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-936">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-936">Parameters</span></span>

|<span data-ttu-id="898ce-937">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-937">Name</span></span>| <span data-ttu-id="898ce-938">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-938">Type</span></span>| <span data-ttu-id="898ce-939">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-939">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="898ce-940">String</span><span class="sxs-lookup"><span data-stu-id="898ce-940">String</span></span>|<span data-ttu-id="898ce-941">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="898ce-941">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="898ce-942">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-942">Requirements</span></span>

|<span data-ttu-id="898ce-943">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-943">Requirement</span></span>| <span data-ttu-id="898ce-944">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-944">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-945">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-945">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-946">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-946">1.0</span></span>|
|[<span data-ttu-id="898ce-947">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-947">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-948">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-948">ReadItem</span></span>|
|[<span data-ttu-id="898ce-949">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-949">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-950">Чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-950">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="898ce-951">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="898ce-951">Returns:</span></span>

<span data-ttu-id="898ce-952">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="898ce-952">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="898ce-953">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="898ce-953">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="898ce-954">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="898ce-954">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="898ce-955">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-955">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="898ce-956">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="898ce-956">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="898ce-957">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-957">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="898ce-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="898ce-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-960">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-960">Parameters</span></span>

|<span data-ttu-id="898ce-961">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-961">Name</span></span>| <span data-ttu-id="898ce-962">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-962">Type</span></span>| <span data-ttu-id="898ce-963">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="898ce-963">Attributes</span></span>| <span data-ttu-id="898ce-964">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-964">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="898ce-965">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="898ce-965">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="898ce-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="898ce-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="898ce-969">Объект</span><span class="sxs-lookup"><span data-stu-id="898ce-969">Object</span></span>| <span data-ttu-id="898ce-970">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-970">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-971">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="898ce-971">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="898ce-972">Объект</span><span class="sxs-lookup"><span data-stu-id="898ce-972">Object</span></span>| <span data-ttu-id="898ce-973">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-973">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-974">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="898ce-974">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="898ce-975">функция</span><span class="sxs-lookup"><span data-stu-id="898ce-975">function</span></span>||<span data-ttu-id="898ce-976">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="898ce-976">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="898ce-977">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="898ce-977">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="898ce-978">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="898ce-978">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="898ce-979">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-979">Requirements</span></span>

|<span data-ttu-id="898ce-980">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-980">Requirement</span></span>| <span data-ttu-id="898ce-981">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-981">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-982">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="898ce-982">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-983">1.2</span><span class="sxs-lookup"><span data-stu-id="898ce-983">1.2</span></span>|
|[<span data-ttu-id="898ce-984">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-984">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-985">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="898ce-985">ReadWriteItem</span></span>|
|[<span data-ttu-id="898ce-986">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-986">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-987">Создание</span><span class="sxs-lookup"><span data-stu-id="898ce-987">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="898ce-988">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="898ce-988">Returns:</span></span>

<span data-ttu-id="898ce-989">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="898ce-989">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="898ce-990">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="898ce-990">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="898ce-991">String</span><span class="sxs-lookup"><span data-stu-id="898ce-991">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="898ce-992">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-992">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="898ce-993">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="898ce-993">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="898ce-994">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-994">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="898ce-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="898ce-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-998">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-998">Parameters</span></span>

|<span data-ttu-id="898ce-999">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-999">Name</span></span>| <span data-ttu-id="898ce-1000">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-1000">Type</span></span>| <span data-ttu-id="898ce-1001">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="898ce-1001">Attributes</span></span>| <span data-ttu-id="898ce-1002">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-1002">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="898ce-1003">function</span><span class="sxs-lookup"><span data-stu-id="898ce-1003">function</span></span>||<span data-ttu-id="898ce-1004">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="898ce-1004">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="898ce-1005">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="898ce-1005">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="898ce-1006">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="898ce-1006">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="898ce-1007">Объект</span><span class="sxs-lookup"><span data-stu-id="898ce-1007">Object</span></span>| <span data-ttu-id="898ce-1008">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-1009">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="898ce-1009">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="898ce-1010">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="898ce-1010">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="898ce-1011">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-1011">Requirements</span></span>

|<span data-ttu-id="898ce-1012">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-1012">Requirement</span></span>| <span data-ttu-id="898ce-1013">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-1014">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="898ce-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-1015">1.0</span><span class="sxs-lookup"><span data-stu-id="898ce-1015">1.0</span></span>|
|[<span data-ttu-id="898ce-1016">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-1017">ReadItem</span><span class="sxs-lookup"><span data-stu-id="898ce-1017">ReadItem</span></span>|
|[<span data-ttu-id="898ce-1018">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-1019">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="898ce-1019">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-1020">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-1020">Example</span></span>

<span data-ttu-id="898ce-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="898ce-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="898ce-1024">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="898ce-1024">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="898ce-1025">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="898ce-1025">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="898ce-1026">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="898ce-1026">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="898ce-1027">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="898ce-1027">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="898ce-1028">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="898ce-1028">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="898ce-1029">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="898ce-1029">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-1030">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-1030">Parameters</span></span>

|<span data-ttu-id="898ce-1031">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-1031">Name</span></span>| <span data-ttu-id="898ce-1032">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-1032">Type</span></span>| <span data-ttu-id="898ce-1033">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="898ce-1033">Attributes</span></span>| <span data-ttu-id="898ce-1034">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-1034">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="898ce-1035">String</span><span class="sxs-lookup"><span data-stu-id="898ce-1035">String</span></span>||<span data-ttu-id="898ce-1036">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="898ce-1036">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="898ce-1037">Object</span><span class="sxs-lookup"><span data-stu-id="898ce-1037">Object</span></span>| <span data-ttu-id="898ce-1038">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-1039">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="898ce-1039">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="898ce-1040">Объект</span><span class="sxs-lookup"><span data-stu-id="898ce-1040">Object</span></span>| <span data-ttu-id="898ce-1041">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-1042">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="898ce-1042">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="898ce-1043">функция</span><span class="sxs-lookup"><span data-stu-id="898ce-1043">function</span></span>| <span data-ttu-id="898ce-1044">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-1044">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-1045">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="898ce-1045">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="898ce-1046">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="898ce-1046">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="898ce-1047">Ошибки</span><span class="sxs-lookup"><span data-stu-id="898ce-1047">Errors</span></span>

| <span data-ttu-id="898ce-1048">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="898ce-1048">Error code</span></span> | <span data-ttu-id="898ce-1049">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-1049">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="898ce-1050">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="898ce-1050">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="898ce-1051">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-1051">Requirements</span></span>

|<span data-ttu-id="898ce-1052">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-1052">Requirement</span></span>| <span data-ttu-id="898ce-1053">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-1054">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="898ce-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-1055">1.1</span><span class="sxs-lookup"><span data-stu-id="898ce-1055">1.1</span></span>|
|[<span data-ttu-id="898ce-1056">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-1056">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-1057">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="898ce-1057">ReadWriteItem</span></span>|
|[<span data-ttu-id="898ce-1058">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-1058">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-1059">Создание</span><span class="sxs-lookup"><span data-stu-id="898ce-1059">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-1060">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-1060">Example</span></span>

<span data-ttu-id="898ce-1061">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="898ce-1061">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="898ce-1062">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="898ce-1062">saveAsync([options], callback)</span></span>

<span data-ttu-id="898ce-1063">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="898ce-1063">Asynchronously saves an item.</span></span>

<span data-ttu-id="898ce-1064">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="898ce-1064">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="898ce-1065">В Outlook в Интернете или Outlook в интерактивном режиме элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="898ce-1065">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="898ce-1066">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="898ce-1066">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-1067">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="898ce-1067">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="898ce-1068">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="898ce-1068">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="898ce-p168">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="898ce-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="898ce-1072">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="898ce-1072">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="898ce-1073">Outlook в Mac не поддерживает сохранение собраний.</span><span class="sxs-lookup"><span data-stu-id="898ce-1073">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="898ce-1074">`saveAsync` Метод завершается с ошибкой при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="898ce-1074">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="898ce-1075">Просмотреть [не удается сохранить собрание в виде черновика в Outlook для Mac с помощью API Office JS](https://support.microsoft.com/help/4505745) для обхода.</span><span class="sxs-lookup"><span data-stu-id="898ce-1075">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="898ce-1076">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="898ce-1076">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-1077">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-1077">Parameters</span></span>

|<span data-ttu-id="898ce-1078">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-1078">Name</span></span>| <span data-ttu-id="898ce-1079">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-1079">Type</span></span>| <span data-ttu-id="898ce-1080">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="898ce-1080">Attributes</span></span>| <span data-ttu-id="898ce-1081">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-1081">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="898ce-1082">Объект</span><span class="sxs-lookup"><span data-stu-id="898ce-1082">Object</span></span>| <span data-ttu-id="898ce-1083">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-1084">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="898ce-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="898ce-1085">Объект</span><span class="sxs-lookup"><span data-stu-id="898ce-1085">Object</span></span>| <span data-ttu-id="898ce-1086">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-1087">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="898ce-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="898ce-1088">функция</span><span class="sxs-lookup"><span data-stu-id="898ce-1088">function</span></span>||<span data-ttu-id="898ce-1089">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="898ce-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="898ce-1090">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="898ce-1090">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="898ce-1091">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-1091">Requirements</span></span>

|<span data-ttu-id="898ce-1092">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-1092">Requirement</span></span>| <span data-ttu-id="898ce-1093">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-1093">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-1094">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="898ce-1094">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-1095">1.3</span><span class="sxs-lookup"><span data-stu-id="898ce-1095">1.3</span></span>|
|[<span data-ttu-id="898ce-1096">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-1096">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-1097">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="898ce-1097">ReadWriteItem</span></span>|
|[<span data-ttu-id="898ce-1098">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-1098">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-1099">Создание</span><span class="sxs-lookup"><span data-stu-id="898ce-1099">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="898ce-1100">Примеры</span><span class="sxs-lookup"><span data-stu-id="898ce-1100">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="898ce-p170">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="898ce-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="898ce-1103">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="898ce-1103">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="898ce-1104">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="898ce-1104">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="898ce-p171">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="898ce-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="898ce-1108">Параметры</span><span class="sxs-lookup"><span data-stu-id="898ce-1108">Parameters</span></span>

|<span data-ttu-id="898ce-1109">Имя</span><span class="sxs-lookup"><span data-stu-id="898ce-1109">Name</span></span>| <span data-ttu-id="898ce-1110">Тип</span><span class="sxs-lookup"><span data-stu-id="898ce-1110">Type</span></span>| <span data-ttu-id="898ce-1111">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="898ce-1111">Attributes</span></span>| <span data-ttu-id="898ce-1112">Описание</span><span class="sxs-lookup"><span data-stu-id="898ce-1112">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="898ce-1113">String</span><span class="sxs-lookup"><span data-stu-id="898ce-1113">String</span></span>||<span data-ttu-id="898ce-p172">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="898ce-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="898ce-1117">Object</span><span class="sxs-lookup"><span data-stu-id="898ce-1117">Object</span></span>| <span data-ttu-id="898ce-1118">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-1118">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-1119">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="898ce-1119">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="898ce-1120">Объект</span><span class="sxs-lookup"><span data-stu-id="898ce-1120">Object</span></span>| <span data-ttu-id="898ce-1121">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-1121">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-1122">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="898ce-1122">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="898ce-1123">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="898ce-1123">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="898ce-1124">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="898ce-1124">&lt;optional&gt;</span></span>|<span data-ttu-id="898ce-1125">Если `text`текущий стиль применяется в Outlook для веб-клиентов и клиентов для настольных ПК.</span><span class="sxs-lookup"><span data-stu-id="898ce-1125">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="898ce-1126">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="898ce-1126">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="898ce-1127">Если `html` и поле поддерживает HTML (тема не используется), текущий стиль применяется в Outlook в Интернете, а в настольных клиентах Outlook применяется стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="898ce-1127">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="898ce-1128">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="898ce-1128">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="898ce-1129">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="898ce-1129">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="898ce-1130">функция</span><span class="sxs-lookup"><span data-stu-id="898ce-1130">function</span></span>||<span data-ttu-id="898ce-1131">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="898ce-1131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="898ce-1132">Требования</span><span class="sxs-lookup"><span data-stu-id="898ce-1132">Requirements</span></span>

|<span data-ttu-id="898ce-1133">Требование</span><span class="sxs-lookup"><span data-stu-id="898ce-1133">Requirement</span></span>| <span data-ttu-id="898ce-1134">Значение</span><span class="sxs-lookup"><span data-stu-id="898ce-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="898ce-1135">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="898ce-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="898ce-1136">1.2</span><span class="sxs-lookup"><span data-stu-id="898ce-1136">1.2</span></span>|
|[<span data-ttu-id="898ce-1137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="898ce-1137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="898ce-1138">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="898ce-1138">ReadWriteItem</span></span>|
|[<span data-ttu-id="898ce-1139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="898ce-1139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="898ce-1140">Создание</span><span class="sxs-lookup"><span data-stu-id="898ce-1140">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="898ce-1141">Пример</span><span class="sxs-lookup"><span data-stu-id="898ce-1141">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
