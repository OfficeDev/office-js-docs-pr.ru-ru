---
title: Office. Context. Mailbox. Item — набор требований 1,3
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: e2e91dc196e0c67eed3a358e9f0d864885a01945
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682650"
---
# <a name="item"></a><span data-ttu-id="e971d-102">item</span><span class="sxs-lookup"><span data-stu-id="e971d-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="e971d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="e971d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="e971d-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="e971d-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="e971d-106">Requirements</span></span>

|<span data-ttu-id="e971d-107">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-107">Requirement</span></span>| <span data-ttu-id="e971d-108">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-110">1.0</span></span>|
|[<span data-ttu-id="e971d-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e971d-112">Restricted</span></span>|
|[<span data-ttu-id="e971d-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e971d-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="e971d-115">Members and methods</span></span>

| <span data-ttu-id="e971d-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-116">Member</span></span> | <span data-ttu-id="e971d-117">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e971d-118">attachments</span><span class="sxs-lookup"><span data-stu-id="e971d-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="e971d-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-119">Member</span></span> |
| [<span data-ttu-id="e971d-120">bcc</span><span class="sxs-lookup"><span data-stu-id="e971d-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="e971d-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-121">Member</span></span> |
| [<span data-ttu-id="e971d-122">body</span><span class="sxs-lookup"><span data-stu-id="e971d-122">body</span></span>](#body-body) | <span data-ttu-id="e971d-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-123">Member</span></span> |
| [<span data-ttu-id="e971d-124">cc</span><span class="sxs-lookup"><span data-stu-id="e971d-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e971d-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-125">Member</span></span> |
| [<span data-ttu-id="e971d-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="e971d-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="e971d-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-127">Member</span></span> |
| [<span data-ttu-id="e971d-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="e971d-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="e971d-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-129">Member</span></span> |
| [<span data-ttu-id="e971d-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="e971d-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="e971d-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-131">Member</span></span> |
| [<span data-ttu-id="e971d-132">end</span><span class="sxs-lookup"><span data-stu-id="e971d-132">end</span></span>](#end-datetime) | <span data-ttu-id="e971d-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-133">Member</span></span> |
| [<span data-ttu-id="e971d-134">from</span><span class="sxs-lookup"><span data-stu-id="e971d-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="e971d-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-135">Member</span></span> |
| [<span data-ttu-id="e971d-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="e971d-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="e971d-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-137">Member</span></span> |
| [<span data-ttu-id="e971d-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="e971d-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="e971d-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-139">Member</span></span> |
| [<span data-ttu-id="e971d-140">itemId</span><span class="sxs-lookup"><span data-stu-id="e971d-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="e971d-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-141">Member</span></span> |
| [<span data-ttu-id="e971d-142">itemType</span><span class="sxs-lookup"><span data-stu-id="e971d-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="e971d-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-143">Member</span></span> |
| [<span data-ttu-id="e971d-144">location</span><span class="sxs-lookup"><span data-stu-id="e971d-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="e971d-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-145">Member</span></span> |
| [<span data-ttu-id="e971d-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="e971d-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="e971d-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-147">Member</span></span> |
| [<span data-ttu-id="e971d-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="e971d-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="e971d-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-149">Member</span></span> |
| [<span data-ttu-id="e971d-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="e971d-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e971d-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-151">Member</span></span> |
| [<span data-ttu-id="e971d-152">organizer</span><span class="sxs-lookup"><span data-stu-id="e971d-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="e971d-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-153">Member</span></span> |
| [<span data-ttu-id="e971d-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="e971d-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e971d-155">Member</span><span class="sxs-lookup"><span data-stu-id="e971d-155">Member</span></span> |
| [<span data-ttu-id="e971d-156">sender</span><span class="sxs-lookup"><span data-stu-id="e971d-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="e971d-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-157">Member</span></span> |
| [<span data-ttu-id="e971d-158">start</span><span class="sxs-lookup"><span data-stu-id="e971d-158">start</span></span>](#start-datetime) | <span data-ttu-id="e971d-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-159">Member</span></span> |
| [<span data-ttu-id="e971d-160">subject</span><span class="sxs-lookup"><span data-stu-id="e971d-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="e971d-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-161">Member</span></span> |
| [<span data-ttu-id="e971d-162">to</span><span class="sxs-lookup"><span data-stu-id="e971d-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e971d-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="e971d-163">Member</span></span> |
| [<span data-ttu-id="e971d-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e971d-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="e971d-165">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-165">Method</span></span> |
| [<span data-ttu-id="e971d-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e971d-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="e971d-167">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-167">Method</span></span> |
| [<span data-ttu-id="e971d-168">close</span><span class="sxs-lookup"><span data-stu-id="e971d-168">close</span></span>](#close) | <span data-ttu-id="e971d-169">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-169">Method</span></span> |
| [<span data-ttu-id="e971d-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="e971d-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="e971d-171">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-171">Method</span></span> |
| [<span data-ttu-id="e971d-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="e971d-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="e971d-173">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-173">Method</span></span> |
| [<span data-ttu-id="e971d-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="e971d-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="e971d-175">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-175">Method</span></span> |
| [<span data-ttu-id="e971d-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="e971d-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="e971d-177">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-177">Method</span></span> |
| [<span data-ttu-id="e971d-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="e971d-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="e971d-179">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-179">Method</span></span> |
| [<span data-ttu-id="e971d-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e971d-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="e971d-181">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-181">Method</span></span> |
| [<span data-ttu-id="e971d-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="e971d-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="e971d-183">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-183">Method</span></span> |
| [<span data-ttu-id="e971d-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e971d-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="e971d-185">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-185">Method</span></span> |
| [<span data-ttu-id="e971d-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e971d-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="e971d-187">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-187">Method</span></span> |
| [<span data-ttu-id="e971d-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e971d-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="e971d-189">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-189">Method</span></span> |
| [<span data-ttu-id="e971d-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="e971d-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="e971d-191">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-191">Method</span></span> |
| [<span data-ttu-id="e971d-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e971d-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="e971d-193">Метод</span><span class="sxs-lookup"><span data-stu-id="e971d-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="e971d-194">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-194">Example</span></span>

<span data-ttu-id="e971d-195">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="e971d-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="e971d-196">Members</span><span class="sxs-lookup"><span data-stu-id="e971d-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-13"></a><span data-ttu-id="e971d-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="e971d-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

<span data-ttu-id="e971d-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-200">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="e971d-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="e971d-201">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="e971d-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-202">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-202">Type</span></span>

*   <span data-ttu-id="e971d-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span><span class="sxs-lookup"><span data-stu-id="e971d-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.3)></span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-204">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-204">Requirements</span></span>

|<span data-ttu-id="e971d-205">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-205">Requirement</span></span>| <span data-ttu-id="e971d-206">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-208">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-208">1.0</span></span>|
|[<span data-ttu-id="e971d-209">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-210">ReadItem</span></span>|
|[<span data-ttu-id="e971d-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-212">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-213">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-213">Example</span></span>

<span data-ttu-id="e971d-214">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="e971d-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-216">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="e971d-217">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e971d-217">Compose mode only.</span></span>

<span data-ttu-id="e971d-218">По умолчанию коллекция ограничена максимум 100 членами.</span><span class="sxs-lookup"><span data-stu-id="e971d-218">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e971d-219">Однако в Windows и Mac применяются следующие пределы.</span><span class="sxs-lookup"><span data-stu-id="e971d-219">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e971d-220">Максимальное число участников для получения 500.</span><span class="sxs-lookup"><span data-stu-id="e971d-220">Get 500 members maximum.</span></span>
- <span data-ttu-id="e971d-221">Задайте не более 100 членов для каждого вызова, до 500 всего членов.</span><span class="sxs-lookup"><span data-stu-id="e971d-221">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-222">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-222">Type</span></span>

*   [<span data-ttu-id="e971d-223">Получатели</span><span class="sxs-lookup"><span data-stu-id="e971d-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="e971d-224">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-224">Requirements</span></span>

|<span data-ttu-id="e971d-225">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-225">Requirement</span></span>| <span data-ttu-id="e971d-226">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-227">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-228">1.1</span><span class="sxs-lookup"><span data-stu-id="e971d-228">1.1</span></span>|
|[<span data-ttu-id="e971d-229">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-230">ReadItem</span></span>|
|[<span data-ttu-id="e971d-231">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-232">Создание</span><span class="sxs-lookup"><span data-stu-id="e971d-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-233">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-233">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-13"></a><span data-ttu-id="e971d-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-235">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-236">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-236">Type</span></span>

*   [<span data-ttu-id="e971d-237">Body</span><span class="sxs-lookup"><span data-stu-id="e971d-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="e971d-238">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-238">Requirements</span></span>

|<span data-ttu-id="e971d-239">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-239">Requirement</span></span>| <span data-ttu-id="e971d-240">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-241">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-242">1.1</span><span class="sxs-lookup"><span data-stu-id="e971d-242">1.1</span></span>|
|[<span data-ttu-id="e971d-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-244">ReadItem</span></span>|
|[<span data-ttu-id="e971d-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-247">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-247">Example</span></span>

<span data-ttu-id="e971d-248">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="e971d-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="e971d-249">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-249">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="e971d-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-251">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="e971d-252">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e971d-253">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e971d-253">Read mode</span></span>

<span data-ttu-id="e971d-254">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-254">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="e971d-255">По умолчанию коллекция ограничена максимум 100 членами.</span><span class="sxs-lookup"><span data-stu-id="e971d-255">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e971d-256">Однако в Windows и Mac вы можете получить максимум 500 членов.</span><span class="sxs-lookup"><span data-stu-id="e971d-256">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="e971d-257">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e971d-257">Compose mode</span></span>

<span data-ttu-id="e971d-258">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-258">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="e971d-259">По умолчанию коллекция ограничена максимум 100 членами.</span><span class="sxs-lookup"><span data-stu-id="e971d-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e971d-260">Однако в Windows и Mac применяются следующие пределы.</span><span class="sxs-lookup"><span data-stu-id="e971d-260">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e971d-261">Максимальное число участников для получения 500.</span><span class="sxs-lookup"><span data-stu-id="e971d-261">Get 500 members maximum.</span></span>
- <span data-ttu-id="e971d-262">Задайте не более 100 членов для каждого вызова, до 500 всего членов.</span><span class="sxs-lookup"><span data-stu-id="e971d-262">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

<br>

---
---

##### <a name="type"></a><span data-ttu-id="e971d-263">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-263">Type</span></span>

*   <span data-ttu-id="e971d-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-265">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-265">Requirements</span></span>

|<span data-ttu-id="e971d-266">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-266">Requirement</span></span>| <span data-ttu-id="e971d-267">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-268">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-269">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-269">1.0</span></span>|
|[<span data-ttu-id="e971d-270">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-271">ReadItem</span></span>|
|[<span data-ttu-id="e971d-272">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-273">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="e971d-274">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="e971d-274">(nullable) conversationId: String</span></span>

<span data-ttu-id="e971d-275">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="e971d-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="e971d-p109">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="e971d-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="e971d-p110">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="e971d-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-280">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-280">Type</span></span>

*   <span data-ttu-id="e971d-281">String</span><span class="sxs-lookup"><span data-stu-id="e971d-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-282">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-282">Requirements</span></span>

|<span data-ttu-id="e971d-283">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-283">Requirement</span></span>| <span data-ttu-id="e971d-284">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-285">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-286">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-286">1.0</span></span>|
|[<span data-ttu-id="e971d-287">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-288">ReadItem</span></span>|
|[<span data-ttu-id="e971d-289">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-290">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-291">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-291">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="e971d-292">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="e971d-292">dateTimeCreated: Date</span></span>

<span data-ttu-id="e971d-p111">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-295">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-295">Type</span></span>

*   <span data-ttu-id="e971d-296">Дата</span><span class="sxs-lookup"><span data-stu-id="e971d-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-297">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-297">Requirements</span></span>

|<span data-ttu-id="e971d-298">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-298">Requirement</span></span>| <span data-ttu-id="e971d-299">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-300">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-301">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-301">1.0</span></span>|
|[<span data-ttu-id="e971d-302">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-302">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-303">ReadItem</span></span>|
|[<span data-ttu-id="e971d-304">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-304">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-305">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-306">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-306">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="e971d-307">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="e971d-307">dateTimeModified: Date</span></span>

<span data-ttu-id="e971d-p112">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-310">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e971d-310">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-311">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-311">Type</span></span>

*   <span data-ttu-id="e971d-312">Дата</span><span class="sxs-lookup"><span data-stu-id="e971d-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-313">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-313">Requirements</span></span>

|<span data-ttu-id="e971d-314">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-314">Requirement</span></span>| <span data-ttu-id="e971d-315">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-316">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-317">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-317">1.0</span></span>|
|[<span data-ttu-id="e971d-318">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-319">ReadItem</span></span>|
|[<span data-ttu-id="e971d-320">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-321">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-322">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-322">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="e971d-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-324">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="e971d-p113">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="e971d-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e971d-327">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e971d-327">Read mode</span></span>

<span data-ttu-id="e971d-328">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="e971d-328">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="e971d-329">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e971d-329">Compose mode</span></span>

<span data-ttu-id="e971d-330">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="e971d-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="e971d-331">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="e971d-331">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="e971d-332">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-332">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e971d-333">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-333">Type</span></span>

*   <span data-ttu-id="e971d-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-335">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-335">Requirements</span></span>

|<span data-ttu-id="e971d-336">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-336">Requirement</span></span>| <span data-ttu-id="e971d-337">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-338">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-339">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-339">1.0</span></span>|
|[<span data-ttu-id="e971d-340">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-341">ReadItem</span></span>|
|[<span data-ttu-id="e971d-342">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-343">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-343">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="e971d-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-p114">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="e971d-p115">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="e971d-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-349">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e971d-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-350">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-350">Type</span></span>

*   [<span data-ttu-id="e971d-351">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e971d-351">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="e971d-352">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-352">Requirements</span></span>

|<span data-ttu-id="e971d-353">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-353">Requirement</span></span>| <span data-ttu-id="e971d-354">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-355">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-356">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-356">1.0</span></span>|
|[<span data-ttu-id="e971d-357">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-358">ReadItem</span></span>|
|[<span data-ttu-id="e971d-359">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-360">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-360">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-361">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-361">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="e971d-362">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="e971d-362">internetMessageId: String</span></span>

<span data-ttu-id="e971d-p116">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-365">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-365">Type</span></span>

*   <span data-ttu-id="e971d-366">String</span><span class="sxs-lookup"><span data-stu-id="e971d-366">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-367">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-367">Requirements</span></span>

|<span data-ttu-id="e971d-368">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-368">Requirement</span></span>| <span data-ttu-id="e971d-369">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-370">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-371">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-371">1.0</span></span>|
|[<span data-ttu-id="e971d-372">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-373">ReadItem</span></span>|
|[<span data-ttu-id="e971d-374">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-374">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-375">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-375">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-376">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-376">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="e971d-377">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="e971d-377">itemClass: String</span></span>

<span data-ttu-id="e971d-p117">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="e971d-p118">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="e971d-382">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-382">Type</span></span> | <span data-ttu-id="e971d-383">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-383">Description</span></span> | <span data-ttu-id="e971d-384">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="e971d-384">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="e971d-385">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="e971d-385">Appointment items</span></span> | <span data-ttu-id="e971d-386">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="e971d-386">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="e971d-387">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="e971d-387">Message items</span></span> | <span data-ttu-id="e971d-388">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-388">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="e971d-389">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="e971d-389">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-390">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-390">Type</span></span>

*   <span data-ttu-id="e971d-391">String</span><span class="sxs-lookup"><span data-stu-id="e971d-391">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-392">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-392">Requirements</span></span>

|<span data-ttu-id="e971d-393">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-393">Requirement</span></span>| <span data-ttu-id="e971d-394">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-395">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-396">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-396">1.0</span></span>|
|[<span data-ttu-id="e971d-397">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-397">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-398">ReadItem</span></span>|
|[<span data-ttu-id="e971d-399">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-400">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-400">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-401">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-401">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="e971d-402">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="e971d-402">(nullable) itemId: String</span></span>

<span data-ttu-id="e971d-p119">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-405">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="e971d-405">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e971d-406">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="e971d-406">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="e971d-407">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="e971d-407">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e971d-408">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="e971d-408">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="e971d-p121">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-411">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-411">Type</span></span>

*   <span data-ttu-id="e971d-412">String</span><span class="sxs-lookup"><span data-stu-id="e971d-412">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-413">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-413">Requirements</span></span>

|<span data-ttu-id="e971d-414">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-414">Requirement</span></span>| <span data-ttu-id="e971d-415">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-415">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-416">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-417">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-417">1.0</span></span>|
|[<span data-ttu-id="e971d-418">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-419">ReadItem</span></span>|
|[<span data-ttu-id="e971d-420">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-421">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-421">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-422">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-422">Example</span></span>

<span data-ttu-id="e971d-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-13"></a><span data-ttu-id="e971d-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-426">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="e971d-426">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="e971d-427">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="e971d-427">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-428">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-428">Type</span></span>

*   [<span data-ttu-id="e971d-429">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="e971d-429">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="e971d-430">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-430">Requirements</span></span>

|<span data-ttu-id="e971d-431">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-431">Requirement</span></span>| <span data-ttu-id="e971d-432">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-432">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-433">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-433">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-434">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-434">1.0</span></span>|
|[<span data-ttu-id="e971d-435">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-435">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-436">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-436">ReadItem</span></span>|
|[<span data-ttu-id="e971d-437">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-437">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-438">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-438">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-439">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-439">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-13"></a><span data-ttu-id="e971d-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-441">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-441">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e971d-442">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e971d-442">Read mode</span></span>

<span data-ttu-id="e971d-443">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-443">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="e971d-444">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e971d-444">Compose mode</span></span>

<span data-ttu-id="e971d-445">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-445">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e971d-446">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-446">Type</span></span>

*   <span data-ttu-id="e971d-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-448">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-448">Requirements</span></span>

|<span data-ttu-id="e971d-449">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-449">Requirement</span></span>| <span data-ttu-id="e971d-450">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-451">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-452">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-452">1.0</span></span>|
|[<span data-ttu-id="e971d-453">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-454">ReadItem</span></span>|
|[<span data-ttu-id="e971d-455">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-456">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-456">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="e971d-457">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="e971d-457">normalizedSubject: String</span></span>

<span data-ttu-id="e971d-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="e971d-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="e971d-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-462">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-462">Type</span></span>

*   <span data-ttu-id="e971d-463">String</span><span class="sxs-lookup"><span data-stu-id="e971d-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-464">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-464">Requirements</span></span>

|<span data-ttu-id="e971d-465">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-465">Requirement</span></span>| <span data-ttu-id="e971d-466">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-467">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-468">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-468">1.0</span></span>|
|[<span data-ttu-id="e971d-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-470">ReadItem</span></span>|
|[<span data-ttu-id="e971d-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-472">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-473">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-473">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-13"></a><span data-ttu-id="e971d-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-475">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-475">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-476">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-476">Type</span></span>

*   [<span data-ttu-id="e971d-477">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="e971d-477">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="e971d-478">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-478">Requirements</span></span>

|<span data-ttu-id="e971d-479">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-479">Requirement</span></span>| <span data-ttu-id="e971d-480">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-481">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-482">1.3</span><span class="sxs-lookup"><span data-stu-id="e971d-482">1.3</span></span>|
|[<span data-ttu-id="e971d-483">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-484">ReadItem</span></span>|
|[<span data-ttu-id="e971d-485">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-486">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-486">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-487">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-487">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="e971d-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-489">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="e971d-489">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="e971d-490">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-490">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e971d-491">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e971d-491">Read mode</span></span>

<span data-ttu-id="e971d-492">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="e971d-492">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="e971d-493">По умолчанию коллекция ограничена максимум 100 членами.</span><span class="sxs-lookup"><span data-stu-id="e971d-493">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e971d-494">Однако в Windows и Mac вы можете получить максимум 500 членов.</span><span class="sxs-lookup"><span data-stu-id="e971d-494">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="e971d-495">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e971d-495">Compose mode</span></span>

<span data-ttu-id="e971d-496">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="e971d-496">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="e971d-497">По умолчанию коллекция ограничена максимум 100 членами.</span><span class="sxs-lookup"><span data-stu-id="e971d-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e971d-498">Однако в Windows и Mac применяются следующие пределы.</span><span class="sxs-lookup"><span data-stu-id="e971d-498">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e971d-499">Максимальное число участников для получения 500.</span><span class="sxs-lookup"><span data-stu-id="e971d-499">Get 500 members maximum.</span></span>
- <span data-ttu-id="e971d-500">Задайте не более 100 членов для каждого вызова, до 500 всего членов.</span><span class="sxs-lookup"><span data-stu-id="e971d-500">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e971d-501">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-501">Type</span></span>

*   <span data-ttu-id="e971d-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-503">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-503">Requirements</span></span>

|<span data-ttu-id="e971d-504">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-504">Requirement</span></span>| <span data-ttu-id="e971d-505">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-506">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-507">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-507">1.0</span></span>|
|[<span data-ttu-id="e971d-508">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-509">ReadItem</span></span>|
|[<span data-ttu-id="e971d-510">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-511">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-511">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="e971d-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-p128">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-515">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-515">Type</span></span>

*   [<span data-ttu-id="e971d-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e971d-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="e971d-517">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-517">Requirements</span></span>

|<span data-ttu-id="e971d-518">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-518">Requirement</span></span>| <span data-ttu-id="e971d-519">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-520">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-521">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-521">1.0</span></span>|
|[<span data-ttu-id="e971d-522">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-523">ReadItem</span></span>|
|[<span data-ttu-id="e971d-524">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-525">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-526">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-526">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="e971d-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-528">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="e971d-528">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="e971d-529">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-529">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e971d-530">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e971d-530">Read mode</span></span>

<span data-ttu-id="e971d-531">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="e971d-531">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="e971d-532">По умолчанию коллекция ограничена максимум 100 членами.</span><span class="sxs-lookup"><span data-stu-id="e971d-532">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e971d-533">Однако в Windows и Mac вы можете получить максимум 500 членов.</span><span class="sxs-lookup"><span data-stu-id="e971d-533">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="e971d-534">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e971d-534">Compose mode</span></span>

<span data-ttu-id="e971d-535">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="e971d-535">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="e971d-536">По умолчанию коллекция ограничена максимум 100 членами.</span><span class="sxs-lookup"><span data-stu-id="e971d-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e971d-537">Однако в Windows и Mac применяются следующие пределы.</span><span class="sxs-lookup"><span data-stu-id="e971d-537">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e971d-538">Максимальное число участников для получения 500.</span><span class="sxs-lookup"><span data-stu-id="e971d-538">Get 500 members maximum.</span></span>
- <span data-ttu-id="e971d-539">Задайте не более 100 членов для каждого вызова, до 500 всего членов.</span><span class="sxs-lookup"><span data-stu-id="e971d-539">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="e971d-540">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-540">Type</span></span>

*   <span data-ttu-id="e971d-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-542">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-542">Requirements</span></span>

|<span data-ttu-id="e971d-543">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-543">Requirement</span></span>| <span data-ttu-id="e971d-544">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-545">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-546">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-546">1.0</span></span>|
|[<span data-ttu-id="e971d-547">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-548">ReadItem</span></span>|
|[<span data-ttu-id="e971d-549">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-550">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-550">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13"></a><span data-ttu-id="e971d-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-p132">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e971d-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="e971d-p133">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="e971d-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-556">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e971d-556">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e971d-557">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-557">Type</span></span>

*   [<span data-ttu-id="e971d-558">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e971d-558">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="e971d-559">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-559">Requirements</span></span>

|<span data-ttu-id="e971d-560">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-560">Requirement</span></span>| <span data-ttu-id="e971d-561">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-562">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-563">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-563">1.0</span></span>|
|[<span data-ttu-id="e971d-564">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-565">ReadItem</span></span>|
|[<span data-ttu-id="e971d-566">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-566">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-567">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-567">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-568">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-568">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-13"></a><span data-ttu-id="e971d-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-570">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-570">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="e971d-p134">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="e971d-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e971d-573">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e971d-573">Read mode</span></span>

<span data-ttu-id="e971d-574">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="e971d-574">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="e971d-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e971d-575">Compose mode</span></span>

<span data-ttu-id="e971d-576">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="e971d-576">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="e971d-577">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="e971d-577">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="e971d-578">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e971d-578">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.3#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e971d-579">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-579">Type</span></span>

*   <span data-ttu-id="e971d-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-581">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-581">Requirements</span></span>

|<span data-ttu-id="e971d-582">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-582">Requirement</span></span>| <span data-ttu-id="e971d-583">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-584">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-585">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-585">1.0</span></span>|
|[<span data-ttu-id="e971d-586">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-586">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-587">ReadItem</span></span>|
|[<span data-ttu-id="e971d-588">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-589">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-589">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-13"></a><span data-ttu-id="e971d-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-591">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-591">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="e971d-592">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="e971d-592">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e971d-593">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e971d-593">Read mode</span></span>

<span data-ttu-id="e971d-p135">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="e971d-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="e971d-596">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e971d-596">Compose mode</span></span>

<span data-ttu-id="e971d-597">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="e971d-597">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="e971d-598">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-598">Type</span></span>

*   <span data-ttu-id="e971d-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-600">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-600">Requirements</span></span>

|<span data-ttu-id="e971d-601">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-601">Requirement</span></span>| <span data-ttu-id="e971d-602">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-603">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-604">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-604">1.0</span></span>|
|[<span data-ttu-id="e971d-605">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-605">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-606">ReadItem</span></span>|
|[<span data-ttu-id="e971d-607">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-607">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-608">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-608">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-13recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-13"></a><span data-ttu-id="e971d-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

<span data-ttu-id="e971d-610">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-610">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="e971d-611">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-611">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e971d-612">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e971d-612">Read mode</span></span>

<span data-ttu-id="e971d-613">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-613">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="e971d-614">По умолчанию коллекция ограничена максимум 100 членами.</span><span class="sxs-lookup"><span data-stu-id="e971d-614">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e971d-615">Однако в Windows и Mac вы можете получить максимум 500 членов.</span><span class="sxs-lookup"><span data-stu-id="e971d-615">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="e971d-616">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e971d-616">Compose mode</span></span>

<span data-ttu-id="e971d-617">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-617">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="e971d-618">По умолчанию коллекция ограничена максимум 100 членами.</span><span class="sxs-lookup"><span data-stu-id="e971d-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e971d-619">Однако в Windows и Mac применяются следующие пределы.</span><span class="sxs-lookup"><span data-stu-id="e971d-619">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e971d-620">Максимальное число участников для получения 500.</span><span class="sxs-lookup"><span data-stu-id="e971d-620">Get 500 members maximum.</span></span>
- <span data-ttu-id="e971d-621">Задайте не более 100 членов для каждого вызова, до 500 всего членов.</span><span class="sxs-lookup"><span data-stu-id="e971d-621">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e971d-622">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-622">Type</span></span>

*   <span data-ttu-id="e971d-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.3)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-624">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-624">Requirements</span></span>

|<span data-ttu-id="e971d-625">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-625">Requirement</span></span>| <span data-ttu-id="e971d-626">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-627">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-628">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-628">1.0</span></span>|
|[<span data-ttu-id="e971d-629">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-630">ReadItem</span></span>|
|[<span data-ttu-id="e971d-631">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-632">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-632">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e971d-633">Методы</span><span class="sxs-lookup"><span data-stu-id="e971d-633">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="e971d-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e971d-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e971d-635">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="e971d-635">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e971d-636">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="e971d-636">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="e971d-637">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e971d-637">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-638">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-638">Parameters</span></span>

|<span data-ttu-id="e971d-639">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-639">Name</span></span>| <span data-ttu-id="e971d-640">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-640">Type</span></span>| <span data-ttu-id="e971d-641">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e971d-641">Attributes</span></span>| <span data-ttu-id="e971d-642">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-642">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="e971d-643">String</span><span class="sxs-lookup"><span data-stu-id="e971d-643">String</span></span>||<span data-ttu-id="e971d-p139">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="e971d-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="e971d-646">String</span><span class="sxs-lookup"><span data-stu-id="e971d-646">String</span></span>||<span data-ttu-id="e971d-p140">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e971d-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="e971d-649">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-649">Object</span></span>| <span data-ttu-id="e971d-650">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-650">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-651">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e971d-651">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e971d-652">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-652">Object</span></span>| <span data-ttu-id="e971d-653">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-653">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-654">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-654">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e971d-655">функция</span><span class="sxs-lookup"><span data-stu-id="e971d-655">function</span></span>| <span data-ttu-id="e971d-656">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-656">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-657">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e971d-657">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e971d-658">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e971d-658">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e971d-659">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="e971d-659">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e971d-660">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e971d-660">Errors</span></span>

| <span data-ttu-id="e971d-661">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e971d-661">Error code</span></span> | <span data-ttu-id="e971d-662">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-662">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="e971d-663">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="e971d-663">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="e971d-664">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="e971d-664">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="e971d-665">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="e971d-665">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e971d-666">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-666">Requirements</span></span>

|<span data-ttu-id="e971d-667">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-667">Requirement</span></span>| <span data-ttu-id="e971d-668">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-669">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-670">1.1</span><span class="sxs-lookup"><span data-stu-id="e971d-670">1.1</span></span>|
|[<span data-ttu-id="e971d-671">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-672">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e971d-672">ReadWriteItem</span></span>|
|[<span data-ttu-id="e971d-673">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-674">Создание</span><span class="sxs-lookup"><span data-stu-id="e971d-674">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-675">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-675">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="e971d-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e971d-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e971d-677">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="e971d-677">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="e971d-p141">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="e971d-681">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e971d-681">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="e971d-682">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="e971d-682">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-683">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-683">Parameters</span></span>

|<span data-ttu-id="e971d-684">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-684">Name</span></span>| <span data-ttu-id="e971d-685">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-685">Type</span></span>| <span data-ttu-id="e971d-686">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e971d-686">Attributes</span></span>| <span data-ttu-id="e971d-687">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-687">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="e971d-688">String</span><span class="sxs-lookup"><span data-stu-id="e971d-688">String</span></span>||<span data-ttu-id="e971d-p142">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="e971d-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="e971d-691">String</span><span class="sxs-lookup"><span data-stu-id="e971d-691">String</span></span>||<span data-ttu-id="e971d-692">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-692">The subject of the item to be attached.</span></span> <span data-ttu-id="e971d-693">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e971d-693">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="e971d-694">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-694">Object</span></span>| <span data-ttu-id="e971d-695">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-695">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-696">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e971d-696">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e971d-697">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-697">Object</span></span>| <span data-ttu-id="e971d-698">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-698">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-699">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-699">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e971d-700">функция</span><span class="sxs-lookup"><span data-stu-id="e971d-700">function</span></span>| <span data-ttu-id="e971d-701">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-701">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-702">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e971d-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e971d-703">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e971d-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e971d-704">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="e971d-704">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e971d-705">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e971d-705">Errors</span></span>

| <span data-ttu-id="e971d-706">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e971d-706">Error code</span></span> | <span data-ttu-id="e971d-707">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-707">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="e971d-708">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="e971d-708">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e971d-709">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-709">Requirements</span></span>

|<span data-ttu-id="e971d-710">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-710">Requirement</span></span>| <span data-ttu-id="e971d-711">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-712">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-713">1.1</span><span class="sxs-lookup"><span data-stu-id="e971d-713">1.1</span></span>|
|[<span data-ttu-id="e971d-714">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-714">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-715">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e971d-715">ReadWriteItem</span></span>|
|[<span data-ttu-id="e971d-716">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-716">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-717">Создание</span><span class="sxs-lookup"><span data-stu-id="e971d-717">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-718">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-718">Example</span></span>

<span data-ttu-id="e971d-719">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="e971d-719">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="e971d-720">close()</span><span class="sxs-lookup"><span data-stu-id="e971d-720">close()</span></span>

<span data-ttu-id="e971d-721">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="e971d-721">Closes the current item that is being composed.</span></span>

<span data-ttu-id="e971d-p144">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="e971d-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-724">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="e971d-724">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="e971d-725">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="e971d-725">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-726">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-726">Requirements</span></span>

|<span data-ttu-id="e971d-727">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-727">Requirement</span></span>| <span data-ttu-id="e971d-728">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-729">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-730">1.3</span><span class="sxs-lookup"><span data-stu-id="e971d-730">1.3</span></span>|
|[<span data-ttu-id="e971d-731">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-732">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e971d-732">Restricted</span></span>|
|[<span data-ttu-id="e971d-733">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-734">Создание</span><span class="sxs-lookup"><span data-stu-id="e971d-734">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="e971d-735">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="e971d-735">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="e971d-736">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-736">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-737">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e971d-737">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e971d-738">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="e971d-738">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e971d-739">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="e971d-739">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="e971d-p145">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="e971d-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-743">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-743">Parameters</span></span>

|<span data-ttu-id="e971d-744">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-744">Name</span></span>| <span data-ttu-id="e971d-745">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-745">Type</span></span>| <span data-ttu-id="e971d-746">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-746">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="e971d-747">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e971d-747">String &#124; Object</span></span>| |<span data-ttu-id="e971d-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e971d-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e971d-750">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="e971d-750">**OR**</span></span><br/><span data-ttu-id="e971d-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="e971d-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="e971d-753">String</span><span class="sxs-lookup"><span data-stu-id="e971d-753">String</span></span> | <span data-ttu-id="e971d-754">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-754">&lt;optional&gt;</span></span> | <span data-ttu-id="e971d-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e971d-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="e971d-757">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-757">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="e971d-758">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-758">&lt;optional&gt;</span></span> | <span data-ttu-id="e971d-759">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="e971d-759">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="e971d-760">String</span><span class="sxs-lookup"><span data-stu-id="e971d-760">String</span></span> | | <span data-ttu-id="e971d-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="e971d-763">Строка</span><span class="sxs-lookup"><span data-stu-id="e971d-763">String</span></span> | | <span data-ttu-id="e971d-764">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e971d-764">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="e971d-765">String</span><span class="sxs-lookup"><span data-stu-id="e971d-765">String</span></span> | | <span data-ttu-id="e971d-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="e971d-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="e971d-768">String</span><span class="sxs-lookup"><span data-stu-id="e971d-768">String</span></span> | | <span data-ttu-id="e971d-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="e971d-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="e971d-772">function</span><span class="sxs-lookup"><span data-stu-id="e971d-772">function</span></span> | <span data-ttu-id="e971d-773">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-773">&lt;optional&gt;</span></span> | <span data-ttu-id="e971d-774">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e971d-774">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e971d-775">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-775">Requirements</span></span>

|<span data-ttu-id="e971d-776">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-776">Requirement</span></span>| <span data-ttu-id="e971d-777">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-777">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-778">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-778">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-779">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-779">1.0</span></span>|
|[<span data-ttu-id="e971d-780">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-780">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-781">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-781">ReadItem</span></span>|
|[<span data-ttu-id="e971d-782">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-782">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-783">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-783">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e971d-784">Примеры</span><span class="sxs-lookup"><span data-stu-id="e971d-784">Examples</span></span>

<span data-ttu-id="e971d-785">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="e971d-785">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="e971d-786">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-786">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="e971d-787">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-787">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e971d-788">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="e971d-788">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e971d-789">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="e971d-789">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e971d-790">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="e971d-790">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="e971d-791">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="e971d-791">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="e971d-792">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-792">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-793">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e971d-793">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e971d-794">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="e971d-794">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e971d-795">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="e971d-795">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="e971d-p152">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="e971d-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-799">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-799">Parameters</span></span>

|<span data-ttu-id="e971d-800">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-800">Name</span></span>| <span data-ttu-id="e971d-801">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-801">Type</span></span>| <span data-ttu-id="e971d-802">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-802">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="e971d-803">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e971d-803">String &#124; Object</span></span>| | <span data-ttu-id="e971d-p153">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e971d-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e971d-806">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="e971d-806">**OR**</span></span><br/><span data-ttu-id="e971d-p154">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="e971d-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="e971d-809">String</span><span class="sxs-lookup"><span data-stu-id="e971d-809">String</span></span> | <span data-ttu-id="e971d-810">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-810">&lt;optional&gt;</span></span> | <span data-ttu-id="e971d-p155">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e971d-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="e971d-813">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-813">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="e971d-814">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-814">&lt;optional&gt;</span></span> | <span data-ttu-id="e971d-815">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="e971d-815">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="e971d-816">String</span><span class="sxs-lookup"><span data-stu-id="e971d-816">String</span></span> | | <span data-ttu-id="e971d-p156">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="e971d-819">Строка</span><span class="sxs-lookup"><span data-stu-id="e971d-819">String</span></span> | | <span data-ttu-id="e971d-820">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e971d-820">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="e971d-821">Строка</span><span class="sxs-lookup"><span data-stu-id="e971d-821">String</span></span> | | <span data-ttu-id="e971d-p157">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="e971d-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="e971d-824">String</span><span class="sxs-lookup"><span data-stu-id="e971d-824">String</span></span> | | <span data-ttu-id="e971d-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="e971d-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="e971d-828">function</span><span class="sxs-lookup"><span data-stu-id="e971d-828">function</span></span> | <span data-ttu-id="e971d-829">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-829">&lt;optional&gt;</span></span> | <span data-ttu-id="e971d-830">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e971d-830">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e971d-831">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-831">Requirements</span></span>

|<span data-ttu-id="e971d-832">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-832">Requirement</span></span>| <span data-ttu-id="e971d-833">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-834">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-835">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-835">1.0</span></span>|
|[<span data-ttu-id="e971d-836">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-837">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-837">ReadItem</span></span>|
|[<span data-ttu-id="e971d-838">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-839">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-839">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e971d-840">Примеры</span><span class="sxs-lookup"><span data-stu-id="e971d-840">Examples</span></span>

<span data-ttu-id="e971d-841">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="e971d-841">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="e971d-842">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-842">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="e971d-843">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-843">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e971d-844">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="e971d-844">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e971d-845">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="e971d-845">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e971d-846">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="e971d-846">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-13"></a><span data-ttu-id="e971d-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="e971d-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="e971d-848">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-848">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-849">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e971d-849">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-850">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-850">Requirements</span></span>

|<span data-ttu-id="e971d-851">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-851">Requirement</span></span>| <span data-ttu-id="e971d-852">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-853">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-853">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-854">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-854">1.0</span></span>|
|[<span data-ttu-id="e971d-855">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-855">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-856">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-856">ReadItem</span></span>|
|[<span data-ttu-id="e971d-857">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-857">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-858">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-858">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e971d-859">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e971d-859">Returns:</span></span>

<span data-ttu-id="e971d-860">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e971d-860">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.3)</span></span>

##### <a name="example"></a><span data-ttu-id="e971d-861">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-861">Example</span></span>

<span data-ttu-id="e971d-862">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-862">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="e971d-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="e971d-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="e971d-864">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-864">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-865">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e971d-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-866">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-866">Parameters</span></span>

|<span data-ttu-id="e971d-867">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-867">Name</span></span>| <span data-ttu-id="e971d-868">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-868">Type</span></span>| <span data-ttu-id="e971d-869">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-869">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="e971d-870">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="e971d-870">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.3)|<span data-ttu-id="e971d-871">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="e971d-871">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e971d-872">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-872">Requirements</span></span>

|<span data-ttu-id="e971d-873">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-873">Requirement</span></span>| <span data-ttu-id="e971d-874">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-874">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-875">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-875">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-876">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-876">1.0</span></span>|
|[<span data-ttu-id="e971d-877">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-877">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-878">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e971d-878">Restricted</span></span>|
|[<span data-ttu-id="e971d-879">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-879">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-880">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-880">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e971d-881">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e971d-881">Returns:</span></span>

<span data-ttu-id="e971d-882">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="e971d-882">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="e971d-883">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="e971d-883">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="e971d-884">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="e971d-884">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="e971d-885">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="e971d-885">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="e971d-886">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="e971d-886">Value of `entityType`</span></span> | <span data-ttu-id="e971d-887">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="e971d-887">Type of objects in returned array</span></span> | <span data-ttu-id="e971d-888">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-888">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="e971d-889">String</span><span class="sxs-lookup"><span data-stu-id="e971d-889">String</span></span> | <span data-ttu-id="e971d-890">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="e971d-890">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="e971d-891">Contact</span><span class="sxs-lookup"><span data-stu-id="e971d-891">Contact</span></span> | <span data-ttu-id="e971d-892">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e971d-892">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="e971d-893">String</span><span class="sxs-lookup"><span data-stu-id="e971d-893">String</span></span> | <span data-ttu-id="e971d-894">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e971d-894">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="e971d-895">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="e971d-895">MeetingSuggestion</span></span> | <span data-ttu-id="e971d-896">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e971d-896">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="e971d-897">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="e971d-897">PhoneNumber</span></span> | <span data-ttu-id="e971d-898">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="e971d-898">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="e971d-899">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="e971d-899">TaskSuggestion</span></span> | <span data-ttu-id="e971d-900">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e971d-900">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="e971d-901">String</span><span class="sxs-lookup"><span data-stu-id="e971d-901">String</span></span> | <span data-ttu-id="e971d-902">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="e971d-902">**Restricted**</span></span> |

<span data-ttu-id="e971d-903">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="e971d-903">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

##### <a name="example"></a><span data-ttu-id="e971d-904">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-904">Example</span></span>

<span data-ttu-id="e971d-905">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-905">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-13meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-13phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-13tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-13"></a><span data-ttu-id="e971d-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span><span class="sxs-lookup"><span data-stu-id="e971d-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))>}</span></span>

<span data-ttu-id="e971d-907">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e971d-907">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-908">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e971d-908">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e971d-909">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="e971d-909">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-910">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-910">Parameters</span></span>

|<span data-ttu-id="e971d-911">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-911">Name</span></span>| <span data-ttu-id="e971d-912">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-912">Type</span></span>| <span data-ttu-id="e971d-913">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-913">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="e971d-914">String</span><span class="sxs-lookup"><span data-stu-id="e971d-914">String</span></span>|<span data-ttu-id="e971d-915">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="e971d-915">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e971d-916">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-916">Requirements</span></span>

|<span data-ttu-id="e971d-917">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-917">Requirement</span></span>| <span data-ttu-id="e971d-918">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-918">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-919">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-919">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-920">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-920">1.0</span></span>|
|[<span data-ttu-id="e971d-921">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-921">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-922">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-922">ReadItem</span></span>|
|[<span data-ttu-id="e971d-923">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-923">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-924">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-924">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e971d-925">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e971d-925">Returns:</span></span>

<span data-ttu-id="e971d-p160">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="e971d-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="e971d-928">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span><span class="sxs-lookup"><span data-stu-id="e971d-928">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.3)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.3)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.3)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.3))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="e971d-929">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e971d-929">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="e971d-930">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e971d-930">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-931">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e971d-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e971d-p161">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="e971d-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e971d-935">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="e971d-935">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e971d-936">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="e971d-936">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e971d-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="e971d-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e971d-940">Requirements</span><span class="sxs-lookup"><span data-stu-id="e971d-940">Requirements</span></span>

|<span data-ttu-id="e971d-941">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-941">Requirement</span></span>| <span data-ttu-id="e971d-942">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-943">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-944">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-944">1.0</span></span>|
|[<span data-ttu-id="e971d-945">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-946">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-946">ReadItem</span></span>|
|[<span data-ttu-id="e971d-947">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-948">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-948">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e971d-949">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e971d-949">Returns:</span></span>

<span data-ttu-id="e971d-p163">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="e971d-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="e971d-952">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="e971d-952">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="e971d-953">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-953">Example</span></span>

<span data-ttu-id="e971d-954">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="e971d-954">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="e971d-955">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="e971d-955">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="e971d-956">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e971d-956">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-957">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e971d-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e971d-958">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="e971d-958">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="e971d-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="e971d-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-961">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-961">Parameters</span></span>

|<span data-ttu-id="e971d-962">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-962">Name</span></span>| <span data-ttu-id="e971d-963">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-963">Type</span></span>| <span data-ttu-id="e971d-964">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-964">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="e971d-965">String</span><span class="sxs-lookup"><span data-stu-id="e971d-965">String</span></span>|<span data-ttu-id="e971d-966">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="e971d-966">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e971d-967">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-967">Requirements</span></span>

|<span data-ttu-id="e971d-968">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-968">Requirement</span></span>| <span data-ttu-id="e971d-969">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-969">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-970">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-970">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-971">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-971">1.0</span></span>|
|[<span data-ttu-id="e971d-972">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-972">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-973">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-973">ReadItem</span></span>|
|[<span data-ttu-id="e971d-974">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-974">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-975">Чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-975">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e971d-976">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e971d-976">Returns:</span></span>

<span data-ttu-id="e971d-977">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e971d-977">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="e971d-978">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="e971d-978">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="e971d-979">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-979">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="e971d-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="e971d-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="e971d-981">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-981">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="e971d-p165">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="e971d-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-984">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-984">Parameters</span></span>

|<span data-ttu-id="e971d-985">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-985">Name</span></span>| <span data-ttu-id="e971d-986">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-986">Type</span></span>| <span data-ttu-id="e971d-987">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e971d-987">Attributes</span></span>| <span data-ttu-id="e971d-988">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-988">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="e971d-989">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e971d-989">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="e971d-p166">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="e971d-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="e971d-993">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-993">Object</span></span>| <span data-ttu-id="e971d-994">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-994">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-995">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e971d-995">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e971d-996">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-996">Object</span></span>| <span data-ttu-id="e971d-997">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-997">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-998">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-998">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e971d-999">функция</span><span class="sxs-lookup"><span data-stu-id="e971d-999">function</span></span>||<span data-ttu-id="e971d-1000">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e971d-1000">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e971d-1001">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="e971d-1001">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="e971d-1002">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="e971d-1002">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e971d-1003">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-1003">Requirements</span></span>

|<span data-ttu-id="e971d-1004">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-1004">Requirement</span></span>| <span data-ttu-id="e971d-1005">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-1006">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-1007">1.2</span><span class="sxs-lookup"><span data-stu-id="e971d-1007">1.2</span></span>|
|[<span data-ttu-id="e971d-1008">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-1008">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-1009">ReadItem</span></span>|
|[<span data-ttu-id="e971d-1010">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-1010">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-1011">Создание</span><span class="sxs-lookup"><span data-stu-id="e971d-1011">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e971d-1012">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e971d-1012">Returns:</span></span>

<span data-ttu-id="e971d-1013">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="e971d-1013">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="e971d-1014">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="e971d-1014">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e971d-1015">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-1015">Example</span></span>

```js
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

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="e971d-1016">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e971d-1016">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="e971d-1017">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-1017">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="e971d-p168">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="e971d-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-1021">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-1021">Parameters</span></span>

|<span data-ttu-id="e971d-1022">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-1022">Name</span></span>| <span data-ttu-id="e971d-1023">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-1023">Type</span></span>| <span data-ttu-id="e971d-1024">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e971d-1024">Attributes</span></span>| <span data-ttu-id="e971d-1025">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-1025">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e971d-1026">function</span><span class="sxs-lookup"><span data-stu-id="e971d-1026">function</span></span>||<span data-ttu-id="e971d-1027">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e971d-1027">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e971d-1028">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e971d-1028">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.3) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e971d-1029">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="e971d-1029">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="e971d-1030">Объект</span><span class="sxs-lookup"><span data-stu-id="e971d-1030">Object</span></span>| <span data-ttu-id="e971d-1031">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-1031">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-1032">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-1032">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="e971d-1033">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-1033">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e971d-1034">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-1034">Requirements</span></span>

|<span data-ttu-id="e971d-1035">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-1035">Requirement</span></span>| <span data-ttu-id="e971d-1036">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-1036">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-1037">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-1037">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-1038">1.0</span><span class="sxs-lookup"><span data-stu-id="e971d-1038">1.0</span></span>|
|[<span data-ttu-id="e971d-1039">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-1039">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-1040">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e971d-1040">ReadItem</span></span>|
|[<span data-ttu-id="e971d-1041">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-1041">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-1042">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e971d-1042">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-1043">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-1043">Example</span></span>

<span data-ttu-id="e971d-p171">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="e971d-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="e971d-1047">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e971d-1047">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="e971d-1048">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="e971d-1048">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="e971d-1049">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="e971d-1049">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="e971d-1050">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e971d-1050">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="e971d-1051">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="e971d-1051">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="e971d-1052">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="e971d-1052">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-1053">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-1053">Parameters</span></span>

|<span data-ttu-id="e971d-1054">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-1054">Name</span></span>| <span data-ttu-id="e971d-1055">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-1055">Type</span></span>| <span data-ttu-id="e971d-1056">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e971d-1056">Attributes</span></span>| <span data-ttu-id="e971d-1057">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-1057">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="e971d-1058">String</span><span class="sxs-lookup"><span data-stu-id="e971d-1058">String</span></span>||<span data-ttu-id="e971d-1059">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="e971d-1059">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="e971d-1060">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-1060">Object</span></span>| <span data-ttu-id="e971d-1061">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-1061">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-1062">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e971d-1062">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e971d-1063">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-1063">Object</span></span>| <span data-ttu-id="e971d-1064">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-1065">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-1065">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e971d-1066">функция</span><span class="sxs-lookup"><span data-stu-id="e971d-1066">function</span></span>| <span data-ttu-id="e971d-1067">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-1068">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e971d-1068">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e971d-1069">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="e971d-1069">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e971d-1070">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e971d-1070">Errors</span></span>

| <span data-ttu-id="e971d-1071">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e971d-1071">Error code</span></span> | <span data-ttu-id="e971d-1072">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-1072">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="e971d-1073">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="e971d-1073">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e971d-1074">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-1074">Requirements</span></span>

|<span data-ttu-id="e971d-1075">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-1075">Requirement</span></span>| <span data-ttu-id="e971d-1076">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-1076">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-1077">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e971d-1077">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-1078">1.1</span><span class="sxs-lookup"><span data-stu-id="e971d-1078">1.1</span></span>|
|[<span data-ttu-id="e971d-1079">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-1079">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-1080">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e971d-1080">ReadWriteItem</span></span>|
|[<span data-ttu-id="e971d-1081">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-1081">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-1082">Создание</span><span class="sxs-lookup"><span data-stu-id="e971d-1082">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-1083">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-1083">Example</span></span>

<span data-ttu-id="e971d-1084">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="e971d-1084">The following code removes an attachment with an identifier of '0'.</span></span>

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

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="e971d-1085">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e971d-1085">saveAsync([options], callback)</span></span>

<span data-ttu-id="e971d-1086">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="e971d-1086">Asynchronously saves an item.</span></span>

<span data-ttu-id="e971d-1087">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-1087">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="e971d-1088">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="e971d-1088">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="e971d-1089">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="e971d-1089">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-1090">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="e971d-1090">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="e971d-1091">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="e971d-1091">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="e971d-p175">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="e971d-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="e971d-1095">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="e971d-1095">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="e971d-1096">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="e971d-1096">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="e971d-1097">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e971d-1097">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="e971d-1098">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="e971d-1098">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="e971d-1099">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e971d-1099">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-1100">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-1100">Parameters</span></span>

|<span data-ttu-id="e971d-1101">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-1101">Name</span></span>| <span data-ttu-id="e971d-1102">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-1102">Type</span></span>| <span data-ttu-id="e971d-1103">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e971d-1103">Attributes</span></span>| <span data-ttu-id="e971d-1104">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-1104">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="e971d-1105">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-1105">Object</span></span>| <span data-ttu-id="e971d-1106">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-1106">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-1107">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e971d-1107">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e971d-1108">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-1108">Object</span></span>| <span data-ttu-id="e971d-1109">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-1109">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-1110">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e971d-1110">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e971d-1111">функция</span><span class="sxs-lookup"><span data-stu-id="e971d-1111">function</span></span>||<span data-ttu-id="e971d-1112">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e971d-1112">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e971d-1113">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e971d-1113">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e971d-1114">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-1114">Requirements</span></span>

|<span data-ttu-id="e971d-1115">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-1115">Requirement</span></span>| <span data-ttu-id="e971d-1116">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-1116">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-1117">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-1117">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-1118">1.3</span><span class="sxs-lookup"><span data-stu-id="e971d-1118">1.3</span></span>|
|[<span data-ttu-id="e971d-1119">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-1119">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-1120">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e971d-1120">ReadWriteItem</span></span>|
|[<span data-ttu-id="e971d-1121">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-1121">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-1122">Создание</span><span class="sxs-lookup"><span data-stu-id="e971d-1122">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e971d-1123">Примеры</span><span class="sxs-lookup"><span data-stu-id="e971d-1123">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="e971d-p177">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e971d-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="e971d-1126">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="e971d-1126">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="e971d-1127">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="e971d-1127">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="e971d-p178">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="e971d-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e971d-1131">Параметры</span><span class="sxs-lookup"><span data-stu-id="e971d-1131">Parameters</span></span>

|<span data-ttu-id="e971d-1132">Имя</span><span class="sxs-lookup"><span data-stu-id="e971d-1132">Name</span></span>| <span data-ttu-id="e971d-1133">Тип</span><span class="sxs-lookup"><span data-stu-id="e971d-1133">Type</span></span>| <span data-ttu-id="e971d-1134">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e971d-1134">Attributes</span></span>| <span data-ttu-id="e971d-1135">Описание</span><span class="sxs-lookup"><span data-stu-id="e971d-1135">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e971d-1136">String</span><span class="sxs-lookup"><span data-stu-id="e971d-1136">String</span></span>||<span data-ttu-id="e971d-p179">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="e971d-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="e971d-1140">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-1140">Object</span></span>| <span data-ttu-id="e971d-1141">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-1141">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-1142">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e971d-1142">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e971d-1143">Object</span><span class="sxs-lookup"><span data-stu-id="e971d-1143">Object</span></span>| <span data-ttu-id="e971d-1144">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-1144">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-1145">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="e971d-1145">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="e971d-1146">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e971d-1146">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="e971d-1147">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e971d-1147">&lt;optional&gt;</span></span>|<span data-ttu-id="e971d-1148">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="e971d-1148">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="e971d-1149">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="e971d-1149">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="e971d-1150">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e971d-1150">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="e971d-1151">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="e971d-1151">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="e971d-1152">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="e971d-1152">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="e971d-1153">функция</span><span class="sxs-lookup"><span data-stu-id="e971d-1153">function</span></span>||<span data-ttu-id="e971d-1154">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e971d-1154">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e971d-1155">Требования</span><span class="sxs-lookup"><span data-stu-id="e971d-1155">Requirements</span></span>

|<span data-ttu-id="e971d-1156">Требование</span><span class="sxs-lookup"><span data-stu-id="e971d-1156">Requirement</span></span>| <span data-ttu-id="e971d-1157">Значение</span><span class="sxs-lookup"><span data-stu-id="e971d-1157">Value</span></span>|
|---|---|
|[<span data-ttu-id="e971d-1158">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e971d-1158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e971d-1159">1.2</span><span class="sxs-lookup"><span data-stu-id="e971d-1159">1.2</span></span>|
|[<span data-ttu-id="e971d-1160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e971d-1160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e971d-1161">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e971d-1161">ReadWriteItem</span></span>|
|[<span data-ttu-id="e971d-1162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e971d-1162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e971d-1163">Создание</span><span class="sxs-lookup"><span data-stu-id="e971d-1163">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e971d-1164">Пример</span><span class="sxs-lookup"><span data-stu-id="e971d-1164">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
