---
title: Office. Context. Mailbox. Item — набор требований 1,2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 536c8b7bece6df6f9609406f3eccc50b330d7925
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268694"
---
# <a name="item"></a><span data-ttu-id="60c40-102">item</span><span class="sxs-lookup"><span data-stu-id="60c40-102">item</span></span>

### <span data-ttu-id="60c40-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="60c40-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="60c40-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="60c40-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="60c40-107">Requirements</span></span>

|<span data-ttu-id="60c40-108">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-108">Requirement</span></span>| <span data-ttu-id="60c40-109">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-111">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-111">1.0</span></span>|
|[<span data-ttu-id="60c40-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="60c40-113">Restricted</span></span>|
|[<span data-ttu-id="60c40-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="60c40-116">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="60c40-116">Members and methods</span></span>

| <span data-ttu-id="60c40-117">Элемент	</span><span class="sxs-lookup"><span data-stu-id="60c40-117">Member</span></span> | <span data-ttu-id="60c40-118">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="60c40-119">attachments</span><span class="sxs-lookup"><span data-stu-id="60c40-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="60c40-120">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-120">Member</span></span> |
| [<span data-ttu-id="60c40-121">bcc</span><span class="sxs-lookup"><span data-stu-id="60c40-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="60c40-122">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-122">Member</span></span> |
| [<span data-ttu-id="60c40-123">body</span><span class="sxs-lookup"><span data-stu-id="60c40-123">body</span></span>](#body-body) | <span data-ttu-id="60c40-124">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-124">Member</span></span> |
| [<span data-ttu-id="60c40-125">cc</span><span class="sxs-lookup"><span data-stu-id="60c40-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="60c40-126">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-126">Member</span></span> |
| [<span data-ttu-id="60c40-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="60c40-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="60c40-128">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-128">Member</span></span> |
| [<span data-ttu-id="60c40-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="60c40-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="60c40-130">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-130">Member</span></span> |
| [<span data-ttu-id="60c40-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="60c40-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="60c40-132">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-132">Member</span></span> |
| [<span data-ttu-id="60c40-133">end</span><span class="sxs-lookup"><span data-stu-id="60c40-133">end</span></span>](#end-datetime) | <span data-ttu-id="60c40-134">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-134">Member</span></span> |
| [<span data-ttu-id="60c40-135">from</span><span class="sxs-lookup"><span data-stu-id="60c40-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="60c40-136">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-136">Member</span></span> |
| [<span data-ttu-id="60c40-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="60c40-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="60c40-138">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-138">Member</span></span> |
| [<span data-ttu-id="60c40-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="60c40-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="60c40-140">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-140">Member</span></span> |
| [<span data-ttu-id="60c40-141">itemId</span><span class="sxs-lookup"><span data-stu-id="60c40-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="60c40-142">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-142">Member</span></span> |
| [<span data-ttu-id="60c40-143">itemType</span><span class="sxs-lookup"><span data-stu-id="60c40-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="60c40-144">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-144">Member</span></span> |
| [<span data-ttu-id="60c40-145">location</span><span class="sxs-lookup"><span data-stu-id="60c40-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="60c40-146">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-146">Member</span></span> |
| [<span data-ttu-id="60c40-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="60c40-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="60c40-148">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-148">Member</span></span> |
| [<span data-ttu-id="60c40-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="60c40-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="60c40-150">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-150">Member</span></span> |
| [<span data-ttu-id="60c40-151">organizer</span><span class="sxs-lookup"><span data-stu-id="60c40-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="60c40-152">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-152">Member</span></span> |
| [<span data-ttu-id="60c40-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="60c40-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="60c40-154">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-154">Member</span></span> |
| [<span data-ttu-id="60c40-155">sender</span><span class="sxs-lookup"><span data-stu-id="60c40-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="60c40-156">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-156">Member</span></span> |
| [<span data-ttu-id="60c40-157">start</span><span class="sxs-lookup"><span data-stu-id="60c40-157">start</span></span>](#start-datetime) | <span data-ttu-id="60c40-158">Member</span><span class="sxs-lookup"><span data-stu-id="60c40-158">Member</span></span> |
| [<span data-ttu-id="60c40-159">subject</span><span class="sxs-lookup"><span data-stu-id="60c40-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="60c40-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="60c40-160">Member</span></span> |
| [<span data-ttu-id="60c40-161">to</span><span class="sxs-lookup"><span data-stu-id="60c40-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="60c40-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="60c40-162">Member</span></span> |
| [<span data-ttu-id="60c40-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="60c40-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="60c40-164">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-164">Method</span></span> |
| [<span data-ttu-id="60c40-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="60c40-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="60c40-166">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-166">Method</span></span> |
| [<span data-ttu-id="60c40-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="60c40-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="60c40-168">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-168">Method</span></span> |
| [<span data-ttu-id="60c40-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="60c40-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="60c40-170">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-170">Method</span></span> |
| [<span data-ttu-id="60c40-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="60c40-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="60c40-172">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-172">Method</span></span> |
| [<span data-ttu-id="60c40-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="60c40-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="60c40-174">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-174">Method</span></span> |
| [<span data-ttu-id="60c40-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="60c40-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="60c40-176">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-176">Method</span></span> |
| [<span data-ttu-id="60c40-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="60c40-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="60c40-178">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-178">Method</span></span> |
| [<span data-ttu-id="60c40-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="60c40-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="60c40-180">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-180">Method</span></span> |
| [<span data-ttu-id="60c40-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="60c40-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="60c40-182">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-182">Method</span></span> |
| [<span data-ttu-id="60c40-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="60c40-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="60c40-184">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-184">Method</span></span> |
| [<span data-ttu-id="60c40-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="60c40-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="60c40-186">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-186">Method</span></span> |
| [<span data-ttu-id="60c40-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="60c40-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="60c40-188">Метод</span><span class="sxs-lookup"><span data-stu-id="60c40-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="60c40-189">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-189">Example</span></span>

<span data-ttu-id="60c40-190">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="60c40-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="60c40-191">Элементы</span><span class="sxs-lookup"><span data-stu-id="60c40-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="60c40-192">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="60c40-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="60c40-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-195">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="60c40-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="60c40-196">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="60c40-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-197">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-197">Type</span></span>

*   <span data-ttu-id="60c40-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="60c40-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-199">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-199">Requirements</span></span>

|<span data-ttu-id="60c40-200">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-200">Requirement</span></span>| <span data-ttu-id="60c40-201">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-202">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-203">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-203">1.0</span></span>|
|[<span data-ttu-id="60c40-204">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-205">ReadItem</span></span>|
|[<span data-ttu-id="60c40-206">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-207">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-208">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-208">Example</span></span>

<span data-ttu-id="60c40-209">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="60c40-210">СК: [получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-211">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="60c40-212">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="60c40-212">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-213">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-213">Type</span></span>

*   [<span data-ttu-id="60c40-214">Получатели</span><span class="sxs-lookup"><span data-stu-id="60c40-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="60c40-215">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-215">Requirements</span></span>

|<span data-ttu-id="60c40-216">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-216">Requirement</span></span>| <span data-ttu-id="60c40-217">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-218">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-219">1.1</span><span class="sxs-lookup"><span data-stu-id="60c40-219">1.1</span></span>|
|[<span data-ttu-id="60c40-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-221">ReadItem</span></span>|
|[<span data-ttu-id="60c40-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-223">Создание</span><span class="sxs-lookup"><span data-stu-id="60c40-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-224">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-224">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="60c40-225">основной текст: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-226">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-227">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-227">Type</span></span>

*   [<span data-ttu-id="60c40-228">Body</span><span class="sxs-lookup"><span data-stu-id="60c40-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="60c40-229">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-229">Requirements</span></span>

|<span data-ttu-id="60c40-230">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-230">Requirement</span></span>| <span data-ttu-id="60c40-231">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-232">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-233">1.1</span><span class="sxs-lookup"><span data-stu-id="60c40-233">1.1</span></span>|
|[<span data-ttu-id="60c40-234">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-235">ReadItem</span></span>|
|[<span data-ttu-id="60c40-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-238">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-238">Example</span></span>

<span data-ttu-id="60c40-239">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="60c40-239">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="60c40-240">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="60c40-240">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="60c40-241">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-242">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="60c40-243">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="60c40-244">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="60c40-244">Read mode</span></span>

<span data-ttu-id="60c40-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="60c40-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="60c40-247">Режим создания</span><span class="sxs-lookup"><span data-stu-id="60c40-247">Compose mode</span></span>

<span data-ttu-id="60c40-248">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-248">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="60c40-249">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-249">Type</span></span>

*   <span data-ttu-id="60c40-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-251">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-251">Requirements</span></span>

|<span data-ttu-id="60c40-252">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-252">Requirement</span></span>| <span data-ttu-id="60c40-253">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-254">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-255">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-255">1.0</span></span>|
|[<span data-ttu-id="60c40-256">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-257">ReadItem</span></span>|
|[<span data-ttu-id="60c40-258">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-259">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-259">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="60c40-260">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="60c40-260">(nullable) conversationId: String</span></span>

<span data-ttu-id="60c40-261">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="60c40-261">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="60c40-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="60c40-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="60c40-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="60c40-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-266">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-266">Type</span></span>

*   <span data-ttu-id="60c40-267">String</span><span class="sxs-lookup"><span data-stu-id="60c40-267">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-268">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-268">Requirements</span></span>

|<span data-ttu-id="60c40-269">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-269">Requirement</span></span>| <span data-ttu-id="60c40-270">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-271">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-272">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-272">1.0</span></span>|
|[<span data-ttu-id="60c40-273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-274">ReadItem</span></span>|
|[<span data-ttu-id="60c40-275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-277">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-277">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="60c40-278">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="60c40-278">dateTimeCreated: Date</span></span>

<span data-ttu-id="60c40-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-281">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-281">Type</span></span>

*   <span data-ttu-id="60c40-282">Дата</span><span class="sxs-lookup"><span data-stu-id="60c40-282">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-283">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-283">Requirements</span></span>

|<span data-ttu-id="60c40-284">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-284">Requirement</span></span>| <span data-ttu-id="60c40-285">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-286">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-287">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-287">1.0</span></span>|
|[<span data-ttu-id="60c40-288">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-288">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-289">ReadItem</span></span>|
|[<span data-ttu-id="60c40-290">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-290">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-291">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-291">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-292">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-292">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="60c40-293">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="60c40-293">dateTimeModified: Date</span></span>

<span data-ttu-id="60c40-294">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-294">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="60c40-295">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-295">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-296">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="60c40-296">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-297">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-297">Type</span></span>

*   <span data-ttu-id="60c40-298">Дата</span><span class="sxs-lookup"><span data-stu-id="60c40-298">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-299">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-299">Requirements</span></span>

|<span data-ttu-id="60c40-300">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-300">Requirement</span></span>| <span data-ttu-id="60c40-301">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-302">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-303">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-303">1.0</span></span>|
|[<span data-ttu-id="60c40-304">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-305">ReadItem</span></span>|
|[<span data-ttu-id="60c40-306">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-307">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-308">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-308">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="60c40-309">конец: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.2) )</span><span class="sxs-lookup"><span data-stu-id="60c40-309">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-310">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-310">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="60c40-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="60c40-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="60c40-313">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="60c40-313">Read mode</span></span>

<span data-ttu-id="60c40-314">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="60c40-314">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="60c40-315">Режим создания</span><span class="sxs-lookup"><span data-stu-id="60c40-315">Compose mode</span></span>

<span data-ttu-id="60c40-316">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="60c40-316">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="60c40-317">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="60c40-317">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="60c40-318">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-318">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="60c40-319">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-319">Type</span></span>

*   <span data-ttu-id="60c40-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-321">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-321">Requirements</span></span>

|<span data-ttu-id="60c40-322">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-322">Requirement</span></span>| <span data-ttu-id="60c40-323">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-324">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-325">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-325">1.0</span></span>|
|[<span data-ttu-id="60c40-326">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-326">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-327">ReadItem</span></span>|
|[<span data-ttu-id="60c40-328">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-328">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-329">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-329">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="60c40-330">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-330">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="60c40-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="60c40-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-335">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="60c40-335">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-336">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-336">Type</span></span>

*   [<span data-ttu-id="60c40-337">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="60c40-337">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="60c40-338">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-338">Requirements</span></span>

|<span data-ttu-id="60c40-339">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-339">Requirement</span></span>| <span data-ttu-id="60c40-340">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-341">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-342">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-342">1.0</span></span>|
|[<span data-ttu-id="60c40-343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-344">ReadItem</span></span>|
|[<span data-ttu-id="60c40-345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-346">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-346">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-347">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="60c40-348">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="60c40-348">internetMessageId: String</span></span>

<span data-ttu-id="60c40-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-351">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-351">Type</span></span>

*   <span data-ttu-id="60c40-352">String</span><span class="sxs-lookup"><span data-stu-id="60c40-352">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-353">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-353">Requirements</span></span>

|<span data-ttu-id="60c40-354">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-354">Requirement</span></span>| <span data-ttu-id="60c40-355">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-356">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-357">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-357">1.0</span></span>|
|[<span data-ttu-id="60c40-358">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-359">ReadItem</span></span>|
|[<span data-ttu-id="60c40-360">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-361">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-361">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-362">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-362">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="60c40-363">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="60c40-363">itemClass: String</span></span>

<span data-ttu-id="60c40-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="60c40-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="60c40-368">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-368">Type</span></span> | <span data-ttu-id="60c40-369">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-369">Description</span></span> | <span data-ttu-id="60c40-370">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="60c40-370">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="60c40-371">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="60c40-371">Appointment items</span></span> | <span data-ttu-id="60c40-372">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="60c40-372">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="60c40-373">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="60c40-373">Message items</span></span> | <span data-ttu-id="60c40-374">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-374">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="60c40-375">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="60c40-375">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-376">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-376">Type</span></span>

*   <span data-ttu-id="60c40-377">String</span><span class="sxs-lookup"><span data-stu-id="60c40-377">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-378">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-378">Requirements</span></span>

|<span data-ttu-id="60c40-379">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-379">Requirement</span></span>| <span data-ttu-id="60c40-380">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-381">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-382">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-382">1.0</span></span>|
|[<span data-ttu-id="60c40-383">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-384">ReadItem</span></span>|
|[<span data-ttu-id="60c40-385">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-386">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-386">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-387">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-387">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="60c40-388">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="60c40-388">(nullable) itemId: String</span></span>

<span data-ttu-id="60c40-389">Получает идентификатор элемента веб-служб Exchange для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-389">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="60c40-390">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-390">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-391">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="60c40-391">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="60c40-392">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="60c40-392">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="60c40-393">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="60c40-393">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="60c40-394">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="60c40-394">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-395">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-395">Type</span></span>

*   <span data-ttu-id="60c40-396">String</span><span class="sxs-lookup"><span data-stu-id="60c40-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-397">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-397">Requirements</span></span>

|<span data-ttu-id="60c40-398">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-398">Requirement</span></span>| <span data-ttu-id="60c40-399">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-400">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-401">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-401">1.0</span></span>|
|[<span data-ttu-id="60c40-402">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-403">ReadItem</span></span>|
|[<span data-ttu-id="60c40-404">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-405">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-406">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-406">Example</span></span>

<span data-ttu-id="60c40-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="60c40-409">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-409">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-410">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="60c40-410">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="60c40-411">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="60c40-411">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-412">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-412">Type</span></span>

*   [<span data-ttu-id="60c40-413">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="60c40-413">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="60c40-414">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-414">Requirements</span></span>

|<span data-ttu-id="60c40-415">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-415">Requirement</span></span>| <span data-ttu-id="60c40-416">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-416">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-417">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-417">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-418">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-418">1.0</span></span>|
|[<span data-ttu-id="60c40-419">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-419">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-420">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-420">ReadItem</span></span>|
|[<span data-ttu-id="60c40-421">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-421">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-422">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-422">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-423">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-423">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="60c40-424">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location?view=outlook-js-1.2) )</span><span class="sxs-lookup"><span data-stu-id="60c40-424">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-425">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-425">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="60c40-426">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="60c40-426">Read mode</span></span>

<span data-ttu-id="60c40-427">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-427">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="60c40-428">Режим создания</span><span class="sxs-lookup"><span data-stu-id="60c40-428">Compose mode</span></span>

<span data-ttu-id="60c40-429">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-429">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="60c40-430">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-430">Type</span></span>

*   <span data-ttu-id="60c40-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-432">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-432">Requirements</span></span>

|<span data-ttu-id="60c40-433">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-433">Requirement</span></span>| <span data-ttu-id="60c40-434">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-434">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-435">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-435">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-436">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-436">1.0</span></span>|
|[<span data-ttu-id="60c40-437">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-437">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-438">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-438">ReadItem</span></span>|
|[<span data-ttu-id="60c40-439">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-439">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-440">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-440">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="60c40-441">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="60c40-441">normalizedSubject: String</span></span>

<span data-ttu-id="60c40-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="60c40-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="60c40-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-446">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-446">Type</span></span>

*   <span data-ttu-id="60c40-447">String</span><span class="sxs-lookup"><span data-stu-id="60c40-447">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-448">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-448">Requirements</span></span>

|<span data-ttu-id="60c40-449">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-449">Requirement</span></span>| <span data-ttu-id="60c40-450">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-451">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-452">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-452">1.0</span></span>|
|[<span data-ttu-id="60c40-453">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-454">ReadItem</span></span>|
|[<span data-ttu-id="60c40-455">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-456">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-456">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-457">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-457">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="60c40-458">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-458">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-459">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="60c40-459">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="60c40-460">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-460">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="60c40-461">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="60c40-461">Read mode</span></span>

<span data-ttu-id="60c40-462">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="60c40-462">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="60c40-463">Режим создания</span><span class="sxs-lookup"><span data-stu-id="60c40-463">Compose mode</span></span>

<span data-ttu-id="60c40-464">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="60c40-464">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="60c40-465">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-465">Type</span></span>

*   <span data-ttu-id="60c40-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-467">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-467">Requirements</span></span>

|<span data-ttu-id="60c40-468">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-468">Requirement</span></span>| <span data-ttu-id="60c40-469">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-470">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-471">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-471">1.0</span></span>|
|[<span data-ttu-id="60c40-472">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-473">ReadItem</span></span>|
|[<span data-ttu-id="60c40-474">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-475">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-475">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="60c40-476">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-476">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-479">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-479">Type</span></span>

*   [<span data-ttu-id="60c40-480">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="60c40-480">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="60c40-481">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-481">Requirements</span></span>

|<span data-ttu-id="60c40-482">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-482">Requirement</span></span>| <span data-ttu-id="60c40-483">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-484">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-485">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-485">1.0</span></span>|
|[<span data-ttu-id="60c40-486">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-487">ReadItem</span></span>|
|[<span data-ttu-id="60c40-488">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-489">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-490">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-490">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="60c40-491">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-491">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-492">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="60c40-492">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="60c40-493">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-493">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="60c40-494">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="60c40-494">Read mode</span></span>

<span data-ttu-id="60c40-495">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="60c40-495">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="60c40-496">Режим создания</span><span class="sxs-lookup"><span data-stu-id="60c40-496">Compose mode</span></span>

<span data-ttu-id="60c40-497">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="60c40-497">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="60c40-498">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-498">Type</span></span>

*   <span data-ttu-id="60c40-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-500">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-500">Requirements</span></span>

|<span data-ttu-id="60c40-501">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-501">Requirement</span></span>| <span data-ttu-id="60c40-502">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-503">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-504">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-504">1.0</span></span>|
|[<span data-ttu-id="60c40-505">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-506">ReadItem</span></span>|
|[<span data-ttu-id="60c40-507">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-508">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-508">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="60c40-509">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-509">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="60c40-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="60c40-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="60c40-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-514">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="60c40-514">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="60c40-515">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-515">Type</span></span>

*   [<span data-ttu-id="60c40-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="60c40-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="60c40-517">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-517">Requirements</span></span>

|<span data-ttu-id="60c40-518">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-518">Requirement</span></span>| <span data-ttu-id="60c40-519">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-520">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-521">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-521">1.0</span></span>|
|[<span data-ttu-id="60c40-522">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-523">ReadItem</span></span>|
|[<span data-ttu-id="60c40-524">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-525">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-526">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-526">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="60c40-527">Начало: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.2) )</span><span class="sxs-lookup"><span data-stu-id="60c40-527">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-528">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-528">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="60c40-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="60c40-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="60c40-531">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="60c40-531">Read mode</span></span>

<span data-ttu-id="60c40-532">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="60c40-532">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="60c40-533">Режим создания</span><span class="sxs-lookup"><span data-stu-id="60c40-533">Compose mode</span></span>

<span data-ttu-id="60c40-534">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="60c40-534">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="60c40-535">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="60c40-535">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="60c40-536">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="60c40-536">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="60c40-537">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-537">Type</span></span>

*   <span data-ttu-id="60c40-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-539">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-539">Requirements</span></span>

|<span data-ttu-id="60c40-540">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-540">Requirement</span></span>| <span data-ttu-id="60c40-541">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-542">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-543">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-543">1.0</span></span>|
|[<span data-ttu-id="60c40-544">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-545">ReadItem</span></span>|
|[<span data-ttu-id="60c40-546">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-547">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-547">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="60c40-548">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject?view=outlook-js-1.2) )</span><span class="sxs-lookup"><span data-stu-id="60c40-548">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-549">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-549">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="60c40-550">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="60c40-550">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="60c40-551">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="60c40-551">Read mode</span></span>

<span data-ttu-id="60c40-p130">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="60c40-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="60c40-554">Режим создания</span><span class="sxs-lookup"><span data-stu-id="60c40-554">Compose mode</span></span>

<span data-ttu-id="60c40-555">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="60c40-555">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="60c40-556">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-556">Type</span></span>

*   <span data-ttu-id="60c40-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-558">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-558">Requirements</span></span>

|<span data-ttu-id="60c40-559">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-559">Requirement</span></span>| <span data-ttu-id="60c40-560">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-561">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-562">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-562">1.0</span></span>|
|[<span data-ttu-id="60c40-563">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-564">ReadItem</span></span>|
|[<span data-ttu-id="60c40-565">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-566">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-566">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="60c40-567">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-567">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="60c40-568">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-568">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="60c40-569">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-569">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="60c40-570">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="60c40-570">Read mode</span></span>

<span data-ttu-id="60c40-p132">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="60c40-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="60c40-573">Режим создания</span><span class="sxs-lookup"><span data-stu-id="60c40-573">Compose mode</span></span>

<span data-ttu-id="60c40-574">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-574">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="60c40-575">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-575">Type</span></span>

*   <span data-ttu-id="60c40-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-577">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-577">Requirements</span></span>

|<span data-ttu-id="60c40-578">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-578">Requirement</span></span>| <span data-ttu-id="60c40-579">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-580">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-581">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-581">1.0</span></span>|
|[<span data-ttu-id="60c40-582">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-583">ReadItem</span></span>|
|[<span data-ttu-id="60c40-584">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-585">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-585">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="60c40-586">Методы</span><span class="sxs-lookup"><span data-stu-id="60c40-586">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="60c40-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="60c40-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="60c40-588">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="60c40-588">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="60c40-589">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="60c40-589">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="60c40-590">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="60c40-590">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-591">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-591">Parameters</span></span>

|<span data-ttu-id="60c40-592">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-592">Name</span></span>| <span data-ttu-id="60c40-593">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-593">Type</span></span>| <span data-ttu-id="60c40-594">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="60c40-594">Attributes</span></span>| <span data-ttu-id="60c40-595">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-595">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="60c40-596">String</span><span class="sxs-lookup"><span data-stu-id="60c40-596">String</span></span>||<span data-ttu-id="60c40-p133">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="60c40-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="60c40-599">String</span><span class="sxs-lookup"><span data-stu-id="60c40-599">String</span></span>||<span data-ttu-id="60c40-p134">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="60c40-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="60c40-602">Объект</span><span class="sxs-lookup"><span data-stu-id="60c40-602">Object</span></span>| <span data-ttu-id="60c40-603">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-603">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-604">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="60c40-604">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="60c40-605">Объект</span><span class="sxs-lookup"><span data-stu-id="60c40-605">Object</span></span>| <span data-ttu-id="60c40-606">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-606">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-607">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="60c40-607">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="60c40-608">функция</span><span class="sxs-lookup"><span data-stu-id="60c40-608">function</span></span>| <span data-ttu-id="60c40-609">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-609">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-610">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="60c40-610">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="60c40-611">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="60c40-611">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="60c40-612">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="60c40-612">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="60c40-613">Ошибки</span><span class="sxs-lookup"><span data-stu-id="60c40-613">Errors</span></span>

| <span data-ttu-id="60c40-614">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="60c40-614">Error code</span></span> | <span data-ttu-id="60c40-615">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-615">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="60c40-616">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="60c40-616">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="60c40-617">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="60c40-617">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="60c40-618">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="60c40-618">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="60c40-619">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-619">Requirements</span></span>

|<span data-ttu-id="60c40-620">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-620">Requirement</span></span>| <span data-ttu-id="60c40-621">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-621">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-622">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-622">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-623">1.1</span><span class="sxs-lookup"><span data-stu-id="60c40-623">1.1</span></span>|
|[<span data-ttu-id="60c40-624">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-624">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-625">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="60c40-625">ReadWriteItem</span></span>|
|[<span data-ttu-id="60c40-626">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-626">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-627">Создание</span><span class="sxs-lookup"><span data-stu-id="60c40-627">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-628">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-628">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="60c40-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="60c40-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="60c40-630">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="60c40-630">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="60c40-p135">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="60c40-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="60c40-634">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="60c40-634">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="60c40-635">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="60c40-635">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-636">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-636">Parameters</span></span>

|<span data-ttu-id="60c40-637">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-637">Name</span></span>| <span data-ttu-id="60c40-638">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-638">Type</span></span>| <span data-ttu-id="60c40-639">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="60c40-639">Attributes</span></span>| <span data-ttu-id="60c40-640">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-640">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="60c40-641">String</span><span class="sxs-lookup"><span data-stu-id="60c40-641">String</span></span>||<span data-ttu-id="60c40-p136">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="60c40-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="60c40-644">String</span><span class="sxs-lookup"><span data-stu-id="60c40-644">String</span></span>||<span data-ttu-id="60c40-645">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-645">The subject of the item to be attached.</span></span> <span data-ttu-id="60c40-646">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="60c40-646">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="60c40-647">Object</span><span class="sxs-lookup"><span data-stu-id="60c40-647">Object</span></span>| <span data-ttu-id="60c40-648">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-648">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-649">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="60c40-649">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="60c40-650">Объект</span><span class="sxs-lookup"><span data-stu-id="60c40-650">Object</span></span>| <span data-ttu-id="60c40-651">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-651">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-652">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="60c40-652">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="60c40-653">функция</span><span class="sxs-lookup"><span data-stu-id="60c40-653">function</span></span>| <span data-ttu-id="60c40-654">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-654">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-655">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="60c40-655">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="60c40-656">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="60c40-656">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="60c40-657">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="60c40-657">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="60c40-658">Ошибки</span><span class="sxs-lookup"><span data-stu-id="60c40-658">Errors</span></span>

| <span data-ttu-id="60c40-659">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="60c40-659">Error code</span></span> | <span data-ttu-id="60c40-660">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-660">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="60c40-661">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="60c40-661">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="60c40-662">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-662">Requirements</span></span>

|<span data-ttu-id="60c40-663">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-663">Requirement</span></span>| <span data-ttu-id="60c40-664">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-665">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-666">1.1</span><span class="sxs-lookup"><span data-stu-id="60c40-666">1.1</span></span>|
|[<span data-ttu-id="60c40-667">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-668">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="60c40-668">ReadWriteItem</span></span>|
|[<span data-ttu-id="60c40-669">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-670">Создание</span><span class="sxs-lookup"><span data-stu-id="60c40-670">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-671">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-671">Example</span></span>

<span data-ttu-id="60c40-672">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="60c40-672">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="60c40-673">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="60c40-673">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="60c40-674">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-674">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-675">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="60c40-675">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="60c40-676">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="60c40-676">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="60c40-677">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="60c40-677">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="60c40-678">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="60c40-678">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="60c40-679">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="60c40-679">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="60c40-680">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="60c40-680">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-681">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-681">Parameters</span></span>

|<span data-ttu-id="60c40-682">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-682">Name</span></span>| <span data-ttu-id="60c40-683">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-683">Type</span></span>| <span data-ttu-id="60c40-684">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-684">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="60c40-685">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="60c40-685">String &#124; Object</span></span>| |<span data-ttu-id="60c40-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="60c40-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="60c40-688">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="60c40-688">**OR**</span></span><br/><span data-ttu-id="60c40-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="60c40-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="60c40-691">String</span><span class="sxs-lookup"><span data-stu-id="60c40-691">String</span></span> | <span data-ttu-id="60c40-692">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-692">&lt;optional&gt;</span></span> | <span data-ttu-id="60c40-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="60c40-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="60c40-695">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-695">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="60c40-696">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-696">&lt;optional&gt;</span></span> | <span data-ttu-id="60c40-697">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="60c40-697">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="60c40-698">String</span><span class="sxs-lookup"><span data-stu-id="60c40-698">String</span></span> | | <span data-ttu-id="60c40-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="60c40-701">Строка</span><span class="sxs-lookup"><span data-stu-id="60c40-701">String</span></span> | | <span data-ttu-id="60c40-702">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="60c40-702">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="60c40-703">String</span><span class="sxs-lookup"><span data-stu-id="60c40-703">String</span></span> | | <span data-ttu-id="60c40-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="60c40-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="60c40-706">String</span><span class="sxs-lookup"><span data-stu-id="60c40-706">String</span></span> | | <span data-ttu-id="60c40-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="60c40-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="60c40-710">function</span><span class="sxs-lookup"><span data-stu-id="60c40-710">function</span></span> | <span data-ttu-id="60c40-711">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-711">&lt;optional&gt;</span></span> | <span data-ttu-id="60c40-712">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="60c40-712">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="60c40-713">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-713">Requirements</span></span>

|<span data-ttu-id="60c40-714">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-714">Requirement</span></span>| <span data-ttu-id="60c40-715">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-716">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-717">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-717">1.0</span></span>|
|[<span data-ttu-id="60c40-718">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-719">ReadItem</span></span>|
|[<span data-ttu-id="60c40-720">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-721">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-721">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="60c40-722">Примеры</span><span class="sxs-lookup"><span data-stu-id="60c40-722">Examples</span></span>

<span data-ttu-id="60c40-723">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="60c40-723">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="60c40-724">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-724">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="60c40-725">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-725">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="60c40-726">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="60c40-726">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="60c40-727">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="60c40-727">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="60c40-728">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="60c40-728">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="60c40-729">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="60c40-729">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="60c40-730">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-730">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-731">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="60c40-731">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="60c40-732">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="60c40-732">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="60c40-733">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="60c40-733">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="60c40-734">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="60c40-734">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="60c40-735">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="60c40-735">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="60c40-736">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="60c40-736">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-737">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-737">Parameters</span></span>

|<span data-ttu-id="60c40-738">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-738">Name</span></span>| <span data-ttu-id="60c40-739">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-739">Type</span></span>| <span data-ttu-id="60c40-740">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-740">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="60c40-741">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="60c40-741">String &#124; Object</span></span>| | <span data-ttu-id="60c40-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="60c40-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="60c40-744">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="60c40-744">**OR**</span></span><br/><span data-ttu-id="60c40-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="60c40-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="60c40-747">String</span><span class="sxs-lookup"><span data-stu-id="60c40-747">String</span></span> | <span data-ttu-id="60c40-748">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-748">&lt;optional&gt;</span></span> | <span data-ttu-id="60c40-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="60c40-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="60c40-751">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-751">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="60c40-752">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-752">&lt;optional&gt;</span></span> | <span data-ttu-id="60c40-753">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="60c40-753">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="60c40-754">String</span><span class="sxs-lookup"><span data-stu-id="60c40-754">String</span></span> | | <span data-ttu-id="60c40-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="60c40-757">Строка</span><span class="sxs-lookup"><span data-stu-id="60c40-757">String</span></span> | | <span data-ttu-id="60c40-758">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="60c40-758">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="60c40-759">String</span><span class="sxs-lookup"><span data-stu-id="60c40-759">String</span></span> | | <span data-ttu-id="60c40-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="60c40-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="60c40-762">String</span><span class="sxs-lookup"><span data-stu-id="60c40-762">String</span></span> | | <span data-ttu-id="60c40-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="60c40-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="60c40-766">function</span><span class="sxs-lookup"><span data-stu-id="60c40-766">function</span></span> | <span data-ttu-id="60c40-767">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-767">&lt;optional&gt;</span></span> | <span data-ttu-id="60c40-768">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="60c40-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="60c40-769">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-769">Requirements</span></span>

|<span data-ttu-id="60c40-770">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-770">Requirement</span></span>| <span data-ttu-id="60c40-771">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-772">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-773">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-773">1.0</span></span>|
|[<span data-ttu-id="60c40-774">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-774">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-775">ReadItem</span></span>|
|[<span data-ttu-id="60c40-776">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-776">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-777">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-777">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="60c40-778">Примеры</span><span class="sxs-lookup"><span data-stu-id="60c40-778">Examples</span></span>

<span data-ttu-id="60c40-779">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="60c40-779">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="60c40-780">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-780">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="60c40-781">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-781">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="60c40-782">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="60c40-782">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="60c40-783">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="60c40-783">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="60c40-784">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="60c40-784">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="60c40-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="60c40-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="60c40-786">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-786">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-787">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="60c40-787">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-788">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-788">Requirements</span></span>

|<span data-ttu-id="60c40-789">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-789">Requirement</span></span>| <span data-ttu-id="60c40-790">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-791">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-792">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-792">1.0</span></span>|
|[<span data-ttu-id="60c40-793">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-794">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-794">ReadItem</span></span>|
|[<span data-ttu-id="60c40-795">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-796">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-796">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="60c40-797">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="60c40-797">Returns:</span></span>

<span data-ttu-id="60c40-798">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="60c40-798">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="60c40-799">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-799">Example</span></span>

<span data-ttu-id="60c40-800">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-800">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="60c40-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="60c40-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="60c40-802">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-802">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-803">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="60c40-803">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-804">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-804">Parameters</span></span>

|<span data-ttu-id="60c40-805">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-805">Name</span></span>| <span data-ttu-id="60c40-806">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-806">Type</span></span>| <span data-ttu-id="60c40-807">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-807">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="60c40-808">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="60c40-808">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="60c40-809">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="60c40-809">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="60c40-810">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-810">Requirements</span></span>

|<span data-ttu-id="60c40-811">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-811">Requirement</span></span>| <span data-ttu-id="60c40-812">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-813">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-814">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-814">1.0</span></span>|
|[<span data-ttu-id="60c40-815">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-815">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-816">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="60c40-816">Restricted</span></span>|
|[<span data-ttu-id="60c40-817">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-817">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-818">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-818">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="60c40-819">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="60c40-819">Returns:</span></span>

<span data-ttu-id="60c40-820">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="60c40-820">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="60c40-821">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="60c40-821">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="60c40-822">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="60c40-822">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="60c40-823">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="60c40-823">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="60c40-824">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="60c40-824">Value of `entityType`</span></span> | <span data-ttu-id="60c40-825">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="60c40-825">Type of objects in returned array</span></span> | <span data-ttu-id="60c40-826">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-826">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="60c40-827">String</span><span class="sxs-lookup"><span data-stu-id="60c40-827">String</span></span> | <span data-ttu-id="60c40-828">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="60c40-828">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="60c40-829">Contact</span><span class="sxs-lookup"><span data-stu-id="60c40-829">Contact</span></span> | <span data-ttu-id="60c40-830">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="60c40-830">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="60c40-831">String</span><span class="sxs-lookup"><span data-stu-id="60c40-831">String</span></span> | <span data-ttu-id="60c40-832">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="60c40-832">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="60c40-833">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="60c40-833">MeetingSuggestion</span></span> | <span data-ttu-id="60c40-834">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="60c40-834">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="60c40-835">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="60c40-835">PhoneNumber</span></span> | <span data-ttu-id="60c40-836">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="60c40-836">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="60c40-837">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="60c40-837">TaskSuggestion</span></span> | <span data-ttu-id="60c40-838">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="60c40-838">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="60c40-839">String</span><span class="sxs-lookup"><span data-stu-id="60c40-839">String</span></span> | <span data-ttu-id="60c40-840">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="60c40-840">**Restricted**</span></span> |

<span data-ttu-id="60c40-841">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="60c40-841">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="60c40-842">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-842">Example</span></span>

<span data-ttu-id="60c40-843">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-843">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="60c40-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="60c40-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="60c40-845">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="60c40-845">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-846">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="60c40-846">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="60c40-847">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="60c40-847">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-848">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-848">Parameters</span></span>

|<span data-ttu-id="60c40-849">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-849">Name</span></span>| <span data-ttu-id="60c40-850">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-850">Type</span></span>| <span data-ttu-id="60c40-851">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-851">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="60c40-852">String</span><span class="sxs-lookup"><span data-stu-id="60c40-852">String</span></span>|<span data-ttu-id="60c40-853">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="60c40-853">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="60c40-854">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-854">Requirements</span></span>

|<span data-ttu-id="60c40-855">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-855">Requirement</span></span>| <span data-ttu-id="60c40-856">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-857">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-858">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-858">1.0</span></span>|
|[<span data-ttu-id="60c40-859">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-859">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-860">ReadItem</span></span>|
|[<span data-ttu-id="60c40-861">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-861">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-862">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-862">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="60c40-863">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="60c40-863">Returns:</span></span>

<span data-ttu-id="60c40-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="60c40-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="60c40-866">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="60c40-866">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="60c40-867">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="60c40-867">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="60c40-868">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="60c40-868">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-869">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="60c40-869">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="60c40-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="60c40-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="60c40-873">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="60c40-873">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="60c40-874">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="60c40-874">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="60c40-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="60c40-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="60c40-877">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-877">Requirements</span></span>

|<span data-ttu-id="60c40-878">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-878">Requirement</span></span>| <span data-ttu-id="60c40-879">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-879">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-880">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-880">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-881">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-881">1.0</span></span>|
|[<span data-ttu-id="60c40-882">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-882">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-883">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-883">ReadItem</span></span>|
|[<span data-ttu-id="60c40-884">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-884">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-885">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-885">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="60c40-886">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="60c40-886">Returns:</span></span>

<span data-ttu-id="60c40-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="60c40-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="60c40-889">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="60c40-889">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="60c40-890">Object</span><span class="sxs-lookup"><span data-stu-id="60c40-890">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="60c40-891">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-891">Example</span></span>

<span data-ttu-id="60c40-892">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="60c40-892">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="60c40-893">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="60c40-893">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="60c40-894">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="60c40-894">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="60c40-895">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="60c40-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="60c40-896">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="60c40-896">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="60c40-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="60c40-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-899">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-899">Parameters</span></span>

|<span data-ttu-id="60c40-900">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-900">Name</span></span>| <span data-ttu-id="60c40-901">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-901">Type</span></span>| <span data-ttu-id="60c40-902">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-902">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="60c40-903">String</span><span class="sxs-lookup"><span data-stu-id="60c40-903">String</span></span>|<span data-ttu-id="60c40-904">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="60c40-904">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="60c40-905">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-905">Requirements</span></span>

|<span data-ttu-id="60c40-906">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-906">Requirement</span></span>| <span data-ttu-id="60c40-907">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-907">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-908">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-908">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-909">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-909">1.0</span></span>|
|[<span data-ttu-id="60c40-910">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-910">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-911">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-911">ReadItem</span></span>|
|[<span data-ttu-id="60c40-912">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-912">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-913">Чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-913">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="60c40-914">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="60c40-914">Returns:</span></span>

<span data-ttu-id="60c40-915">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="60c40-915">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="60c40-916">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="60c40-916">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="60c40-917">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="60c40-917">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="60c40-918">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-918">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="60c40-919">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="60c40-919">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="60c40-920">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-920">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="60c40-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="60c40-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-923">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-923">Parameters</span></span>

|<span data-ttu-id="60c40-924">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-924">Name</span></span>| <span data-ttu-id="60c40-925">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-925">Type</span></span>| <span data-ttu-id="60c40-926">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="60c40-926">Attributes</span></span>| <span data-ttu-id="60c40-927">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-927">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="60c40-928">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="60c40-928">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="60c40-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="60c40-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="60c40-932">Объект</span><span class="sxs-lookup"><span data-stu-id="60c40-932">Object</span></span>| <span data-ttu-id="60c40-933">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-933">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-934">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="60c40-934">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="60c40-935">Объект</span><span class="sxs-lookup"><span data-stu-id="60c40-935">Object</span></span>| <span data-ttu-id="60c40-936">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-936">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-937">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="60c40-937">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="60c40-938">функция</span><span class="sxs-lookup"><span data-stu-id="60c40-938">function</span></span>||<span data-ttu-id="60c40-939">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="60c40-939">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="60c40-940">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="60c40-940">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="60c40-941">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="60c40-941">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="60c40-942">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-942">Requirements</span></span>

|<span data-ttu-id="60c40-943">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-943">Requirement</span></span>| <span data-ttu-id="60c40-944">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-944">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-945">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-945">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-946">1.2</span><span class="sxs-lookup"><span data-stu-id="60c40-946">1.2</span></span>|
|[<span data-ttu-id="60c40-947">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-947">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-948">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="60c40-948">ReadWriteItem</span></span>|
|[<span data-ttu-id="60c40-949">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-949">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-950">Создание</span><span class="sxs-lookup"><span data-stu-id="60c40-950">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="60c40-951">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="60c40-951">Returns:</span></span>

<span data-ttu-id="60c40-952">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="60c40-952">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="60c40-953">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="60c40-953">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="60c40-954">String</span><span class="sxs-lookup"><span data-stu-id="60c40-954">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="60c40-955">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-955">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="60c40-956">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="60c40-956">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="60c40-957">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="60c40-957">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="60c40-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="60c40-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-961">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-961">Parameters</span></span>

|<span data-ttu-id="60c40-962">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-962">Name</span></span>| <span data-ttu-id="60c40-963">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-963">Type</span></span>| <span data-ttu-id="60c40-964">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="60c40-964">Attributes</span></span>| <span data-ttu-id="60c40-965">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-965">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="60c40-966">function</span><span class="sxs-lookup"><span data-stu-id="60c40-966">function</span></span>||<span data-ttu-id="60c40-967">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="60c40-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="60c40-968">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="60c40-968">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="60c40-969">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="60c40-969">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="60c40-970">Объект</span><span class="sxs-lookup"><span data-stu-id="60c40-970">Object</span></span>| <span data-ttu-id="60c40-971">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-971">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-972">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="60c40-972">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="60c40-973">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="60c40-973">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="60c40-974">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-974">Requirements</span></span>

|<span data-ttu-id="60c40-975">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-975">Requirement</span></span>| <span data-ttu-id="60c40-976">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-976">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-977">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-977">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-978">1.0</span><span class="sxs-lookup"><span data-stu-id="60c40-978">1.0</span></span>|
|[<span data-ttu-id="60c40-979">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-979">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-980">ReadItem</span><span class="sxs-lookup"><span data-stu-id="60c40-980">ReadItem</span></span>|
|[<span data-ttu-id="60c40-981">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-981">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-982">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="60c40-982">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-983">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-983">Example</span></span>

<span data-ttu-id="60c40-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="60c40-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="60c40-987">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="60c40-987">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="60c40-988">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="60c40-988">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="60c40-989">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="60c40-989">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="60c40-990">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="60c40-990">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="60c40-991">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="60c40-991">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="60c40-992">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="60c40-992">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-993">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-993">Parameters</span></span>

|<span data-ttu-id="60c40-994">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-994">Name</span></span>| <span data-ttu-id="60c40-995">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-995">Type</span></span>| <span data-ttu-id="60c40-996">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="60c40-996">Attributes</span></span>| <span data-ttu-id="60c40-997">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-997">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="60c40-998">String</span><span class="sxs-lookup"><span data-stu-id="60c40-998">String</span></span>||<span data-ttu-id="60c40-999">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="60c40-999">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="60c40-1000">Объект</span><span class="sxs-lookup"><span data-stu-id="60c40-1000">Object</span></span>| <span data-ttu-id="60c40-1001">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-1002">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="60c40-1002">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="60c40-1003">Объект</span><span class="sxs-lookup"><span data-stu-id="60c40-1003">Object</span></span>| <span data-ttu-id="60c40-1004">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-1005">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="60c40-1005">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="60c40-1006">функция</span><span class="sxs-lookup"><span data-stu-id="60c40-1006">function</span></span>| <span data-ttu-id="60c40-1007">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-1008">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="60c40-1008">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="60c40-1009">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="60c40-1009">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="60c40-1010">Ошибки</span><span class="sxs-lookup"><span data-stu-id="60c40-1010">Errors</span></span>

| <span data-ttu-id="60c40-1011">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="60c40-1011">Error code</span></span> | <span data-ttu-id="60c40-1012">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-1012">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="60c40-1013">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="60c40-1013">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="60c40-1014">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-1014">Requirements</span></span>

|<span data-ttu-id="60c40-1015">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-1015">Requirement</span></span>| <span data-ttu-id="60c40-1016">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-1016">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-1017">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="60c40-1017">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-1018">1.1</span><span class="sxs-lookup"><span data-stu-id="60c40-1018">1.1</span></span>|
|[<span data-ttu-id="60c40-1019">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-1019">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-1020">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="60c40-1020">ReadWriteItem</span></span>|
|[<span data-ttu-id="60c40-1021">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-1021">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-1022">Создание</span><span class="sxs-lookup"><span data-stu-id="60c40-1022">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-1023">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-1023">Example</span></span>

<span data-ttu-id="60c40-1024">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="60c40-1024">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="60c40-1025">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="60c40-1025">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="60c40-1026">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="60c40-1026">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="60c40-p166">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="60c40-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="60c40-1030">Параметры</span><span class="sxs-lookup"><span data-stu-id="60c40-1030">Parameters</span></span>

|<span data-ttu-id="60c40-1031">Имя</span><span class="sxs-lookup"><span data-stu-id="60c40-1031">Name</span></span>| <span data-ttu-id="60c40-1032">Тип</span><span class="sxs-lookup"><span data-stu-id="60c40-1032">Type</span></span>| <span data-ttu-id="60c40-1033">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="60c40-1033">Attributes</span></span>| <span data-ttu-id="60c40-1034">Описание</span><span class="sxs-lookup"><span data-stu-id="60c40-1034">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="60c40-1035">String</span><span class="sxs-lookup"><span data-stu-id="60c40-1035">String</span></span>||<span data-ttu-id="60c40-p167">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="60c40-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="60c40-1039">Object</span><span class="sxs-lookup"><span data-stu-id="60c40-1039">Object</span></span>| <span data-ttu-id="60c40-1040">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-1041">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="60c40-1041">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="60c40-1042">Объект</span><span class="sxs-lookup"><span data-stu-id="60c40-1042">Object</span></span>| <span data-ttu-id="60c40-1043">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-1044">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="60c40-1044">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="60c40-1045">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="60c40-1045">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="60c40-1046">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="60c40-1046">&lt;optional&gt;</span></span>|<span data-ttu-id="60c40-1047">Если `text`текущий стиль применяется в Outlook для веб-клиентов и клиентов для настольных ПК.</span><span class="sxs-lookup"><span data-stu-id="60c40-1047">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="60c40-1048">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="60c40-1048">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="60c40-1049">Если `html` и поле поддерживает HTML (тема не используется), текущий стиль применяется в Outlook в Интернете, а в настольных клиентах Outlook применяется стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="60c40-1049">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="60c40-1050">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="60c40-1050">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="60c40-1051">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="60c40-1051">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="60c40-1052">функция</span><span class="sxs-lookup"><span data-stu-id="60c40-1052">function</span></span>||<span data-ttu-id="60c40-1053">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="60c40-1053">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="60c40-1054">Требования</span><span class="sxs-lookup"><span data-stu-id="60c40-1054">Requirements</span></span>

|<span data-ttu-id="60c40-1055">Требование</span><span class="sxs-lookup"><span data-stu-id="60c40-1055">Requirement</span></span>| <span data-ttu-id="60c40-1056">Значение</span><span class="sxs-lookup"><span data-stu-id="60c40-1056">Value</span></span>|
|---|---|
|[<span data-ttu-id="60c40-1057">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="60c40-1057">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60c40-1058">1.2</span><span class="sxs-lookup"><span data-stu-id="60c40-1058">1.2</span></span>|
|[<span data-ttu-id="60c40-1059">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="60c40-1059">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60c40-1060">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="60c40-1060">ReadWriteItem</span></span>|
|[<span data-ttu-id="60c40-1061">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="60c40-1061">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60c40-1062">Создание</span><span class="sxs-lookup"><span data-stu-id="60c40-1062">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="60c40-1063">Пример</span><span class="sxs-lookup"><span data-stu-id="60c40-1063">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
