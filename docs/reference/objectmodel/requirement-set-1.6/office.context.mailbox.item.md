---
title: Office. Context. Mailbox. Item — набор требований 1,6
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: f5789037ab5486fecf6e821dc39dc4b627e7f825
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001588"
---
# <a name="item"></a><span data-ttu-id="68dae-102">item</span><span class="sxs-lookup"><span data-stu-id="68dae-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="68dae-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="68dae-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="68dae-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="68dae-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-106">Requirements</span></span>

|<span data-ttu-id="68dae-107">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-107">Requirement</span></span>| <span data-ttu-id="68dae-108">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-110">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-110">1.0</span></span>|
|[<span data-ttu-id="68dae-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="68dae-112">Restricted</span></span>|
|[<span data-ttu-id="68dae-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="68dae-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="68dae-115">Members and methods</span></span>

| <span data-ttu-id="68dae-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-116">Member</span></span> | <span data-ttu-id="68dae-117">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="68dae-118">attachments</span><span class="sxs-lookup"><span data-stu-id="68dae-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="68dae-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-119">Member</span></span> |
| [<span data-ttu-id="68dae-120">bcc</span><span class="sxs-lookup"><span data-stu-id="68dae-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="68dae-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-121">Member</span></span> |
| [<span data-ttu-id="68dae-122">body</span><span class="sxs-lookup"><span data-stu-id="68dae-122">body</span></span>](#body-body) | <span data-ttu-id="68dae-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-123">Member</span></span> |
| [<span data-ttu-id="68dae-124">cc</span><span class="sxs-lookup"><span data-stu-id="68dae-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="68dae-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-125">Member</span></span> |
| [<span data-ttu-id="68dae-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="68dae-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="68dae-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-127">Member</span></span> |
| [<span data-ttu-id="68dae-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="68dae-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="68dae-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-129">Member</span></span> |
| [<span data-ttu-id="68dae-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="68dae-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="68dae-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-131">Member</span></span> |
| [<span data-ttu-id="68dae-132">end</span><span class="sxs-lookup"><span data-stu-id="68dae-132">end</span></span>](#end-datetime) | <span data-ttu-id="68dae-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-133">Member</span></span> |
| [<span data-ttu-id="68dae-134">from</span><span class="sxs-lookup"><span data-stu-id="68dae-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="68dae-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-135">Member</span></span> |
| [<span data-ttu-id="68dae-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="68dae-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="68dae-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-137">Member</span></span> |
| [<span data-ttu-id="68dae-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="68dae-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="68dae-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-139">Member</span></span> |
| [<span data-ttu-id="68dae-140">itemId</span><span class="sxs-lookup"><span data-stu-id="68dae-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="68dae-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-141">Member</span></span> |
| [<span data-ttu-id="68dae-142">itemType</span><span class="sxs-lookup"><span data-stu-id="68dae-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="68dae-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-143">Member</span></span> |
| [<span data-ttu-id="68dae-144">location</span><span class="sxs-lookup"><span data-stu-id="68dae-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="68dae-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-145">Member</span></span> |
| [<span data-ttu-id="68dae-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="68dae-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="68dae-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-147">Member</span></span> |
| [<span data-ttu-id="68dae-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="68dae-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="68dae-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-149">Member</span></span> |
| [<span data-ttu-id="68dae-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="68dae-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="68dae-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-151">Member</span></span> |
| [<span data-ttu-id="68dae-152">organizer</span><span class="sxs-lookup"><span data-stu-id="68dae-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="68dae-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-153">Member</span></span> |
| [<span data-ttu-id="68dae-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="68dae-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="68dae-155">Member</span><span class="sxs-lookup"><span data-stu-id="68dae-155">Member</span></span> |
| [<span data-ttu-id="68dae-156">sender</span><span class="sxs-lookup"><span data-stu-id="68dae-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="68dae-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-157">Member</span></span> |
| [<span data-ttu-id="68dae-158">start</span><span class="sxs-lookup"><span data-stu-id="68dae-158">start</span></span>](#start-datetime) | <span data-ttu-id="68dae-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-159">Member</span></span> |
| [<span data-ttu-id="68dae-160">subject</span><span class="sxs-lookup"><span data-stu-id="68dae-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="68dae-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-161">Member</span></span> |
| [<span data-ttu-id="68dae-162">to</span><span class="sxs-lookup"><span data-stu-id="68dae-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="68dae-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="68dae-163">Member</span></span> |
| [<span data-ttu-id="68dae-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="68dae-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="68dae-165">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-165">Method</span></span> |
| [<span data-ttu-id="68dae-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="68dae-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="68dae-167">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-167">Method</span></span> |
| [<span data-ttu-id="68dae-168">close</span><span class="sxs-lookup"><span data-stu-id="68dae-168">close</span></span>](#close) | <span data-ttu-id="68dae-169">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-169">Method</span></span> |
| [<span data-ttu-id="68dae-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="68dae-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="68dae-171">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-171">Method</span></span> |
| [<span data-ttu-id="68dae-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="68dae-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="68dae-173">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-173">Method</span></span> |
| [<span data-ttu-id="68dae-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="68dae-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="68dae-175">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-175">Method</span></span> |
| [<span data-ttu-id="68dae-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="68dae-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="68dae-177">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-177">Method</span></span> |
| [<span data-ttu-id="68dae-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="68dae-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="68dae-179">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-179">Method</span></span> |
| [<span data-ttu-id="68dae-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="68dae-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="68dae-181">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-181">Method</span></span> |
| [<span data-ttu-id="68dae-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="68dae-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="68dae-183">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-183">Method</span></span> |
| [<span data-ttu-id="68dae-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="68dae-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="68dae-185">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-185">Method</span></span> |
| [<span data-ttu-id="68dae-186">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="68dae-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="68dae-187">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-187">Method</span></span> |
| [<span data-ttu-id="68dae-188">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="68dae-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="68dae-189">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-189">Method</span></span> |
| [<span data-ttu-id="68dae-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="68dae-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="68dae-191">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-191">Method</span></span> |
| [<span data-ttu-id="68dae-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="68dae-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="68dae-193">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-193">Method</span></span> |
| [<span data-ttu-id="68dae-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="68dae-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="68dae-195">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-195">Method</span></span> |
| [<span data-ttu-id="68dae-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="68dae-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="68dae-197">Метод</span><span class="sxs-lookup"><span data-stu-id="68dae-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="68dae-198">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-198">Example</span></span>

<span data-ttu-id="68dae-199">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="68dae-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="68dae-200">Members</span><span class="sxs-lookup"><span data-stu-id="68dae-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="68dae-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="68dae-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="68dae-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-204">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="68dae-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="68dae-205">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="68dae-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-206">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-206">Type</span></span>

*   <span data-ttu-id="68dae-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="68dae-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-208">Requirements</span></span>

|<span data-ttu-id="68dae-209">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-209">Requirement</span></span>| <span data-ttu-id="68dae-210">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-211">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-212">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-212">1.0</span></span>|
|[<span data-ttu-id="68dae-213">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-214">ReadItem</span></span>|
|[<span data-ttu-id="68dae-215">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-216">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-217">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-217">Example</span></span>

<span data-ttu-id="68dae-218">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="68dae-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-220">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="68dae-221">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="68dae-221">Compose mode only.</span></span>

<span data-ttu-id="68dae-222">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-222">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="68dae-223">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="68dae-223">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="68dae-224">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-224">Get 500 members maximum.</span></span>
- <span data-ttu-id="68dae-225">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-225">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-226">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-226">Type</span></span>

*   [<span data-ttu-id="68dae-227">Получатели</span><span class="sxs-lookup"><span data-stu-id="68dae-227">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="68dae-228">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-228">Requirements</span></span>

|<span data-ttu-id="68dae-229">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-229">Requirement</span></span>| <span data-ttu-id="68dae-230">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-231">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-232">1.1</span><span class="sxs-lookup"><span data-stu-id="68dae-232">1.1</span></span>|
|[<span data-ttu-id="68dae-233">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-233">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-234">ReadItem</span></span>|
|[<span data-ttu-id="68dae-235">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-236">Создание</span><span class="sxs-lookup"><span data-stu-id="68dae-236">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-237">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-237">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="68dae-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-239">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-239">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-240">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-240">Type</span></span>

*   [<span data-ttu-id="68dae-241">Body</span><span class="sxs-lookup"><span data-stu-id="68dae-241">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="68dae-242">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-242">Requirements</span></span>

|<span data-ttu-id="68dae-243">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-243">Requirement</span></span>| <span data-ttu-id="68dae-244">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-245">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-246">1.1</span><span class="sxs-lookup"><span data-stu-id="68dae-246">1.1</span></span>|
|[<span data-ttu-id="68dae-247">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-248">ReadItem</span></span>|
|[<span data-ttu-id="68dae-249">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-250">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-251">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-251">Example</span></span>

<span data-ttu-id="68dae-252">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="68dae-252">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="68dae-253">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="68dae-253">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="68dae-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-255">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="68dae-256">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="68dae-257">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="68dae-257">Read mode</span></span>

<span data-ttu-id="68dae-258">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-258">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="68dae-259">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="68dae-260">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="68dae-260">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="68dae-261">Режим создания</span><span class="sxs-lookup"><span data-stu-id="68dae-261">Compose mode</span></span>

<span data-ttu-id="68dae-262">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="68dae-263">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-263">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="68dae-264">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="68dae-264">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="68dae-265">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-265">Get 500 members maximum.</span></span>
- <span data-ttu-id="68dae-266">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-266">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="68dae-267">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-267">Type</span></span>

*   <span data-ttu-id="68dae-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-269">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-269">Requirements</span></span>

|<span data-ttu-id="68dae-270">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-270">Requirement</span></span>| <span data-ttu-id="68dae-271">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-272">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-272">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-273">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-273">1.0</span></span>|
|[<span data-ttu-id="68dae-274">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-274">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-275">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-275">ReadItem</span></span>|
|[<span data-ttu-id="68dae-276">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-276">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-277">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-277">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="68dae-278">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="68dae-278">(nullable) conversationId: String</span></span>

<span data-ttu-id="68dae-279">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="68dae-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="68dae-p109">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="68dae-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="68dae-p110">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="68dae-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-284">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-284">Type</span></span>

*   <span data-ttu-id="68dae-285">String</span><span class="sxs-lookup"><span data-stu-id="68dae-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-286">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-286">Requirements</span></span>

|<span data-ttu-id="68dae-287">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-287">Requirement</span></span>| <span data-ttu-id="68dae-288">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-289">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-289">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-290">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-290">1.0</span></span>|
|[<span data-ttu-id="68dae-291">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-291">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-292">ReadItem</span></span>|
|[<span data-ttu-id="68dae-293">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-293">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-294">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-294">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-295">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-295">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="68dae-296">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="68dae-296">dateTimeCreated: Date</span></span>

<span data-ttu-id="68dae-p111">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-299">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-299">Type</span></span>

*   <span data-ttu-id="68dae-300">Дата</span><span class="sxs-lookup"><span data-stu-id="68dae-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-301">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-301">Requirements</span></span>

|<span data-ttu-id="68dae-302">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-302">Requirement</span></span>| <span data-ttu-id="68dae-303">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-304">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-305">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-305">1.0</span></span>|
|[<span data-ttu-id="68dae-306">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-307">ReadItem</span></span>|
|[<span data-ttu-id="68dae-308">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-309">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-310">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-310">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="68dae-311">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="68dae-311">dateTimeModified: Date</span></span>

<span data-ttu-id="68dae-p112">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-314">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-314">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-315">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-315">Type</span></span>

*   <span data-ttu-id="68dae-316">Дата</span><span class="sxs-lookup"><span data-stu-id="68dae-316">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-317">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-317">Requirements</span></span>

|<span data-ttu-id="68dae-318">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-318">Requirement</span></span>| <span data-ttu-id="68dae-319">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-320">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-321">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-321">1.0</span></span>|
|[<span data-ttu-id="68dae-322">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-323">ReadItem</span></span>|
|[<span data-ttu-id="68dae-324">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-325">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-325">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-326">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-326">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="68dae-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-328">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-328">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="68dae-p113">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="68dae-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="68dae-331">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="68dae-331">Read mode</span></span>

<span data-ttu-id="68dae-332">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="68dae-332">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="68dae-333">Режим создания</span><span class="sxs-lookup"><span data-stu-id="68dae-333">Compose mode</span></span>

<span data-ttu-id="68dae-334">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="68dae-334">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="68dae-335">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="68dae-335">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="68dae-336">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-336">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="68dae-337">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-337">Type</span></span>

*   <span data-ttu-id="68dae-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-339">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-339">Requirements</span></span>

|<span data-ttu-id="68dae-340">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-340">Requirement</span></span>| <span data-ttu-id="68dae-341">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-342">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-343">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-343">1.0</span></span>|
|[<span data-ttu-id="68dae-344">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-344">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-345">ReadItem</span></span>|
|[<span data-ttu-id="68dae-346">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-346">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-347">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-347">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="68dae-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-p114">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="68dae-p115">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="68dae-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-353">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="68dae-353">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-354">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-354">Type</span></span>

*   [<span data-ttu-id="68dae-355">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="68dae-355">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="68dae-356">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="68dae-357">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-357">Requirements</span></span>

|<span data-ttu-id="68dae-358">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-358">Requirement</span></span>| <span data-ttu-id="68dae-359">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-360">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-361">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-361">1.0</span></span>|
|[<span data-ttu-id="68dae-362">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-363">ReadItem</span></span>|
|[<span data-ttu-id="68dae-364">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-365">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-365">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="68dae-366">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="68dae-366">internetMessageId: String</span></span>

<span data-ttu-id="68dae-p116">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-369">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-369">Type</span></span>

*   <span data-ttu-id="68dae-370">String</span><span class="sxs-lookup"><span data-stu-id="68dae-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-371">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-371">Requirements</span></span>

|<span data-ttu-id="68dae-372">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-372">Requirement</span></span>| <span data-ttu-id="68dae-373">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-374">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-375">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-375">1.0</span></span>|
|[<span data-ttu-id="68dae-376">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-376">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-377">ReadItem</span></span>|
|[<span data-ttu-id="68dae-378">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-378">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-379">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-380">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-380">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="68dae-381">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="68dae-381">itemClass: String</span></span>

<span data-ttu-id="68dae-p117">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="68dae-p118">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="68dae-386">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-386">Type</span></span> | <span data-ttu-id="68dae-387">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-387">Description</span></span> | <span data-ttu-id="68dae-388">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="68dae-388">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="68dae-389">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="68dae-389">Appointment items</span></span> | <span data-ttu-id="68dae-390">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="68dae-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="68dae-391">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="68dae-391">Message items</span></span> | <span data-ttu-id="68dae-392">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="68dae-393">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="68dae-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-394">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-394">Type</span></span>

*   <span data-ttu-id="68dae-395">String</span><span class="sxs-lookup"><span data-stu-id="68dae-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-396">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-396">Requirements</span></span>

|<span data-ttu-id="68dae-397">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-397">Requirement</span></span>| <span data-ttu-id="68dae-398">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-399">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-400">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-400">1.0</span></span>|
|[<span data-ttu-id="68dae-401">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-401">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-402">ReadItem</span></span>|
|[<span data-ttu-id="68dae-403">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-403">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-404">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-405">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-405">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="68dae-406">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="68dae-406">(nullable) itemId: String</span></span>

<span data-ttu-id="68dae-407">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-407">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="68dae-408">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-408">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-409">Идентификатор, возвращаемый `itemId` свойством, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="68dae-409">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="68dae-410">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="68dae-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="68dae-411">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="68dae-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="68dae-412">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="68dae-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="68dae-p121">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-415">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-415">Type</span></span>

*   <span data-ttu-id="68dae-416">String</span><span class="sxs-lookup"><span data-stu-id="68dae-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-417">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-417">Requirements</span></span>

|<span data-ttu-id="68dae-418">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-418">Requirement</span></span>| <span data-ttu-id="68dae-419">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-420">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-421">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-421">1.0</span></span>|
|[<span data-ttu-id="68dae-422">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-422">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-423">ReadItem</span></span>|
|[<span data-ttu-id="68dae-424">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-424">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-425">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-426">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-426">Example</span></span>

<span data-ttu-id="68dae-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="68dae-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-430">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="68dae-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="68dae-431">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="68dae-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-432">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-432">Type</span></span>

*   [<span data-ttu-id="68dae-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="68dae-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="68dae-434">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-434">Requirements</span></span>

|<span data-ttu-id="68dae-435">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-435">Requirement</span></span>| <span data-ttu-id="68dae-436">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-437">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-438">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-438">1.0</span></span>|
|[<span data-ttu-id="68dae-439">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-439">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-440">ReadItem</span></span>|
|[<span data-ttu-id="68dae-441">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-441">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-442">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-442">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-443">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-443">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="68dae-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-445">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="68dae-446">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="68dae-446">Read mode</span></span>

<span data-ttu-id="68dae-447">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-447">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="68dae-448">Режим создания</span><span class="sxs-lookup"><span data-stu-id="68dae-448">Compose mode</span></span>

<span data-ttu-id="68dae-449">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="68dae-450">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-450">Type</span></span>

*   <span data-ttu-id="68dae-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-452">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-452">Requirements</span></span>

|<span data-ttu-id="68dae-453">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-453">Requirement</span></span>| <span data-ttu-id="68dae-454">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-455">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-456">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-456">1.0</span></span>|
|[<span data-ttu-id="68dae-457">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-457">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-458">ReadItem</span></span>|
|[<span data-ttu-id="68dae-459">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-459">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-460">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-460">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="68dae-461">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="68dae-461">normalizedSubject: String</span></span>

<span data-ttu-id="68dae-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="68dae-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="68dae-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-466">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-466">Type</span></span>

*   <span data-ttu-id="68dae-467">String</span><span class="sxs-lookup"><span data-stu-id="68dae-467">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-468">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-468">Requirements</span></span>

|<span data-ttu-id="68dae-469">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-469">Requirement</span></span>| <span data-ttu-id="68dae-470">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-470">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-471">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-471">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-472">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-472">1.0</span></span>|
|[<span data-ttu-id="68dae-473">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-473">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-474">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-474">ReadItem</span></span>|
|[<span data-ttu-id="68dae-475">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-475">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-476">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-476">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-477">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-477">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="68dae-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-479">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-479">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-480">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-480">Type</span></span>

*   [<span data-ttu-id="68dae-481">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="68dae-481">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="68dae-482">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-482">Requirements</span></span>

|<span data-ttu-id="68dae-483">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-483">Requirement</span></span>| <span data-ttu-id="68dae-484">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-485">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-486">1.3</span><span class="sxs-lookup"><span data-stu-id="68dae-486">1.3</span></span>|
|[<span data-ttu-id="68dae-487">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-488">ReadItem</span></span>|
|[<span data-ttu-id="68dae-489">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-490">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-490">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-491">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-491">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="68dae-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-493">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="68dae-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="68dae-494">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="68dae-495">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="68dae-495">Read mode</span></span>

<span data-ttu-id="68dae-496">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="68dae-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="68dae-497">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="68dae-498">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="68dae-498">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="68dae-499">Режим создания</span><span class="sxs-lookup"><span data-stu-id="68dae-499">Compose mode</span></span>

<span data-ttu-id="68dae-500">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="68dae-500">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="68dae-501">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-501">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="68dae-502">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="68dae-502">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="68dae-503">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-503">Get 500 members maximum.</span></span>
- <span data-ttu-id="68dae-504">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-504">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="68dae-505">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-505">Type</span></span>

*   <span data-ttu-id="68dae-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-507">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-507">Requirements</span></span>

|<span data-ttu-id="68dae-508">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-508">Requirement</span></span>| <span data-ttu-id="68dae-509">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-510">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-511">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-511">1.0</span></span>|
|[<span data-ttu-id="68dae-512">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-513">ReadItem</span></span>|
|[<span data-ttu-id="68dae-514">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-514">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-515">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-515">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="68dae-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-p128">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-519">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-519">Type</span></span>

*   [<span data-ttu-id="68dae-520">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="68dae-520">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="68dae-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-521">Requirements</span></span>

|<span data-ttu-id="68dae-522">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-522">Requirement</span></span>| <span data-ttu-id="68dae-523">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-524">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-525">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-525">1.0</span></span>|
|[<span data-ttu-id="68dae-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-527">ReadItem</span></span>|
|[<span data-ttu-id="68dae-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-529">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-529">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-530">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-530">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="68dae-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-532">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="68dae-532">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="68dae-533">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-533">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="68dae-534">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="68dae-534">Read mode</span></span>

<span data-ttu-id="68dae-535">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="68dae-535">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="68dae-536">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="68dae-537">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="68dae-537">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="68dae-538">Режим создания</span><span class="sxs-lookup"><span data-stu-id="68dae-538">Compose mode</span></span>

<span data-ttu-id="68dae-539">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="68dae-539">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="68dae-540">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-540">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="68dae-541">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="68dae-541">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="68dae-542">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-542">Get 500 members maximum.</span></span>
- <span data-ttu-id="68dae-543">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-543">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="68dae-544">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-544">Type</span></span>

*   <span data-ttu-id="68dae-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-546">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-546">Requirements</span></span>

|<span data-ttu-id="68dae-547">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-547">Requirement</span></span>| <span data-ttu-id="68dae-548">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-549">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-550">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-550">1.0</span></span>|
|[<span data-ttu-id="68dae-551">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-552">ReadItem</span></span>|
|[<span data-ttu-id="68dae-553">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-554">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-554">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="68dae-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-p132">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="68dae-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="68dae-p133">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="68dae-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-560">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="68dae-560">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="68dae-561">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-561">Type</span></span>

*   [<span data-ttu-id="68dae-562">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="68dae-562">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="68dae-563">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-563">Requirements</span></span>

|<span data-ttu-id="68dae-564">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-564">Requirement</span></span>| <span data-ttu-id="68dae-565">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-565">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-566">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-566">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-567">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-567">1.0</span></span>|
|[<span data-ttu-id="68dae-568">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-568">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-569">ReadItem</span></span>|
|[<span data-ttu-id="68dae-570">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-570">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-571">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-571">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-572">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-572">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="68dae-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-574">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-574">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="68dae-p134">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="68dae-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="68dae-577">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="68dae-577">Read mode</span></span>

<span data-ttu-id="68dae-578">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="68dae-578">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="68dae-579">Режим создания</span><span class="sxs-lookup"><span data-stu-id="68dae-579">Compose mode</span></span>

<span data-ttu-id="68dae-580">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="68dae-580">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="68dae-581">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="68dae-581">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="68dae-582">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="68dae-582">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="68dae-583">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-583">Type</span></span>

*   <span data-ttu-id="68dae-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-585">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-585">Requirements</span></span>

|<span data-ttu-id="68dae-586">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-586">Requirement</span></span>| <span data-ttu-id="68dae-587">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-588">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-589">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-589">1.0</span></span>|
|[<span data-ttu-id="68dae-590">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-591">ReadItem</span></span>|
|[<span data-ttu-id="68dae-592">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-593">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-593">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="68dae-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-595">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-595">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="68dae-596">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="68dae-596">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="68dae-597">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="68dae-597">Read mode</span></span>

<span data-ttu-id="68dae-p135">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="68dae-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="68dae-600">Режим создания</span><span class="sxs-lookup"><span data-stu-id="68dae-600">Compose mode</span></span>

<span data-ttu-id="68dae-601">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="68dae-601">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="68dae-602">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-602">Type</span></span>

*   <span data-ttu-id="68dae-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-604">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-604">Requirements</span></span>

|<span data-ttu-id="68dae-605">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-605">Requirement</span></span>| <span data-ttu-id="68dae-606">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-607">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-608">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-608">1.0</span></span>|
|[<span data-ttu-id="68dae-609">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-610">ReadItem</span></span>|
|[<span data-ttu-id="68dae-611">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-612">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-612">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="68dae-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="68dae-614">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-614">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="68dae-615">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-615">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="68dae-616">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="68dae-616">Read mode</span></span>

<span data-ttu-id="68dae-617">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-617">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="68dae-618">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="68dae-619">Тем не менее, в Windows и Mac вы можете настроить максимальную длину участников 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-619">However, on Windows and Mac, you can set up to get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="68dae-620">Режим создания</span><span class="sxs-lookup"><span data-stu-id="68dae-620">Compose mode</span></span>

<span data-ttu-id="68dae-621">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-621">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="68dae-622">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-622">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="68dae-623">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="68dae-623">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="68dae-624">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-624">Get 500 members maximum.</span></span>
- <span data-ttu-id="68dae-625">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="68dae-625">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="68dae-626">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-626">Type</span></span>

*   <span data-ttu-id="68dae-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-628">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-628">Requirements</span></span>

|<span data-ttu-id="68dae-629">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-629">Requirement</span></span>| <span data-ttu-id="68dae-630">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-630">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-631">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-631">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-632">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-632">1.0</span></span>|
|[<span data-ttu-id="68dae-633">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-633">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-634">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-634">ReadItem</span></span>|
|[<span data-ttu-id="68dae-635">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-635">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-636">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-636">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="68dae-637">Методы</span><span class="sxs-lookup"><span data-stu-id="68dae-637">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="68dae-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="68dae-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="68dae-639">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="68dae-639">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="68dae-640">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="68dae-640">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="68dae-641">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="68dae-641">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-642">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-642">Parameters</span></span>

|<span data-ttu-id="68dae-643">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-643">Name</span></span>| <span data-ttu-id="68dae-644">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-644">Type</span></span>| <span data-ttu-id="68dae-645">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="68dae-645">Attributes</span></span>| <span data-ttu-id="68dae-646">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-646">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="68dae-647">String</span><span class="sxs-lookup"><span data-stu-id="68dae-647">String</span></span>||<span data-ttu-id="68dae-p139">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="68dae-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="68dae-650">String</span><span class="sxs-lookup"><span data-stu-id="68dae-650">String</span></span>||<span data-ttu-id="68dae-p140">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="68dae-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="68dae-653">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-653">Object</span></span>| <span data-ttu-id="68dae-654">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-654">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-655">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="68dae-655">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="68dae-656">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-656">Object</span></span> | <span data-ttu-id="68dae-657">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-657">&lt;optional&gt;</span></span> | <span data-ttu-id="68dae-658">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="68dae-658">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="68dae-659">Boolean</span><span class="sxs-lookup"><span data-stu-id="68dae-659">Boolean</span></span> | <span data-ttu-id="68dae-660">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-660">&lt;optional&gt;</span></span> | <span data-ttu-id="68dae-661">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="68dae-661">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="68dae-662">function</span><span class="sxs-lookup"><span data-stu-id="68dae-662">function</span></span>| <span data-ttu-id="68dae-663">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-663">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-664">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="68dae-664">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="68dae-665">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="68dae-665">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="68dae-666">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="68dae-666">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="68dae-667">Ошибки</span><span class="sxs-lookup"><span data-stu-id="68dae-667">Errors</span></span>

| <span data-ttu-id="68dae-668">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="68dae-668">Error code</span></span> | <span data-ttu-id="68dae-669">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-669">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="68dae-670">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="68dae-670">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="68dae-671">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="68dae-671">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="68dae-672">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="68dae-672">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="68dae-673">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-673">Requirements</span></span>

|<span data-ttu-id="68dae-674">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-674">Requirement</span></span>| <span data-ttu-id="68dae-675">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-676">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-677">1.1</span><span class="sxs-lookup"><span data-stu-id="68dae-677">1.1</span></span>|
|[<span data-ttu-id="68dae-678">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-679">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="68dae-679">ReadWriteItem</span></span>|
|[<span data-ttu-id="68dae-680">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-681">Создание</span><span class="sxs-lookup"><span data-stu-id="68dae-681">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="68dae-682">Примеры</span><span class="sxs-lookup"><span data-stu-id="68dae-682">Examples</span></span>

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

<span data-ttu-id="68dae-683">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-683">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="68dae-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="68dae-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="68dae-685">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="68dae-685">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="68dae-p141">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="68dae-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="68dae-689">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="68dae-689">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="68dae-690">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="68dae-690">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-691">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-691">Parameters</span></span>

|<span data-ttu-id="68dae-692">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-692">Name</span></span>| <span data-ttu-id="68dae-693">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-693">Type</span></span>| <span data-ttu-id="68dae-694">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="68dae-694">Attributes</span></span>| <span data-ttu-id="68dae-695">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-695">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="68dae-696">String</span><span class="sxs-lookup"><span data-stu-id="68dae-696">String</span></span>||<span data-ttu-id="68dae-p142">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="68dae-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="68dae-699">String</span><span class="sxs-lookup"><span data-stu-id="68dae-699">String</span></span>||<span data-ttu-id="68dae-700">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-700">The subject of the item to be attached.</span></span> <span data-ttu-id="68dae-701">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="68dae-701">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="68dae-702">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-702">Object</span></span>| <span data-ttu-id="68dae-703">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-703">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-704">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="68dae-704">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="68dae-705">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-705">Object</span></span>| <span data-ttu-id="68dae-706">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-706">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-707">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="68dae-707">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="68dae-708">функция</span><span class="sxs-lookup"><span data-stu-id="68dae-708">function</span></span>| <span data-ttu-id="68dae-709">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-709">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-710">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="68dae-710">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="68dae-711">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="68dae-711">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="68dae-712">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="68dae-712">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="68dae-713">Ошибки</span><span class="sxs-lookup"><span data-stu-id="68dae-713">Errors</span></span>

| <span data-ttu-id="68dae-714">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="68dae-714">Error code</span></span> | <span data-ttu-id="68dae-715">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-715">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="68dae-716">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="68dae-716">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="68dae-717">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-717">Requirements</span></span>

|<span data-ttu-id="68dae-718">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-718">Requirement</span></span>| <span data-ttu-id="68dae-719">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-720">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-721">1.1</span><span class="sxs-lookup"><span data-stu-id="68dae-721">1.1</span></span>|
|[<span data-ttu-id="68dae-722">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-722">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-723">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="68dae-723">ReadWriteItem</span></span>|
|[<span data-ttu-id="68dae-724">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-724">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-725">Создание</span><span class="sxs-lookup"><span data-stu-id="68dae-725">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-726">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-726">Example</span></span>

<span data-ttu-id="68dae-727">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="68dae-727">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="68dae-728">close()</span><span class="sxs-lookup"><span data-stu-id="68dae-728">close()</span></span>

<span data-ttu-id="68dae-729">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="68dae-729">Closes the current item that is being composed.</span></span>

<span data-ttu-id="68dae-p144">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="68dae-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-732">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="68dae-732">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="68dae-733">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="68dae-733">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-734">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-734">Requirements</span></span>

|<span data-ttu-id="68dae-735">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-735">Requirement</span></span>| <span data-ttu-id="68dae-736">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-736">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-737">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-737">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-738">1.3</span><span class="sxs-lookup"><span data-stu-id="68dae-738">1.3</span></span>|
|[<span data-ttu-id="68dae-739">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-739">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-740">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="68dae-740">Restricted</span></span>|
|[<span data-ttu-id="68dae-741">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-741">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-742">Создание</span><span class="sxs-lookup"><span data-stu-id="68dae-742">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="68dae-743">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="68dae-743">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="68dae-744">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-744">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-745">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-745">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="68dae-746">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="68dae-746">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="68dae-747">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="68dae-747">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="68dae-p145">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="68dae-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-751">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-751">Parameters</span></span>

| <span data-ttu-id="68dae-752">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-752">Name</span></span> | <span data-ttu-id="68dae-753">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-753">Type</span></span> | <span data-ttu-id="68dae-754">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="68dae-754">Attributes</span></span> | <span data-ttu-id="68dae-755">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-755">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="68dae-756">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="68dae-756">String &#124; Object</span></span>| |<span data-ttu-id="68dae-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="68dae-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="68dae-759">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="68dae-759">**OR**</span></span><br/><span data-ttu-id="68dae-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="68dae-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="68dae-762">String</span><span class="sxs-lookup"><span data-stu-id="68dae-762">String</span></span> | <span data-ttu-id="68dae-763">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-763">&lt;optional&gt;</span></span> | <span data-ttu-id="68dae-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="68dae-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="68dae-766">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-766">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="68dae-767">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-767">&lt;optional&gt;</span></span> | <span data-ttu-id="68dae-768">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="68dae-768">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="68dae-769">String</span><span class="sxs-lookup"><span data-stu-id="68dae-769">String</span></span> | | <span data-ttu-id="68dae-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="68dae-772">Строка</span><span class="sxs-lookup"><span data-stu-id="68dae-772">String</span></span> | | <span data-ttu-id="68dae-773">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="68dae-773">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="68dae-774">String</span><span class="sxs-lookup"><span data-stu-id="68dae-774">String</span></span> | | <span data-ttu-id="68dae-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="68dae-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="68dae-777">Логический</span><span class="sxs-lookup"><span data-stu-id="68dae-777">Boolean</span></span> | | <span data-ttu-id="68dae-p151">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="68dae-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="68dae-780">String</span><span class="sxs-lookup"><span data-stu-id="68dae-780">String</span></span> | | <span data-ttu-id="68dae-p152">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="68dae-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="68dae-784">function</span><span class="sxs-lookup"><span data-stu-id="68dae-784">function</span></span> | <span data-ttu-id="68dae-785">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-785">&lt;optional&gt;</span></span> | <span data-ttu-id="68dae-786">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="68dae-786">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="68dae-787">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-787">Requirements</span></span>

|<span data-ttu-id="68dae-788">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-788">Requirement</span></span>| <span data-ttu-id="68dae-789">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-789">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-790">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-790">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-791">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-791">1.0</span></span>|
|[<span data-ttu-id="68dae-792">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-792">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-793">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-793">ReadItem</span></span>|
|[<span data-ttu-id="68dae-794">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-794">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-795">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-795">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="68dae-796">Примеры</span><span class="sxs-lookup"><span data-stu-id="68dae-796">Examples</span></span>

<span data-ttu-id="68dae-797">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="68dae-797">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="68dae-798">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-798">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="68dae-799">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-799">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="68dae-800">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="68dae-800">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="68dae-801">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="68dae-801">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="68dae-802">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="68dae-802">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="68dae-803">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="68dae-803">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="68dae-804">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-804">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-805">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-805">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="68dae-806">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="68dae-806">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="68dae-807">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="68dae-807">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="68dae-p153">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="68dae-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-811">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-811">Parameters</span></span>

| <span data-ttu-id="68dae-812">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-812">Name</span></span> | <span data-ttu-id="68dae-813">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-813">Type</span></span> | <span data-ttu-id="68dae-814">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="68dae-814">Attributes</span></span> | <span data-ttu-id="68dae-815">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-815">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="68dae-816">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="68dae-816">String &#124; Object</span></span>| | <span data-ttu-id="68dae-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="68dae-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="68dae-819">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="68dae-819">**OR**</span></span><br/><span data-ttu-id="68dae-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="68dae-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="68dae-822">String</span><span class="sxs-lookup"><span data-stu-id="68dae-822">String</span></span> | <span data-ttu-id="68dae-823">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-823">&lt;optional&gt;</span></span> | <span data-ttu-id="68dae-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="68dae-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="68dae-826">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-826">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="68dae-827">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-827">&lt;optional&gt;</span></span> | <span data-ttu-id="68dae-828">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="68dae-828">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="68dae-829">String</span><span class="sxs-lookup"><span data-stu-id="68dae-829">String</span></span> | | <span data-ttu-id="68dae-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="68dae-832">Строка</span><span class="sxs-lookup"><span data-stu-id="68dae-832">String</span></span> | | <span data-ttu-id="68dae-833">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="68dae-833">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="68dae-834">String</span><span class="sxs-lookup"><span data-stu-id="68dae-834">String</span></span> | | <span data-ttu-id="68dae-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="68dae-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="68dae-837">Логический</span><span class="sxs-lookup"><span data-stu-id="68dae-837">Boolean</span></span> | | <span data-ttu-id="68dae-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="68dae-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="68dae-840">String</span><span class="sxs-lookup"><span data-stu-id="68dae-840">String</span></span> | | <span data-ttu-id="68dae-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="68dae-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="68dae-844">function</span><span class="sxs-lookup"><span data-stu-id="68dae-844">function</span></span> | <span data-ttu-id="68dae-845">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-845">&lt;optional&gt;</span></span> | <span data-ttu-id="68dae-846">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="68dae-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="68dae-847">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-847">Requirements</span></span>

|<span data-ttu-id="68dae-848">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-848">Requirement</span></span>| <span data-ttu-id="68dae-849">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-850">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-851">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-851">1.0</span></span>|
|[<span data-ttu-id="68dae-852">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-853">ReadItem</span></span>|
|[<span data-ttu-id="68dae-854">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-855">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-855">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="68dae-856">Примеры</span><span class="sxs-lookup"><span data-stu-id="68dae-856">Examples</span></span>

<span data-ttu-id="68dae-857">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="68dae-857">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="68dae-858">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-858">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="68dae-859">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-859">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="68dae-860">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="68dae-860">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="68dae-861">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="68dae-861">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="68dae-862">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="68dae-862">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="68dae-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="68dae-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="68dae-864">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-864">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-865">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-866">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-866">Requirements</span></span>

|<span data-ttu-id="68dae-867">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-867">Requirement</span></span>| <span data-ttu-id="68dae-868">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-868">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-869">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-869">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-870">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-870">1.0</span></span>|
|[<span data-ttu-id="68dae-871">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-871">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-872">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-872">ReadItem</span></span>|
|[<span data-ttu-id="68dae-873">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-873">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-874">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-874">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="68dae-875">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="68dae-875">Returns:</span></span>

<span data-ttu-id="68dae-876">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-876">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="68dae-877">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-877">Example</span></span>

<span data-ttu-id="68dae-878">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-878">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="68dae-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="68dae-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="68dae-880">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-880">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-881">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-881">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-882">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-882">Parameters</span></span>

|<span data-ttu-id="68dae-883">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-883">Name</span></span>| <span data-ttu-id="68dae-884">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-884">Type</span></span>| <span data-ttu-id="68dae-885">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-885">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="68dae-886">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="68dae-886">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="68dae-887">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="68dae-887">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="68dae-888">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-888">Requirements</span></span>

|<span data-ttu-id="68dae-889">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-889">Requirement</span></span>| <span data-ttu-id="68dae-890">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-891">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-892">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-892">1.0</span></span>|
|[<span data-ttu-id="68dae-893">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-894">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="68dae-894">Restricted</span></span>|
|[<span data-ttu-id="68dae-895">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-896">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-896">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="68dae-897">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="68dae-897">Returns:</span></span>

<span data-ttu-id="68dae-898">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="68dae-898">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="68dae-899">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="68dae-899">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="68dae-900">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="68dae-900">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="68dae-901">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="68dae-901">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="68dae-902">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="68dae-902">Value of `entityType`</span></span> | <span data-ttu-id="68dae-903">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="68dae-903">Type of objects in returned array</span></span> | <span data-ttu-id="68dae-904">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-904">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="68dae-905">String</span><span class="sxs-lookup"><span data-stu-id="68dae-905">String</span></span> | <span data-ttu-id="68dae-906">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="68dae-906">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="68dae-907">Contact</span><span class="sxs-lookup"><span data-stu-id="68dae-907">Contact</span></span> | <span data-ttu-id="68dae-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="68dae-908">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="68dae-909">String</span><span class="sxs-lookup"><span data-stu-id="68dae-909">String</span></span> | <span data-ttu-id="68dae-910">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="68dae-910">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="68dae-911">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="68dae-911">MeetingSuggestion</span></span> | <span data-ttu-id="68dae-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="68dae-912">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="68dae-913">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="68dae-913">PhoneNumber</span></span> | <span data-ttu-id="68dae-914">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="68dae-914">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="68dae-915">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="68dae-915">TaskSuggestion</span></span> | <span data-ttu-id="68dae-916">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="68dae-916">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="68dae-917">String</span><span class="sxs-lookup"><span data-stu-id="68dae-917">String</span></span> | <span data-ttu-id="68dae-918">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="68dae-918">**Restricted**</span></span> |

<span data-ttu-id="68dae-919">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="68dae-919">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="68dae-920">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-920">Example</span></span>

<span data-ttu-id="68dae-921">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-921">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="68dae-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="68dae-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="68dae-923">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="68dae-923">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-924">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-924">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="68dae-925">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="68dae-925">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-926">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-926">Parameters</span></span>

|<span data-ttu-id="68dae-927">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-927">Name</span></span>| <span data-ttu-id="68dae-928">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-928">Type</span></span>| <span data-ttu-id="68dae-929">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-929">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="68dae-930">String</span><span class="sxs-lookup"><span data-stu-id="68dae-930">String</span></span>|<span data-ttu-id="68dae-931">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="68dae-931">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="68dae-932">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-932">Requirements</span></span>

|<span data-ttu-id="68dae-933">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-933">Requirement</span></span>| <span data-ttu-id="68dae-934">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-935">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-936">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-936">1.0</span></span>|
|[<span data-ttu-id="68dae-937">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-938">ReadItem</span></span>|
|[<span data-ttu-id="68dae-939">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-940">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="68dae-941">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="68dae-941">Returns:</span></span>

<span data-ttu-id="68dae-p162">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="68dae-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="68dae-944">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="68dae-944">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="68dae-945">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="68dae-945">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="68dae-946">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="68dae-946">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-947">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="68dae-p163">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="68dae-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="68dae-951">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="68dae-951">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="68dae-952">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="68dae-952">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="68dae-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="68dae-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-956">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-956">Requirements</span></span>

|<span data-ttu-id="68dae-957">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-957">Requirement</span></span>| <span data-ttu-id="68dae-958">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-958">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-959">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-959">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-960">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-960">1.0</span></span>|
|[<span data-ttu-id="68dae-961">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-961">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-962">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-962">ReadItem</span></span>|
|[<span data-ttu-id="68dae-963">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-963">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-964">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-964">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="68dae-965">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="68dae-965">Returns:</span></span>

<span data-ttu-id="68dae-p165">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="68dae-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="68dae-968">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="68dae-968">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="68dae-969">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-969">Example</span></span>

<span data-ttu-id="68dae-970">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="68dae-970">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="68dae-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="68dae-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="68dae-972">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="68dae-972">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-973">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-973">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="68dae-974">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="68dae-974">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="68dae-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="68dae-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-977">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-977">Parameters</span></span>

|<span data-ttu-id="68dae-978">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-978">Name</span></span>| <span data-ttu-id="68dae-979">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-979">Type</span></span>| <span data-ttu-id="68dae-980">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-980">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="68dae-981">String</span><span class="sxs-lookup"><span data-stu-id="68dae-981">String</span></span>|<span data-ttu-id="68dae-982">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="68dae-982">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="68dae-983">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-983">Requirements</span></span>

|<span data-ttu-id="68dae-984">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-984">Requirement</span></span>| <span data-ttu-id="68dae-985">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-986">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-987">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-987">1.0</span></span>|
|[<span data-ttu-id="68dae-988">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-988">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-989">ReadItem</span></span>|
|[<span data-ttu-id="68dae-990">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-990">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-991">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-991">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="68dae-992">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="68dae-992">Returns:</span></span>

<span data-ttu-id="68dae-993">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="68dae-993">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="68dae-994">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="68dae-994">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="68dae-995">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-995">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="68dae-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="68dae-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="68dae-997">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-997">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="68dae-p167">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="68dae-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-1000">В Outlook в Интернете метод возвращает строку "null", если текст не выбран, но курсор находится в теле.</span><span class="sxs-lookup"><span data-stu-id="68dae-1000">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="68dae-1001">Чтобы проверить эту ситуацию, добавьте код, подобный приведенному ниже:</span><span class="sxs-lookup"><span data-stu-id="68dae-1001">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="68dae-1002">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-1002">Parameters</span></span>

|<span data-ttu-id="68dae-1003">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-1003">Name</span></span>| <span data-ttu-id="68dae-1004">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-1004">Type</span></span>| <span data-ttu-id="68dae-1005">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="68dae-1005">Attributes</span></span>| <span data-ttu-id="68dae-1006">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-1006">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="68dae-1007">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="68dae-1007">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="68dae-p169">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="68dae-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="68dae-1011">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-1011">Object</span></span>| <span data-ttu-id="68dae-1012">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1012">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1013">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="68dae-1013">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="68dae-1014">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-1014">Object</span></span>| <span data-ttu-id="68dae-1015">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1015">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1016">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="68dae-1016">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="68dae-1017">функция</span><span class="sxs-lookup"><span data-stu-id="68dae-1017">function</span></span>||<span data-ttu-id="68dae-1018">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="68dae-1018">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="68dae-1019">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="68dae-1019">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="68dae-1020">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="68dae-1020">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="68dae-1021">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-1021">Requirements</span></span>

|<span data-ttu-id="68dae-1022">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-1022">Requirement</span></span>| <span data-ttu-id="68dae-1023">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-1024">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-1025">1.2</span><span class="sxs-lookup"><span data-stu-id="68dae-1025">1.2</span></span>|
|[<span data-ttu-id="68dae-1026">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-1026">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-1027">ReadItem</span></span>|
|[<span data-ttu-id="68dae-1028">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-1028">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-1029">Создание</span><span class="sxs-lookup"><span data-stu-id="68dae-1029">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="68dae-1030">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="68dae-1030">Returns:</span></span>

<span data-ttu-id="68dae-1031">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="68dae-1031">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="68dae-1032">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="68dae-1032">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="68dae-1033">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-1033">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="68dae-1034">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="68dae-1034">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="68dae-1035">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="68dae-1035">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="68dae-1036">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="68dae-1036">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-1037">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-1037">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-1038">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-1038">Requirements</span></span>

|<span data-ttu-id="68dae-1039">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-1039">Requirement</span></span>| <span data-ttu-id="68dae-1040">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-1041">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="68dae-1042">1.6</span></span> |
|[<span data-ttu-id="68dae-1043">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-1044">ReadItem</span></span>|
|[<span data-ttu-id="68dae-1045">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-1046">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="68dae-1047">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="68dae-1047">Returns:</span></span>

<span data-ttu-id="68dae-1048">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="68dae-1048">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="68dae-1049">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-1049">Example</span></span>

<span data-ttu-id="68dae-1050">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="68dae-1050">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="68dae-1051">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="68dae-1051">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="68dae-p172">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="68dae-p172">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-1054">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="68dae-1054">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="68dae-p173">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="68dae-p173">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="68dae-1058">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="68dae-1058">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="68dae-1059">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="68dae-1059">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="68dae-p174">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="68dae-p174">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="68dae-1063">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-1063">Requirements</span></span>

|<span data-ttu-id="68dae-1064">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-1064">Requirement</span></span>| <span data-ttu-id="68dae-1065">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-1066">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-1067">1.6</span><span class="sxs-lookup"><span data-stu-id="68dae-1067">1.6</span></span> |
|[<span data-ttu-id="68dae-1068">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-1068">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-1069">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-1069">ReadItem</span></span>|
|[<span data-ttu-id="68dae-1070">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-1070">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-1071">Чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-1071">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="68dae-1072">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="68dae-1072">Returns:</span></span>

<span data-ttu-id="68dae-p175">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="68dae-p175">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="68dae-1075">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-1075">Example</span></span>

<span data-ttu-id="68dae-1076">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="68dae-1076">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="68dae-1077">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="68dae-1077">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="68dae-1078">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-1078">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="68dae-p176">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="68dae-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-1082">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-1082">Parameters</span></span>

|<span data-ttu-id="68dae-1083">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-1083">Name</span></span>| <span data-ttu-id="68dae-1084">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-1084">Type</span></span>| <span data-ttu-id="68dae-1085">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="68dae-1085">Attributes</span></span>| <span data-ttu-id="68dae-1086">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-1086">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="68dae-1087">function</span><span class="sxs-lookup"><span data-stu-id="68dae-1087">function</span></span>||<span data-ttu-id="68dae-1088">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="68dae-1088">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="68dae-1089">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="68dae-1089">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="68dae-1090">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="68dae-1090">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="68dae-1091">Объект</span><span class="sxs-lookup"><span data-stu-id="68dae-1091">Object</span></span>| <span data-ttu-id="68dae-1092">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1093">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="68dae-1093">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="68dae-1094">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="68dae-1094">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="68dae-1095">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-1095">Requirements</span></span>

|<span data-ttu-id="68dae-1096">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-1096">Requirement</span></span>| <span data-ttu-id="68dae-1097">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-1098">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-1099">1.0</span><span class="sxs-lookup"><span data-stu-id="68dae-1099">1.0</span></span>|
|[<span data-ttu-id="68dae-1100">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="68dae-1101">ReadItem</span></span>|
|[<span data-ttu-id="68dae-1102">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-1103">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="68dae-1103">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-1104">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-1104">Example</span></span>

<span data-ttu-id="68dae-p179">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="68dae-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="68dae-1108">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="68dae-1108">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="68dae-1109">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="68dae-1109">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="68dae-1110">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="68dae-1110">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="68dae-1111">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="68dae-1111">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="68dae-1112">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="68dae-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="68dae-1113">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="68dae-1113">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-1114">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-1114">Parameters</span></span>

|<span data-ttu-id="68dae-1115">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-1115">Name</span></span>| <span data-ttu-id="68dae-1116">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-1116">Type</span></span>| <span data-ttu-id="68dae-1117">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="68dae-1117">Attributes</span></span>| <span data-ttu-id="68dae-1118">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="68dae-1119">String</span><span class="sxs-lookup"><span data-stu-id="68dae-1119">String</span></span>||<span data-ttu-id="68dae-1120">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="68dae-1120">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="68dae-1121">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-1121">Object</span></span>| <span data-ttu-id="68dae-1122">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1123">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="68dae-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="68dae-1124">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-1124">Object</span></span>| <span data-ttu-id="68dae-1125">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1126">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="68dae-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="68dae-1127">функция</span><span class="sxs-lookup"><span data-stu-id="68dae-1127">function</span></span>| <span data-ttu-id="68dae-1128">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1129">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="68dae-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="68dae-1130">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="68dae-1130">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="68dae-1131">Ошибки</span><span class="sxs-lookup"><span data-stu-id="68dae-1131">Errors</span></span>

| <span data-ttu-id="68dae-1132">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="68dae-1132">Error code</span></span> | <span data-ttu-id="68dae-1133">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-1133">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="68dae-1134">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="68dae-1134">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="68dae-1135">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-1135">Requirements</span></span>

|<span data-ttu-id="68dae-1136">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-1136">Requirement</span></span>| <span data-ttu-id="68dae-1137">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-1138">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="68dae-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-1139">1.1</span><span class="sxs-lookup"><span data-stu-id="68dae-1139">1.1</span></span>|
|[<span data-ttu-id="68dae-1140">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-1141">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="68dae-1141">ReadWriteItem</span></span>|
|[<span data-ttu-id="68dae-1142">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-1143">Создание</span><span class="sxs-lookup"><span data-stu-id="68dae-1143">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-1144">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-1144">Example</span></span>

<span data-ttu-id="68dae-1145">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="68dae-1145">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="68dae-1146">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="68dae-1146">saveAsync([options], callback)</span></span>

<span data-ttu-id="68dae-1147">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="68dae-1147">Asynchronously saves an item.</span></span>

<span data-ttu-id="68dae-1148">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="68dae-1148">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="68dae-1149">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="68dae-1149">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="68dae-1150">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="68dae-1150">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-1151">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="68dae-1151">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="68dae-1152">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="68dae-1152">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="68dae-p183">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="68dae-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="68dae-1156">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="68dae-1156">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="68dae-1157">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="68dae-1157">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="68dae-1158">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="68dae-1158">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="68dae-1159">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="68dae-1159">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="68dae-1160">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="68dae-1160">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-1161">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-1161">Parameters</span></span>

|<span data-ttu-id="68dae-1162">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-1162">Name</span></span>| <span data-ttu-id="68dae-1163">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-1163">Type</span></span>| <span data-ttu-id="68dae-1164">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="68dae-1164">Attributes</span></span>| <span data-ttu-id="68dae-1165">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-1165">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="68dae-1166">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-1166">Object</span></span>| <span data-ttu-id="68dae-1167">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1167">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1168">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="68dae-1168">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="68dae-1169">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-1169">Object</span></span>| <span data-ttu-id="68dae-1170">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1170">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1171">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="68dae-1171">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="68dae-1172">функция</span><span class="sxs-lookup"><span data-stu-id="68dae-1172">function</span></span>||<span data-ttu-id="68dae-1173">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="68dae-1173">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="68dae-1174">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="68dae-1174">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="68dae-1175">Requirements</span><span class="sxs-lookup"><span data-stu-id="68dae-1175">Requirements</span></span>

|<span data-ttu-id="68dae-1176">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-1176">Requirement</span></span>| <span data-ttu-id="68dae-1177">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-1178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-1179">1.3</span><span class="sxs-lookup"><span data-stu-id="68dae-1179">1.3</span></span>|
|[<span data-ttu-id="68dae-1180">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-1181">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="68dae-1181">ReadWriteItem</span></span>|
|[<span data-ttu-id="68dae-1182">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-1183">Создание</span><span class="sxs-lookup"><span data-stu-id="68dae-1183">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="68dae-1184">Примеры</span><span class="sxs-lookup"><span data-stu-id="68dae-1184">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="68dae-p185">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="68dae-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="68dae-1187">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="68dae-1187">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="68dae-1188">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="68dae-1188">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="68dae-p186">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="68dae-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="68dae-1192">Параметры</span><span class="sxs-lookup"><span data-stu-id="68dae-1192">Parameters</span></span>

|<span data-ttu-id="68dae-1193">Имя</span><span class="sxs-lookup"><span data-stu-id="68dae-1193">Name</span></span>| <span data-ttu-id="68dae-1194">Тип</span><span class="sxs-lookup"><span data-stu-id="68dae-1194">Type</span></span>| <span data-ttu-id="68dae-1195">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="68dae-1195">Attributes</span></span>| <span data-ttu-id="68dae-1196">Описание</span><span class="sxs-lookup"><span data-stu-id="68dae-1196">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="68dae-1197">String</span><span class="sxs-lookup"><span data-stu-id="68dae-1197">String</span></span>||<span data-ttu-id="68dae-p187">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="68dae-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="68dae-1201">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-1201">Object</span></span>| <span data-ttu-id="68dae-1202">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1202">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1203">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="68dae-1203">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="68dae-1204">Object</span><span class="sxs-lookup"><span data-stu-id="68dae-1204">Object</span></span>| <span data-ttu-id="68dae-1205">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1205">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1206">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="68dae-1206">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="68dae-1207">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="68dae-1207">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="68dae-1208">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="68dae-1208">&lt;optional&gt;</span></span>|<span data-ttu-id="68dae-1209">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="68dae-1209">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="68dae-1210">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="68dae-1210">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="68dae-1211">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="68dae-1211">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="68dae-1212">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="68dae-1212">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="68dae-1213">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="68dae-1213">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="68dae-1214">функция</span><span class="sxs-lookup"><span data-stu-id="68dae-1214">function</span></span>||<span data-ttu-id="68dae-1215">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="68dae-1215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="68dae-1216">Требования</span><span class="sxs-lookup"><span data-stu-id="68dae-1216">Requirements</span></span>

|<span data-ttu-id="68dae-1217">Требование</span><span class="sxs-lookup"><span data-stu-id="68dae-1217">Requirement</span></span>| <span data-ttu-id="68dae-1218">Значение</span><span class="sxs-lookup"><span data-stu-id="68dae-1218">Value</span></span>|
|---|---|
|[<span data-ttu-id="68dae-1219">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="68dae-1219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68dae-1220">1.2</span><span class="sxs-lookup"><span data-stu-id="68dae-1220">1.2</span></span>|
|[<span data-ttu-id="68dae-1221">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="68dae-1221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="68dae-1222">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="68dae-1222">ReadWriteItem</span></span>|
|[<span data-ttu-id="68dae-1223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="68dae-1223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="68dae-1224">Создание</span><span class="sxs-lookup"><span data-stu-id="68dae-1224">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="68dae-1225">Пример</span><span class="sxs-lookup"><span data-stu-id="68dae-1225">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
