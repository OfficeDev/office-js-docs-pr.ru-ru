---
title: Office. Context. Mailbox. Item — набор требований 1,6
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: d6b77724290d9d100ff098baf11d97ba600bd8ee
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696038"
---
# <a name="item"></a><span data-ttu-id="35eab-102">item</span><span class="sxs-lookup"><span data-stu-id="35eab-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="35eab-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="35eab-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="35eab-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="35eab-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="35eab-106">Requirements</span></span>

|<span data-ttu-id="35eab-107">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-107">Requirement</span></span>| <span data-ttu-id="35eab-108">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-110">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-110">1.0</span></span>|
|[<span data-ttu-id="35eab-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="35eab-112">Restricted</span></span>|
|[<span data-ttu-id="35eab-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="35eab-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="35eab-115">Members and methods</span></span>

| <span data-ttu-id="35eab-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-116">Member</span></span> | <span data-ttu-id="35eab-117">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="35eab-118">attachments</span><span class="sxs-lookup"><span data-stu-id="35eab-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="35eab-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-119">Member</span></span> |
| [<span data-ttu-id="35eab-120">bcc</span><span class="sxs-lookup"><span data-stu-id="35eab-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="35eab-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-121">Member</span></span> |
| [<span data-ttu-id="35eab-122">body</span><span class="sxs-lookup"><span data-stu-id="35eab-122">body</span></span>](#body-body) | <span data-ttu-id="35eab-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-123">Member</span></span> |
| [<span data-ttu-id="35eab-124">cc</span><span class="sxs-lookup"><span data-stu-id="35eab-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="35eab-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-125">Member</span></span> |
| [<span data-ttu-id="35eab-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="35eab-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="35eab-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-127">Member</span></span> |
| [<span data-ttu-id="35eab-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="35eab-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="35eab-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-129">Member</span></span> |
| [<span data-ttu-id="35eab-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="35eab-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="35eab-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-131">Member</span></span> |
| [<span data-ttu-id="35eab-132">end</span><span class="sxs-lookup"><span data-stu-id="35eab-132">end</span></span>](#end-datetime) | <span data-ttu-id="35eab-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-133">Member</span></span> |
| [<span data-ttu-id="35eab-134">from</span><span class="sxs-lookup"><span data-stu-id="35eab-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="35eab-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-135">Member</span></span> |
| [<span data-ttu-id="35eab-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="35eab-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="35eab-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-137">Member</span></span> |
| [<span data-ttu-id="35eab-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="35eab-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="35eab-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-139">Member</span></span> |
| [<span data-ttu-id="35eab-140">itemId</span><span class="sxs-lookup"><span data-stu-id="35eab-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="35eab-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-141">Member</span></span> |
| [<span data-ttu-id="35eab-142">itemType</span><span class="sxs-lookup"><span data-stu-id="35eab-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="35eab-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-143">Member</span></span> |
| [<span data-ttu-id="35eab-144">location</span><span class="sxs-lookup"><span data-stu-id="35eab-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="35eab-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-145">Member</span></span> |
| [<span data-ttu-id="35eab-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="35eab-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="35eab-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-147">Member</span></span> |
| [<span data-ttu-id="35eab-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="35eab-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="35eab-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-149">Member</span></span> |
| [<span data-ttu-id="35eab-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="35eab-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="35eab-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-151">Member</span></span> |
| [<span data-ttu-id="35eab-152">organizer</span><span class="sxs-lookup"><span data-stu-id="35eab-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="35eab-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-153">Member</span></span> |
| [<span data-ttu-id="35eab-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="35eab-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="35eab-155">Member</span><span class="sxs-lookup"><span data-stu-id="35eab-155">Member</span></span> |
| [<span data-ttu-id="35eab-156">sender</span><span class="sxs-lookup"><span data-stu-id="35eab-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="35eab-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-157">Member</span></span> |
| [<span data-ttu-id="35eab-158">start</span><span class="sxs-lookup"><span data-stu-id="35eab-158">start</span></span>](#start-datetime) | <span data-ttu-id="35eab-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-159">Member</span></span> |
| [<span data-ttu-id="35eab-160">subject</span><span class="sxs-lookup"><span data-stu-id="35eab-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="35eab-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-161">Member</span></span> |
| [<span data-ttu-id="35eab-162">to</span><span class="sxs-lookup"><span data-stu-id="35eab-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="35eab-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="35eab-163">Member</span></span> |
| [<span data-ttu-id="35eab-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="35eab-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="35eab-165">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-165">Method</span></span> |
| [<span data-ttu-id="35eab-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="35eab-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="35eab-167">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-167">Method</span></span> |
| [<span data-ttu-id="35eab-168">close</span><span class="sxs-lookup"><span data-stu-id="35eab-168">close</span></span>](#close) | <span data-ttu-id="35eab-169">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-169">Method</span></span> |
| [<span data-ttu-id="35eab-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="35eab-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="35eab-171">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-171">Method</span></span> |
| [<span data-ttu-id="35eab-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="35eab-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="35eab-173">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-173">Method</span></span> |
| [<span data-ttu-id="35eab-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="35eab-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="35eab-175">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-175">Method</span></span> |
| [<span data-ttu-id="35eab-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="35eab-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="35eab-177">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-177">Method</span></span> |
| [<span data-ttu-id="35eab-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="35eab-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="35eab-179">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-179">Method</span></span> |
| [<span data-ttu-id="35eab-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="35eab-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="35eab-181">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-181">Method</span></span> |
| [<span data-ttu-id="35eab-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="35eab-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="35eab-183">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-183">Method</span></span> |
| [<span data-ttu-id="35eab-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="35eab-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="35eab-185">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-185">Method</span></span> |
| [<span data-ttu-id="35eab-186">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="35eab-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="35eab-187">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-187">Method</span></span> |
| [<span data-ttu-id="35eab-188">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="35eab-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="35eab-189">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-189">Method</span></span> |
| [<span data-ttu-id="35eab-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="35eab-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="35eab-191">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-191">Method</span></span> |
| [<span data-ttu-id="35eab-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="35eab-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="35eab-193">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-193">Method</span></span> |
| [<span data-ttu-id="35eab-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="35eab-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="35eab-195">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-195">Method</span></span> |
| [<span data-ttu-id="35eab-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="35eab-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="35eab-197">Метод</span><span class="sxs-lookup"><span data-stu-id="35eab-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="35eab-198">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-198">Example</span></span>

<span data-ttu-id="35eab-199">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="35eab-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="35eab-200">Элементы</span><span class="sxs-lookup"><span data-stu-id="35eab-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="35eab-201">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="35eab-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="35eab-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-204">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="35eab-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="35eab-205">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="35eab-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-206">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-206">Type</span></span>

*   <span data-ttu-id="35eab-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="35eab-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-208">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-208">Requirements</span></span>

|<span data-ttu-id="35eab-209">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-209">Requirement</span></span>| <span data-ttu-id="35eab-210">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-211">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-212">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-212">1.0</span></span>|
|[<span data-ttu-id="35eab-213">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-214">ReadItem</span></span>|
|[<span data-ttu-id="35eab-215">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-216">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-217">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-217">Example</span></span>

<span data-ttu-id="35eab-218">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="35eab-219">СК: [получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-220">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="35eab-221">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="35eab-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-222">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-222">Type</span></span>

*   [<span data-ttu-id="35eab-223">Получатели</span><span class="sxs-lookup"><span data-stu-id="35eab-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="35eab-224">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-224">Requirements</span></span>

|<span data-ttu-id="35eab-225">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-225">Requirement</span></span>| <span data-ttu-id="35eab-226">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-227">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-228">1.1</span><span class="sxs-lookup"><span data-stu-id="35eab-228">1.1</span></span>|
|[<span data-ttu-id="35eab-229">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-230">ReadItem</span></span>|
|[<span data-ttu-id="35eab-231">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-232">Создание</span><span class="sxs-lookup"><span data-stu-id="35eab-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-233">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-233">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="35eab-234">основной текст: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-235">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-236">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-236">Type</span></span>

*   [<span data-ttu-id="35eab-237">Body</span><span class="sxs-lookup"><span data-stu-id="35eab-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="35eab-238">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-238">Requirements</span></span>

|<span data-ttu-id="35eab-239">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-239">Requirement</span></span>| <span data-ttu-id="35eab-240">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-241">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-242">1.1</span><span class="sxs-lookup"><span data-stu-id="35eab-242">1.1</span></span>|
|[<span data-ttu-id="35eab-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-244">ReadItem</span></span>|
|[<span data-ttu-id="35eab-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-247">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-247">Example</span></span>

<span data-ttu-id="35eab-248">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="35eab-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="35eab-249">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="35eab-249">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="35eab-250">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-251">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="35eab-252">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35eab-253">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="35eab-253">Read mode</span></span>

<span data-ttu-id="35eab-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="35eab-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="35eab-256">Режим создания</span><span class="sxs-lookup"><span data-stu-id="35eab-256">Compose mode</span></span>

<span data-ttu-id="35eab-257">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="35eab-258">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-258">Type</span></span>

*   <span data-ttu-id="35eab-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-260">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-260">Requirements</span></span>

|<span data-ttu-id="35eab-261">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-261">Requirement</span></span>| <span data-ttu-id="35eab-262">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-263">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-264">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-264">1.0</span></span>|
|[<span data-ttu-id="35eab-265">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-266">ReadItem</span></span>|
|[<span data-ttu-id="35eab-267">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-268">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="35eab-269">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="35eab-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="35eab-270">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="35eab-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="35eab-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="35eab-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="35eab-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="35eab-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-275">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-275">Type</span></span>

*   <span data-ttu-id="35eab-276">String</span><span class="sxs-lookup"><span data-stu-id="35eab-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-277">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-277">Requirements</span></span>

|<span data-ttu-id="35eab-278">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-278">Requirement</span></span>| <span data-ttu-id="35eab-279">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-280">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-281">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-281">1.0</span></span>|
|[<span data-ttu-id="35eab-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-283">ReadItem</span></span>|
|[<span data-ttu-id="35eab-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-285">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-286">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="35eab-287">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="35eab-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="35eab-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-290">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-290">Type</span></span>

*   <span data-ttu-id="35eab-291">Дата</span><span class="sxs-lookup"><span data-stu-id="35eab-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-292">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-292">Requirements</span></span>

|<span data-ttu-id="35eab-293">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-293">Requirement</span></span>| <span data-ttu-id="35eab-294">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-295">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-296">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-296">1.0</span></span>|
|[<span data-ttu-id="35eab-297">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-298">ReadItem</span></span>|
|[<span data-ttu-id="35eab-299">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-300">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-301">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="35eab-302">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="35eab-302">dateTimeModified: Date</span></span>

<span data-ttu-id="35eab-303">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-303">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="35eab-304">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-304">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-305">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-306">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-306">Type</span></span>

*   <span data-ttu-id="35eab-307">Дата</span><span class="sxs-lookup"><span data-stu-id="35eab-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-308">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-308">Requirements</span></span>

|<span data-ttu-id="35eab-309">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-309">Requirement</span></span>| <span data-ttu-id="35eab-310">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-312">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-312">1.0</span></span>|
|[<span data-ttu-id="35eab-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-314">ReadItem</span></span>|
|[<span data-ttu-id="35eab-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-316">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-317">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="35eab-318">конец: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.6) )</span><span class="sxs-lookup"><span data-stu-id="35eab-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-319">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="35eab-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="35eab-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35eab-322">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="35eab-322">Read mode</span></span>

<span data-ttu-id="35eab-323">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="35eab-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="35eab-324">Режим создания</span><span class="sxs-lookup"><span data-stu-id="35eab-324">Compose mode</span></span>

<span data-ttu-id="35eab-325">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="35eab-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="35eab-326">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="35eab-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="35eab-327">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="35eab-328">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-328">Type</span></span>

*   <span data-ttu-id="35eab-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-330">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-330">Requirements</span></span>

|<span data-ttu-id="35eab-331">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-331">Requirement</span></span>| <span data-ttu-id="35eab-332">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-333">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-334">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-334">1.0</span></span>|
|[<span data-ttu-id="35eab-335">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-336">ReadItem</span></span>|
|[<span data-ttu-id="35eab-337">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-338">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="35eab-339">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="35eab-p113">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="35eab-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-344">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="35eab-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-345">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-345">Type</span></span>

*   [<span data-ttu-id="35eab-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="35eab-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="35eab-347">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-347">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="35eab-348">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-348">Requirements</span></span>

|<span data-ttu-id="35eab-349">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-349">Requirement</span></span>| <span data-ttu-id="35eab-350">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-351">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-352">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-352">1.0</span></span>|
|[<span data-ttu-id="35eab-353">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-354">ReadItem</span></span>|
|[<span data-ttu-id="35eab-355">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-356">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-356">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="35eab-357">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="35eab-357">internetMessageId: String</span></span>

<span data-ttu-id="35eab-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-360">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-360">Type</span></span>

*   <span data-ttu-id="35eab-361">String</span><span class="sxs-lookup"><span data-stu-id="35eab-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-362">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-362">Requirements</span></span>

|<span data-ttu-id="35eab-363">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-363">Requirement</span></span>| <span data-ttu-id="35eab-364">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-365">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-366">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-366">1.0</span></span>|
|[<span data-ttu-id="35eab-367">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-368">ReadItem</span></span>|
|[<span data-ttu-id="35eab-369">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-370">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-371">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="35eab-372">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="35eab-372">itemClass: String</span></span>

<span data-ttu-id="35eab-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="35eab-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="35eab-377">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-377">Type</span></span> | <span data-ttu-id="35eab-378">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-378">Description</span></span> | <span data-ttu-id="35eab-379">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="35eab-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="35eab-380">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="35eab-380">Appointment items</span></span> | <span data-ttu-id="35eab-381">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="35eab-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="35eab-382">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="35eab-382">Message items</span></span> | <span data-ttu-id="35eab-383">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="35eab-384">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="35eab-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-385">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-385">Type</span></span>

*   <span data-ttu-id="35eab-386">String</span><span class="sxs-lookup"><span data-stu-id="35eab-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-387">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-387">Requirements</span></span>

|<span data-ttu-id="35eab-388">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-388">Requirement</span></span>| <span data-ttu-id="35eab-389">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-390">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-391">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-391">1.0</span></span>|
|[<span data-ttu-id="35eab-392">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-393">ReadItem</span></span>|
|[<span data-ttu-id="35eab-394">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-395">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-396">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="35eab-397">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="35eab-397">(nullable) itemId: String</span></span>

<span data-ttu-id="35eab-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-400">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="35eab-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="35eab-401">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="35eab-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="35eab-402">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="35eab-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="35eab-403">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="35eab-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="35eab-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-406">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-406">Type</span></span>

*   <span data-ttu-id="35eab-407">String</span><span class="sxs-lookup"><span data-stu-id="35eab-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-408">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-408">Requirements</span></span>

|<span data-ttu-id="35eab-409">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-409">Requirement</span></span>| <span data-ttu-id="35eab-410">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-412">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-412">1.0</span></span>|
|[<span data-ttu-id="35eab-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-414">ReadItem</span></span>|
|[<span data-ttu-id="35eab-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-417">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-417">Example</span></span>

<span data-ttu-id="35eab-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="35eab-420">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-420">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-421">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="35eab-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="35eab-422">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="35eab-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-423">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-423">Type</span></span>

*   [<span data-ttu-id="35eab-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="35eab-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="35eab-425">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-425">Requirements</span></span>

|<span data-ttu-id="35eab-426">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-426">Requirement</span></span>| <span data-ttu-id="35eab-427">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-428">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-429">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-429">1.0</span></span>|
|[<span data-ttu-id="35eab-430">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-431">ReadItem</span></span>|
|[<span data-ttu-id="35eab-432">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-433">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-434">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-434">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="35eab-435">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location?view=outlook-js-1.6) )</span><span class="sxs-lookup"><span data-stu-id="35eab-435">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-436">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35eab-437">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="35eab-437">Read mode</span></span>

<span data-ttu-id="35eab-438">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="35eab-439">Режим создания</span><span class="sxs-lookup"><span data-stu-id="35eab-439">Compose mode</span></span>

<span data-ttu-id="35eab-440">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="35eab-441">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-441">Type</span></span>

*   <span data-ttu-id="35eab-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-443">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-443">Requirements</span></span>

|<span data-ttu-id="35eab-444">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-444">Requirement</span></span>| <span data-ttu-id="35eab-445">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-446">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-447">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-447">1.0</span></span>|
|[<span data-ttu-id="35eab-448">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-449">ReadItem</span></span>|
|[<span data-ttu-id="35eab-450">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-451">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-451">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="35eab-452">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="35eab-452">normalizedSubject: String</span></span>

<span data-ttu-id="35eab-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="35eab-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="35eab-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-457">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-457">Type</span></span>

*   <span data-ttu-id="35eab-458">String</span><span class="sxs-lookup"><span data-stu-id="35eab-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-459">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-459">Requirements</span></span>

|<span data-ttu-id="35eab-460">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-460">Requirement</span></span>| <span data-ttu-id="35eab-461">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-462">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-463">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-463">1.0</span></span>|
|[<span data-ttu-id="35eab-464">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-465">ReadItem</span></span>|
|[<span data-ttu-id="35eab-466">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-467">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-468">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-468">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="35eab-469">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-469">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-470">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-471">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-471">Type</span></span>

*   [<span data-ttu-id="35eab-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="35eab-472">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="35eab-473">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-473">Requirements</span></span>

|<span data-ttu-id="35eab-474">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-474">Requirement</span></span>| <span data-ttu-id="35eab-475">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-476">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-477">1.3</span><span class="sxs-lookup"><span data-stu-id="35eab-477">1.3</span></span>|
|[<span data-ttu-id="35eab-478">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-479">ReadItem</span></span>|
|[<span data-ttu-id="35eab-480">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-481">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-482">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-482">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="35eab-483">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-483">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-484">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="35eab-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="35eab-485">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35eab-486">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="35eab-486">Read mode</span></span>

<span data-ttu-id="35eab-487">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="35eab-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="35eab-488">Режим создания</span><span class="sxs-lookup"><span data-stu-id="35eab-488">Compose mode</span></span>

<span data-ttu-id="35eab-489">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="35eab-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="35eab-490">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-490">Type</span></span>

*   <span data-ttu-id="35eab-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-492">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-492">Requirements</span></span>

|<span data-ttu-id="35eab-493">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-493">Requirement</span></span>| <span data-ttu-id="35eab-494">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-495">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-496">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-496">1.0</span></span>|
|[<span data-ttu-id="35eab-497">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-498">ReadItem</span></span>|
|[<span data-ttu-id="35eab-499">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-500">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-500">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="35eab-501">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-501">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-504">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-504">Type</span></span>

*   [<span data-ttu-id="35eab-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="35eab-505">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="35eab-506">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-506">Requirements</span></span>

|<span data-ttu-id="35eab-507">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-507">Requirement</span></span>| <span data-ttu-id="35eab-508">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-509">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-510">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-510">1.0</span></span>|
|[<span data-ttu-id="35eab-511">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-512">ReadItem</span></span>|
|[<span data-ttu-id="35eab-513">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-514">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-515">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-515">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="35eab-516">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-516">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-517">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="35eab-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="35eab-518">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35eab-519">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="35eab-519">Read mode</span></span>

<span data-ttu-id="35eab-520">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="35eab-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="35eab-521">Режим создания</span><span class="sxs-lookup"><span data-stu-id="35eab-521">Compose mode</span></span>

<span data-ttu-id="35eab-522">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="35eab-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="35eab-523">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-523">Type</span></span>

*   <span data-ttu-id="35eab-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-525">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-525">Requirements</span></span>

|<span data-ttu-id="35eab-526">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-526">Requirement</span></span>| <span data-ttu-id="35eab-527">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-528">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-529">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-529">1.0</span></span>|
|[<span data-ttu-id="35eab-530">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-531">ReadItem</span></span>|
|[<span data-ttu-id="35eab-532">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-533">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-533">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="35eab-534">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-534">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="35eab-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="35eab-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="35eab-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-539">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="35eab-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="35eab-540">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-540">Type</span></span>

*   [<span data-ttu-id="35eab-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="35eab-541">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="35eab-542">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-542">Requirements</span></span>

|<span data-ttu-id="35eab-543">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-543">Requirement</span></span>| <span data-ttu-id="35eab-544">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-545">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-546">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-546">1.0</span></span>|
|[<span data-ttu-id="35eab-547">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-548">ReadItem</span></span>|
|[<span data-ttu-id="35eab-549">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-550">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-551">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-551">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="35eab-552">Начало: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.6) )</span><span class="sxs-lookup"><span data-stu-id="35eab-552">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-553">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="35eab-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="35eab-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35eab-556">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="35eab-556">Read mode</span></span>

<span data-ttu-id="35eab-557">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="35eab-557">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="35eab-558">Режим создания</span><span class="sxs-lookup"><span data-stu-id="35eab-558">Compose mode</span></span>

<span data-ttu-id="35eab-559">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="35eab-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="35eab-560">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="35eab-560">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="35eab-561">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="35eab-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="35eab-562">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-562">Type</span></span>

*   <span data-ttu-id="35eab-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-564">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-564">Requirements</span></span>

|<span data-ttu-id="35eab-565">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-565">Requirement</span></span>| <span data-ttu-id="35eab-566">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-567">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-568">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-568">1.0</span></span>|
|[<span data-ttu-id="35eab-569">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-570">ReadItem</span></span>|
|[<span data-ttu-id="35eab-571">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-572">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="35eab-573">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject?view=outlook-js-1.6) )</span><span class="sxs-lookup"><span data-stu-id="35eab-573">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-574">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="35eab-575">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="35eab-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35eab-576">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="35eab-576">Read mode</span></span>

<span data-ttu-id="35eab-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="35eab-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="35eab-579">Режим создания</span><span class="sxs-lookup"><span data-stu-id="35eab-579">Compose mode</span></span>

<span data-ttu-id="35eab-580">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="35eab-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="35eab-581">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-581">Type</span></span>

*   <span data-ttu-id="35eab-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-583">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-583">Requirements</span></span>

|<span data-ttu-id="35eab-584">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-584">Requirement</span></span>| <span data-ttu-id="35eab-585">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-586">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-587">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-587">1.0</span></span>|
|[<span data-ttu-id="35eab-588">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-589">ReadItem</span></span>|
|[<span data-ttu-id="35eab-590">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-591">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-591">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="35eab-592">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-592">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="35eab-593">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="35eab-594">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="35eab-595">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="35eab-595">Read mode</span></span>

<span data-ttu-id="35eab-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="35eab-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="35eab-598">Режим создания</span><span class="sxs-lookup"><span data-stu-id="35eab-598">Compose mode</span></span>

<span data-ttu-id="35eab-599">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="35eab-600">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-600">Type</span></span>

*   <span data-ttu-id="35eab-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-602">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-602">Requirements</span></span>

|<span data-ttu-id="35eab-603">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-603">Requirement</span></span>| <span data-ttu-id="35eab-604">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-605">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-606">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-606">1.0</span></span>|
|[<span data-ttu-id="35eab-607">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-608">ReadItem</span></span>|
|[<span data-ttu-id="35eab-609">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-610">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="35eab-611">Методы</span><span class="sxs-lookup"><span data-stu-id="35eab-611">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="35eab-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35eab-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="35eab-613">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="35eab-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="35eab-614">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="35eab-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="35eab-615">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="35eab-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-616">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-616">Parameters</span></span>

|<span data-ttu-id="35eab-617">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-617">Name</span></span>| <span data-ttu-id="35eab-618">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-618">Type</span></span>| <span data-ttu-id="35eab-619">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="35eab-619">Attributes</span></span>| <span data-ttu-id="35eab-620">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="35eab-621">String</span><span class="sxs-lookup"><span data-stu-id="35eab-621">String</span></span>||<span data-ttu-id="35eab-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="35eab-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="35eab-624">String</span><span class="sxs-lookup"><span data-stu-id="35eab-624">String</span></span>||<span data-ttu-id="35eab-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="35eab-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="35eab-627">Объект</span><span class="sxs-lookup"><span data-stu-id="35eab-627">Object</span></span>| <span data-ttu-id="35eab-628">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-628">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-629">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="35eab-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="35eab-630">Object</span><span class="sxs-lookup"><span data-stu-id="35eab-630">Object</span></span> | <span data-ttu-id="35eab-631">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-631">&lt;optional&gt;</span></span> | <span data-ttu-id="35eab-632">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="35eab-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="35eab-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="35eab-633">Boolean</span></span> | <span data-ttu-id="35eab-634">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-634">&lt;optional&gt;</span></span> | <span data-ttu-id="35eab-635">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="35eab-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="35eab-636">function</span><span class="sxs-lookup"><span data-stu-id="35eab-636">function</span></span>| <span data-ttu-id="35eab-637">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-637">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-638">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35eab-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="35eab-639">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="35eab-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="35eab-640">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="35eab-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="35eab-641">Ошибки</span><span class="sxs-lookup"><span data-stu-id="35eab-641">Errors</span></span>

| <span data-ttu-id="35eab-642">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="35eab-642">Error code</span></span> | <span data-ttu-id="35eab-643">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="35eab-644">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="35eab-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="35eab-645">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="35eab-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="35eab-646">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="35eab-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="35eab-647">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-647">Requirements</span></span>

|<span data-ttu-id="35eab-648">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-648">Requirement</span></span>| <span data-ttu-id="35eab-649">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-650">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-651">1.1</span><span class="sxs-lookup"><span data-stu-id="35eab-651">1.1</span></span>|
|[<span data-ttu-id="35eab-652">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35eab-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="35eab-654">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-655">Создание</span><span class="sxs-lookup"><span data-stu-id="35eab-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="35eab-656">Примеры</span><span class="sxs-lookup"><span data-stu-id="35eab-656">Examples</span></span>

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

<span data-ttu-id="35eab-657">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="35eab-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35eab-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="35eab-659">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="35eab-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="35eab-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="35eab-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="35eab-663">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="35eab-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="35eab-664">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="35eab-664">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-665">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-665">Parameters</span></span>

|<span data-ttu-id="35eab-666">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-666">Name</span></span>| <span data-ttu-id="35eab-667">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-667">Type</span></span>| <span data-ttu-id="35eab-668">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="35eab-668">Attributes</span></span>| <span data-ttu-id="35eab-669">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="35eab-670">String</span><span class="sxs-lookup"><span data-stu-id="35eab-670">String</span></span>||<span data-ttu-id="35eab-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="35eab-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="35eab-673">String</span><span class="sxs-lookup"><span data-stu-id="35eab-673">String</span></span>||<span data-ttu-id="35eab-674">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-674">The subject of the item to be attached.</span></span> <span data-ttu-id="35eab-675">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="35eab-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="35eab-676">Object</span><span class="sxs-lookup"><span data-stu-id="35eab-676">Object</span></span>| <span data-ttu-id="35eab-677">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-677">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-678">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="35eab-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="35eab-679">Объект</span><span class="sxs-lookup"><span data-stu-id="35eab-679">Object</span></span>| <span data-ttu-id="35eab-680">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-680">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-681">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="35eab-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="35eab-682">функция</span><span class="sxs-lookup"><span data-stu-id="35eab-682">function</span></span>| <span data-ttu-id="35eab-683">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-683">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-684">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35eab-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="35eab-685">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="35eab-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="35eab-686">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="35eab-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="35eab-687">Ошибки</span><span class="sxs-lookup"><span data-stu-id="35eab-687">Errors</span></span>

| <span data-ttu-id="35eab-688">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="35eab-688">Error code</span></span> | <span data-ttu-id="35eab-689">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="35eab-690">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="35eab-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="35eab-691">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-691">Requirements</span></span>

|<span data-ttu-id="35eab-692">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-692">Requirement</span></span>| <span data-ttu-id="35eab-693">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-694">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-695">1.1</span><span class="sxs-lookup"><span data-stu-id="35eab-695">1.1</span></span>|
|[<span data-ttu-id="35eab-696">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35eab-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="35eab-698">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-699">Создание</span><span class="sxs-lookup"><span data-stu-id="35eab-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-700">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-700">Example</span></span>

<span data-ttu-id="35eab-701">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="35eab-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="35eab-702">close()</span><span class="sxs-lookup"><span data-stu-id="35eab-702">close()</span></span>

<span data-ttu-id="35eab-703">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="35eab-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="35eab-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="35eab-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-706">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="35eab-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="35eab-707">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="35eab-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-708">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-708">Requirements</span></span>

|<span data-ttu-id="35eab-709">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-709">Requirement</span></span>| <span data-ttu-id="35eab-710">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-711">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-712">1.3</span><span class="sxs-lookup"><span data-stu-id="35eab-712">1.3</span></span>|
|[<span data-ttu-id="35eab-713">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-714">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="35eab-714">Restricted</span></span>|
|[<span data-ttu-id="35eab-715">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-716">Создание</span><span class="sxs-lookup"><span data-stu-id="35eab-716">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="35eab-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="35eab-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="35eab-718">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-719">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-719">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="35eab-720">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="35eab-720">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="35eab-721">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="35eab-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="35eab-722">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="35eab-722">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="35eab-723">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="35eab-723">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="35eab-724">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="35eab-724">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-725">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-725">Parameters</span></span>

| <span data-ttu-id="35eab-726">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-726">Name</span></span> | <span data-ttu-id="35eab-727">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-727">Type</span></span> | <span data-ttu-id="35eab-728">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="35eab-728">Attributes</span></span> | <span data-ttu-id="35eab-729">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="35eab-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="35eab-730">String &#124; Object</span></span>| |<span data-ttu-id="35eab-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="35eab-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="35eab-733">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="35eab-733">**OR**</span></span><br/><span data-ttu-id="35eab-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="35eab-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="35eab-736">String.</span><span class="sxs-lookup"><span data-stu-id="35eab-736">String</span></span> | <span data-ttu-id="35eab-737">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-737">&lt;optional&gt;</span></span> | <span data-ttu-id="35eab-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="35eab-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="35eab-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="35eab-741">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-741">&lt;optional&gt;</span></span> | <span data-ttu-id="35eab-742">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="35eab-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="35eab-743">String.</span><span class="sxs-lookup"><span data-stu-id="35eab-743">String</span></span> | | <span data-ttu-id="35eab-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="35eab-746">Строка</span><span class="sxs-lookup"><span data-stu-id="35eab-746">String</span></span> | | <span data-ttu-id="35eab-747">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="35eab-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="35eab-748">String</span><span class="sxs-lookup"><span data-stu-id="35eab-748">String</span></span> | | <span data-ttu-id="35eab-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="35eab-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="35eab-751">Логический</span><span class="sxs-lookup"><span data-stu-id="35eab-751">Boolean</span></span> | | <span data-ttu-id="35eab-p144">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="35eab-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="35eab-754">String</span><span class="sxs-lookup"><span data-stu-id="35eab-754">String</span></span> | | <span data-ttu-id="35eab-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="35eab-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="35eab-758">function</span><span class="sxs-lookup"><span data-stu-id="35eab-758">function</span></span> | <span data-ttu-id="35eab-759">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-759">&lt;optional&gt;</span></span> | <span data-ttu-id="35eab-760">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35eab-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="35eab-761">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-761">Requirements</span></span>

|<span data-ttu-id="35eab-762">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-762">Requirement</span></span>| <span data-ttu-id="35eab-763">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-764">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-765">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-765">1.0</span></span>|
|[<span data-ttu-id="35eab-766">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-767">ReadItem</span></span>|
|[<span data-ttu-id="35eab-768">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-769">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="35eab-770">Примеры</span><span class="sxs-lookup"><span data-stu-id="35eab-770">Examples</span></span>

<span data-ttu-id="35eab-771">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="35eab-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="35eab-772">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-772">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="35eab-773">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-773">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="35eab-774">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="35eab-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="35eab-775">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="35eab-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="35eab-776">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="35eab-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="35eab-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="35eab-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="35eab-778">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-779">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-779">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="35eab-780">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="35eab-780">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="35eab-781">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="35eab-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="35eab-782">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="35eab-782">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="35eab-783">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="35eab-783">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="35eab-784">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="35eab-784">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-785">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-785">Parameters</span></span>

| <span data-ttu-id="35eab-786">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-786">Name</span></span> | <span data-ttu-id="35eab-787">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-787">Type</span></span> | <span data-ttu-id="35eab-788">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="35eab-788">Attributes</span></span> | <span data-ttu-id="35eab-789">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="35eab-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="35eab-790">String &#124; Object</span></span>| | <span data-ttu-id="35eab-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="35eab-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="35eab-793">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="35eab-793">**OR**</span></span><br/><span data-ttu-id="35eab-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="35eab-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="35eab-796">String.</span><span class="sxs-lookup"><span data-stu-id="35eab-796">String</span></span> | <span data-ttu-id="35eab-797">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-797">&lt;optional&gt;</span></span> | <span data-ttu-id="35eab-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="35eab-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="35eab-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="35eab-801">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-801">&lt;optional&gt;</span></span> | <span data-ttu-id="35eab-802">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="35eab-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="35eab-803">String.</span><span class="sxs-lookup"><span data-stu-id="35eab-803">String</span></span> | | <span data-ttu-id="35eab-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="35eab-806">Строка</span><span class="sxs-lookup"><span data-stu-id="35eab-806">String</span></span> | | <span data-ttu-id="35eab-807">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="35eab-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="35eab-808">String</span><span class="sxs-lookup"><span data-stu-id="35eab-808">String</span></span> | | <span data-ttu-id="35eab-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="35eab-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="35eab-811">Логический</span><span class="sxs-lookup"><span data-stu-id="35eab-811">Boolean</span></span> | | <span data-ttu-id="35eab-p152">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="35eab-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="35eab-814">String</span><span class="sxs-lookup"><span data-stu-id="35eab-814">String</span></span> | | <span data-ttu-id="35eab-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="35eab-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="35eab-818">function</span><span class="sxs-lookup"><span data-stu-id="35eab-818">function</span></span> | <span data-ttu-id="35eab-819">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-819">&lt;optional&gt;</span></span> | <span data-ttu-id="35eab-820">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35eab-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="35eab-821">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-821">Requirements</span></span>

|<span data-ttu-id="35eab-822">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-822">Requirement</span></span>| <span data-ttu-id="35eab-823">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-824">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-825">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-825">1.0</span></span>|
|[<span data-ttu-id="35eab-826">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-827">ReadItem</span></span>|
|[<span data-ttu-id="35eab-828">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-829">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="35eab-830">Примеры</span><span class="sxs-lookup"><span data-stu-id="35eab-830">Examples</span></span>

<span data-ttu-id="35eab-831">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="35eab-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="35eab-832">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-832">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="35eab-833">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-833">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="35eab-834">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="35eab-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="35eab-835">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="35eab-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="35eab-836">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="35eab-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="35eab-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="35eab-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="35eab-838">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-839">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-840">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-840">Requirements</span></span>

|<span data-ttu-id="35eab-841">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-841">Requirement</span></span>| <span data-ttu-id="35eab-842">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-843">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-844">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-844">1.0</span></span>|
|[<span data-ttu-id="35eab-845">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-846">ReadItem</span></span>|
|[<span data-ttu-id="35eab-847">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-848">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35eab-849">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="35eab-849">Returns:</span></span>

<span data-ttu-id="35eab-850">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-850">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="35eab-851">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-851">Example</span></span>

<span data-ttu-id="35eab-852">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-852">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="35eab-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="35eab-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="35eab-854">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-855">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-855">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-856">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-856">Parameters</span></span>

|<span data-ttu-id="35eab-857">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-857">Name</span></span>| <span data-ttu-id="35eab-858">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-858">Type</span></span>| <span data-ttu-id="35eab-859">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="35eab-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="35eab-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="35eab-861">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="35eab-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35eab-862">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-862">Requirements</span></span>

|<span data-ttu-id="35eab-863">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-863">Requirement</span></span>| <span data-ttu-id="35eab-864">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-865">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-866">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-866">1.0</span></span>|
|[<span data-ttu-id="35eab-867">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-868">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="35eab-868">Restricted</span></span>|
|[<span data-ttu-id="35eab-869">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-870">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35eab-871">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="35eab-871">Returns:</span></span>

<span data-ttu-id="35eab-872">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="35eab-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="35eab-873">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="35eab-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="35eab-874">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="35eab-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="35eab-875">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="35eab-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="35eab-876">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="35eab-876">Value of `entityType`</span></span> | <span data-ttu-id="35eab-877">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="35eab-877">Type of objects in returned array</span></span> | <span data-ttu-id="35eab-878">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="35eab-879">String</span><span class="sxs-lookup"><span data-stu-id="35eab-879">String</span></span> | <span data-ttu-id="35eab-880">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="35eab-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="35eab-881">Contact</span><span class="sxs-lookup"><span data-stu-id="35eab-881">Contact</span></span> | <span data-ttu-id="35eab-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="35eab-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="35eab-883">String</span><span class="sxs-lookup"><span data-stu-id="35eab-883">String</span></span> | <span data-ttu-id="35eab-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="35eab-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="35eab-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="35eab-885">MeetingSuggestion</span></span> | <span data-ttu-id="35eab-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="35eab-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="35eab-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="35eab-887">PhoneNumber</span></span> | <span data-ttu-id="35eab-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="35eab-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="35eab-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="35eab-889">TaskSuggestion</span></span> | <span data-ttu-id="35eab-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="35eab-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="35eab-891">String</span><span class="sxs-lookup"><span data-stu-id="35eab-891">String</span></span> | <span data-ttu-id="35eab-892">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="35eab-892">**Restricted**</span></span> |

<span data-ttu-id="35eab-893">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="35eab-893">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="35eab-894">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-894">Example</span></span>

<span data-ttu-id="35eab-895">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="35eab-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="35eab-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="35eab-897">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="35eab-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-898">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-898">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="35eab-899">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="35eab-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-900">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-900">Parameters</span></span>

|<span data-ttu-id="35eab-901">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-901">Name</span></span>| <span data-ttu-id="35eab-902">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-902">Type</span></span>| <span data-ttu-id="35eab-903">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="35eab-904">String</span><span class="sxs-lookup"><span data-stu-id="35eab-904">String</span></span>|<span data-ttu-id="35eab-905">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="35eab-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35eab-906">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-906">Requirements</span></span>

|<span data-ttu-id="35eab-907">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-907">Requirement</span></span>| <span data-ttu-id="35eab-908">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-909">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-910">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-910">1.0</span></span>|
|[<span data-ttu-id="35eab-911">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-912">ReadItem</span></span>|
|[<span data-ttu-id="35eab-913">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-914">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35eab-915">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="35eab-915">Returns:</span></span>

<span data-ttu-id="35eab-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="35eab-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="35eab-918">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="35eab-918">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="35eab-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="35eab-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="35eab-920">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="35eab-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-921">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-921">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="35eab-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="35eab-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="35eab-925">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="35eab-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="35eab-926">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="35eab-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="35eab-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="35eab-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="35eab-930">Requirements</span></span>

|<span data-ttu-id="35eab-931">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-931">Requirement</span></span>| <span data-ttu-id="35eab-932">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-933">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-934">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-934">1.0</span></span>|
|[<span data-ttu-id="35eab-935">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-936">ReadItem</span></span>|
|[<span data-ttu-id="35eab-937">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-938">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35eab-939">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="35eab-939">Returns:</span></span>

<span data-ttu-id="35eab-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="35eab-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="35eab-942">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="35eab-942">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="35eab-943">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-943">Example</span></span>

<span data-ttu-id="35eab-944">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="35eab-944">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="35eab-945">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="35eab-945">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="35eab-946">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="35eab-946">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-947">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="35eab-948">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="35eab-948">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="35eab-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="35eab-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-951">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-951">Parameters</span></span>

|<span data-ttu-id="35eab-952">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-952">Name</span></span>| <span data-ttu-id="35eab-953">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-953">Type</span></span>| <span data-ttu-id="35eab-954">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-954">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="35eab-955">String</span><span class="sxs-lookup"><span data-stu-id="35eab-955">String</span></span>|<span data-ttu-id="35eab-956">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="35eab-956">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35eab-957">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-957">Requirements</span></span>

|<span data-ttu-id="35eab-958">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-958">Requirement</span></span>| <span data-ttu-id="35eab-959">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-959">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-960">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-960">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-961">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-961">1.0</span></span>|
|[<span data-ttu-id="35eab-962">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-962">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-963">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-963">ReadItem</span></span>|
|[<span data-ttu-id="35eab-964">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-964">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-965">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-965">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35eab-966">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="35eab-966">Returns:</span></span>

<span data-ttu-id="35eab-967">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="35eab-967">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="35eab-968">Тип: Array. < String ></span><span class="sxs-lookup"><span data-stu-id="35eab-968">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="35eab-969">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-969">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="35eab-970">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="35eab-970">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="35eab-971">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-971">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="35eab-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="35eab-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-974">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-974">Parameters</span></span>

|<span data-ttu-id="35eab-975">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-975">Name</span></span>| <span data-ttu-id="35eab-976">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-976">Type</span></span>| <span data-ttu-id="35eab-977">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="35eab-977">Attributes</span></span>| <span data-ttu-id="35eab-978">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-978">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="35eab-979">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="35eab-979">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="35eab-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="35eab-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="35eab-983">Object</span><span class="sxs-lookup"><span data-stu-id="35eab-983">Object</span></span>| <span data-ttu-id="35eab-984">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-984">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-985">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="35eab-985">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="35eab-986">Объект</span><span class="sxs-lookup"><span data-stu-id="35eab-986">Object</span></span>| <span data-ttu-id="35eab-987">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-987">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-988">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="35eab-988">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="35eab-989">функция</span><span class="sxs-lookup"><span data-stu-id="35eab-989">function</span></span>||<span data-ttu-id="35eab-990">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35eab-990">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="35eab-991">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="35eab-991">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="35eab-992">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="35eab-992">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35eab-993">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-993">Requirements</span></span>

|<span data-ttu-id="35eab-994">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-994">Requirement</span></span>| <span data-ttu-id="35eab-995">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-995">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-996">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-996">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-997">1.2</span><span class="sxs-lookup"><span data-stu-id="35eab-997">1.2</span></span>|
|[<span data-ttu-id="35eab-998">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-998">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-999">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35eab-999">ReadWriteItem</span></span>|
|[<span data-ttu-id="35eab-1000">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-1000">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-1001">Создание</span><span class="sxs-lookup"><span data-stu-id="35eab-1001">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="35eab-1002">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="35eab-1002">Returns:</span></span>

<span data-ttu-id="35eab-1003">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="35eab-1003">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="35eab-1004">Тип: String</span><span class="sxs-lookup"><span data-stu-id="35eab-1004">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="35eab-1005">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-1005">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="35eab-1006">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="35eab-1006">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="35eab-1007">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="35eab-1007">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="35eab-1008">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="35eab-1008">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-1009">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-1009">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-1010">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-1010">Requirements</span></span>

|<span data-ttu-id="35eab-1011">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-1011">Requirement</span></span>| <span data-ttu-id="35eab-1012">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-1013">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-1014">1.6</span><span class="sxs-lookup"><span data-stu-id="35eab-1014">1.6</span></span> |
|[<span data-ttu-id="35eab-1015">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-1015">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-1016">ReadItem</span></span>|
|[<span data-ttu-id="35eab-1017">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-1017">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-1018">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-1018">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35eab-1019">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="35eab-1019">Returns:</span></span>

<span data-ttu-id="35eab-1020">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="35eab-1020">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="35eab-1021">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-1021">Example</span></span>

<span data-ttu-id="35eab-1022">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="35eab-1022">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="35eab-1023">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="35eab-1023">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="35eab-p164">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="35eab-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-1026">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="35eab-1026">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="35eab-p165">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="35eab-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="35eab-1030">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="35eab-1030">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="35eab-1031">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="35eab-1031">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="35eab-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="35eab-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35eab-1035">Requirements</span><span class="sxs-lookup"><span data-stu-id="35eab-1035">Requirements</span></span>

|<span data-ttu-id="35eab-1036">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-1036">Requirement</span></span>| <span data-ttu-id="35eab-1037">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-1037">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-1038">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-1038">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-1039">1.6</span><span class="sxs-lookup"><span data-stu-id="35eab-1039">1.6</span></span> |
|[<span data-ttu-id="35eab-1040">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-1040">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-1041">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-1041">ReadItem</span></span>|
|[<span data-ttu-id="35eab-1042">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-1042">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-1043">Чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-1043">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35eab-1044">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="35eab-1044">Returns:</span></span>

<span data-ttu-id="35eab-p167">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="35eab-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="35eab-1047">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-1047">Example</span></span>

<span data-ttu-id="35eab-1048">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="35eab-1048">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="35eab-1049">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="35eab-1049">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="35eab-1050">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-1050">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="35eab-p168">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="35eab-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-1054">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-1054">Parameters</span></span>

|<span data-ttu-id="35eab-1055">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-1055">Name</span></span>| <span data-ttu-id="35eab-1056">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-1056">Type</span></span>| <span data-ttu-id="35eab-1057">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="35eab-1057">Attributes</span></span>| <span data-ttu-id="35eab-1058">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-1058">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="35eab-1059">function</span><span class="sxs-lookup"><span data-stu-id="35eab-1059">function</span></span>||<span data-ttu-id="35eab-1060">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35eab-1060">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="35eab-1061">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="35eab-1061">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="35eab-1062">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="35eab-1062">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="35eab-1063">Объект</span><span class="sxs-lookup"><span data-stu-id="35eab-1063">Object</span></span>| <span data-ttu-id="35eab-1064">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-1065">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="35eab-1065">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="35eab-1066">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="35eab-1066">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35eab-1067">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-1067">Requirements</span></span>

|<span data-ttu-id="35eab-1068">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-1068">Requirement</span></span>| <span data-ttu-id="35eab-1069">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-1069">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-1070">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-1070">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-1071">1.0</span><span class="sxs-lookup"><span data-stu-id="35eab-1071">1.0</span></span>|
|[<span data-ttu-id="35eab-1072">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-1072">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-1073">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35eab-1073">ReadItem</span></span>|
|[<span data-ttu-id="35eab-1074">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-1074">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-1075">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="35eab-1075">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-1076">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-1076">Example</span></span>

<span data-ttu-id="35eab-p171">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="35eab-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="35eab-1080">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35eab-1080">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="35eab-1081">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="35eab-1081">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="35eab-1082">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="35eab-1082">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="35eab-1083">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="35eab-1083">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="35eab-1084">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="35eab-1084">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="35eab-1085">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="35eab-1085">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-1086">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-1086">Parameters</span></span>

|<span data-ttu-id="35eab-1087">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-1087">Name</span></span>| <span data-ttu-id="35eab-1088">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-1088">Type</span></span>| <span data-ttu-id="35eab-1089">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="35eab-1089">Attributes</span></span>| <span data-ttu-id="35eab-1090">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-1090">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="35eab-1091">String</span><span class="sxs-lookup"><span data-stu-id="35eab-1091">String</span></span>||<span data-ttu-id="35eab-1092">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="35eab-1092">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="35eab-1093">Object</span><span class="sxs-lookup"><span data-stu-id="35eab-1093">Object</span></span>| <span data-ttu-id="35eab-1094">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-1094">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-1095">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="35eab-1095">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="35eab-1096">Объект</span><span class="sxs-lookup"><span data-stu-id="35eab-1096">Object</span></span>| <span data-ttu-id="35eab-1097">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-1098">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="35eab-1098">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="35eab-1099">функция</span><span class="sxs-lookup"><span data-stu-id="35eab-1099">function</span></span>| <span data-ttu-id="35eab-1100">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-1101">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35eab-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="35eab-1102">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="35eab-1102">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="35eab-1103">Ошибки</span><span class="sxs-lookup"><span data-stu-id="35eab-1103">Errors</span></span>

| <span data-ttu-id="35eab-1104">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="35eab-1104">Error code</span></span> | <span data-ttu-id="35eab-1105">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-1105">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="35eab-1106">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="35eab-1106">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="35eab-1107">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-1107">Requirements</span></span>

|<span data-ttu-id="35eab-1108">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-1108">Requirement</span></span>| <span data-ttu-id="35eab-1109">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-1109">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-1110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="35eab-1110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-1111">1.1</span><span class="sxs-lookup"><span data-stu-id="35eab-1111">1.1</span></span>|
|[<span data-ttu-id="35eab-1112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-1112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-1113">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35eab-1113">ReadWriteItem</span></span>|
|[<span data-ttu-id="35eab-1114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-1114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-1115">Создание</span><span class="sxs-lookup"><span data-stu-id="35eab-1115">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-1116">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-1116">Example</span></span>

<span data-ttu-id="35eab-1117">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="35eab-1117">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="35eab-1118">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="35eab-1118">saveAsync([options], callback)</span></span>

<span data-ttu-id="35eab-1119">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="35eab-1119">Asynchronously saves an item.</span></span>

<span data-ttu-id="35eab-1120">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="35eab-1120">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="35eab-1121">В Outlook в Интернете или Outlook в интерактивном режиме элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="35eab-1121">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="35eab-1122">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="35eab-1122">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-1123">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="35eab-1123">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="35eab-1124">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="35eab-1124">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="35eab-p175">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="35eab-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="35eab-1128">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="35eab-1128">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="35eab-1129">Outlook в Mac не поддерживает сохранение собраний.</span><span class="sxs-lookup"><span data-stu-id="35eab-1129">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="35eab-1130">`saveAsync` Метод завершается с ошибкой при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="35eab-1130">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="35eab-1131">Просмотреть [не удается сохранить собрание в виде черновика в Outlook для Mac с помощью API Office JS](https://support.microsoft.com/help/4505745) для обхода.</span><span class="sxs-lookup"><span data-stu-id="35eab-1131">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="35eab-1132">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="35eab-1132">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-1133">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-1133">Parameters</span></span>

|<span data-ttu-id="35eab-1134">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-1134">Name</span></span>| <span data-ttu-id="35eab-1135">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-1135">Type</span></span>| <span data-ttu-id="35eab-1136">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="35eab-1136">Attributes</span></span>| <span data-ttu-id="35eab-1137">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-1137">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="35eab-1138">Объект</span><span class="sxs-lookup"><span data-stu-id="35eab-1138">Object</span></span>| <span data-ttu-id="35eab-1139">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-1139">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-1140">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="35eab-1140">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="35eab-1141">Объект</span><span class="sxs-lookup"><span data-stu-id="35eab-1141">Object</span></span>| <span data-ttu-id="35eab-1142">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-1143">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="35eab-1143">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="35eab-1144">функция</span><span class="sxs-lookup"><span data-stu-id="35eab-1144">function</span></span>||<span data-ttu-id="35eab-1145">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35eab-1145">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="35eab-1146">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="35eab-1146">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35eab-1147">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-1147">Requirements</span></span>

|<span data-ttu-id="35eab-1148">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-1148">Requirement</span></span>| <span data-ttu-id="35eab-1149">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-1149">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-1150">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-1150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-1151">1.3</span><span class="sxs-lookup"><span data-stu-id="35eab-1151">1.3</span></span>|
|[<span data-ttu-id="35eab-1152">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-1152">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-1153">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35eab-1153">ReadWriteItem</span></span>|
|[<span data-ttu-id="35eab-1154">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-1154">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-1155">Создание</span><span class="sxs-lookup"><span data-stu-id="35eab-1155">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="35eab-1156">Примеры</span><span class="sxs-lookup"><span data-stu-id="35eab-1156">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="35eab-p177">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="35eab-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="35eab-1159">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="35eab-1159">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="35eab-1160">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="35eab-1160">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="35eab-p178">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="35eab-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35eab-1164">Параметры</span><span class="sxs-lookup"><span data-stu-id="35eab-1164">Parameters</span></span>

|<span data-ttu-id="35eab-1165">Имя</span><span class="sxs-lookup"><span data-stu-id="35eab-1165">Name</span></span>| <span data-ttu-id="35eab-1166">Тип</span><span class="sxs-lookup"><span data-stu-id="35eab-1166">Type</span></span>| <span data-ttu-id="35eab-1167">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="35eab-1167">Attributes</span></span>| <span data-ttu-id="35eab-1168">Описание</span><span class="sxs-lookup"><span data-stu-id="35eab-1168">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="35eab-1169">String</span><span class="sxs-lookup"><span data-stu-id="35eab-1169">String</span></span>||<span data-ttu-id="35eab-p179">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="35eab-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="35eab-1173">Object</span><span class="sxs-lookup"><span data-stu-id="35eab-1173">Object</span></span>| <span data-ttu-id="35eab-1174">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-1174">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-1175">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="35eab-1175">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="35eab-1176">Объект</span><span class="sxs-lookup"><span data-stu-id="35eab-1176">Object</span></span>| <span data-ttu-id="35eab-1177">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-1178">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="35eab-1178">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="35eab-1179">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="35eab-1179">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="35eab-1180">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="35eab-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="35eab-1181">Если `text`текущий стиль применяется в Outlook для веб-клиентов и клиентов для настольных ПК.</span><span class="sxs-lookup"><span data-stu-id="35eab-1181">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="35eab-1182">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="35eab-1182">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="35eab-1183">Если `html` и поле поддерживает HTML (тема не используется), текущий стиль применяется в Outlook в Интернете, а в настольных клиентах Outlook применяется стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="35eab-1183">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="35eab-1184">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="35eab-1184">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="35eab-1185">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="35eab-1185">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="35eab-1186">функция</span><span class="sxs-lookup"><span data-stu-id="35eab-1186">function</span></span>||<span data-ttu-id="35eab-1187">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35eab-1187">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="35eab-1188">Требования</span><span class="sxs-lookup"><span data-stu-id="35eab-1188">Requirements</span></span>

|<span data-ttu-id="35eab-1189">Требование</span><span class="sxs-lookup"><span data-stu-id="35eab-1189">Requirement</span></span>| <span data-ttu-id="35eab-1190">Значение</span><span class="sxs-lookup"><span data-stu-id="35eab-1190">Value</span></span>|
|---|---|
|[<span data-ttu-id="35eab-1191">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="35eab-1191">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35eab-1192">1.2</span><span class="sxs-lookup"><span data-stu-id="35eab-1192">1.2</span></span>|
|[<span data-ttu-id="35eab-1193">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="35eab-1193">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35eab-1194">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="35eab-1194">ReadWriteItem</span></span>|
|[<span data-ttu-id="35eab-1195">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="35eab-1195">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="35eab-1196">Создание</span><span class="sxs-lookup"><span data-stu-id="35eab-1196">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="35eab-1197">Пример</span><span class="sxs-lookup"><span data-stu-id="35eab-1197">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
