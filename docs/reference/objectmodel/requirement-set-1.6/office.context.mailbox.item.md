---
title: Office. Context. Mailbox. Item — набор требований 1,6
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 46dc6148ea150e9e2ab1b245ead006a2ad377d88
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629702"
---
# <a name="item"></a><span data-ttu-id="4b19d-102">item</span><span class="sxs-lookup"><span data-stu-id="4b19d-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4b19d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4b19d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4b19d-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="4b19d-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-106">Requirements</span></span>

|<span data-ttu-id="4b19d-107">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-107">Requirement</span></span>| <span data-ttu-id="4b19d-108">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-110">1.0</span></span>|
|[<span data-ttu-id="4b19d-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4b19d-112">Restricted</span></span>|
|[<span data-ttu-id="4b19d-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4b19d-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="4b19d-115">Members and methods</span></span>

| <span data-ttu-id="4b19d-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-116">Member</span></span> | <span data-ttu-id="4b19d-117">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4b19d-118">attachments</span><span class="sxs-lookup"><span data-stu-id="4b19d-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="4b19d-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-119">Member</span></span> |
| [<span data-ttu-id="4b19d-120">bcc</span><span class="sxs-lookup"><span data-stu-id="4b19d-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="4b19d-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-121">Member</span></span> |
| [<span data-ttu-id="4b19d-122">body</span><span class="sxs-lookup"><span data-stu-id="4b19d-122">body</span></span>](#body-body) | <span data-ttu-id="4b19d-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-123">Member</span></span> |
| [<span data-ttu-id="4b19d-124">cc</span><span class="sxs-lookup"><span data-stu-id="4b19d-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4b19d-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-125">Member</span></span> |
| [<span data-ttu-id="4b19d-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="4b19d-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4b19d-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-127">Member</span></span> |
| [<span data-ttu-id="4b19d-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4b19d-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4b19d-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-129">Member</span></span> |
| [<span data-ttu-id="4b19d-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4b19d-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4b19d-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-131">Member</span></span> |
| [<span data-ttu-id="4b19d-132">end</span><span class="sxs-lookup"><span data-stu-id="4b19d-132">end</span></span>](#end-datetime) | <span data-ttu-id="4b19d-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-133">Member</span></span> |
| [<span data-ttu-id="4b19d-134">from</span><span class="sxs-lookup"><span data-stu-id="4b19d-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="4b19d-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-135">Member</span></span> |
| [<span data-ttu-id="4b19d-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4b19d-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4b19d-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-137">Member</span></span> |
| [<span data-ttu-id="4b19d-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="4b19d-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4b19d-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-139">Member</span></span> |
| [<span data-ttu-id="4b19d-140">itemId</span><span class="sxs-lookup"><span data-stu-id="4b19d-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4b19d-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-141">Member</span></span> |
| [<span data-ttu-id="4b19d-142">itemType</span><span class="sxs-lookup"><span data-stu-id="4b19d-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="4b19d-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-143">Member</span></span> |
| [<span data-ttu-id="4b19d-144">location</span><span class="sxs-lookup"><span data-stu-id="4b19d-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="4b19d-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-145">Member</span></span> |
| [<span data-ttu-id="4b19d-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4b19d-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4b19d-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-147">Member</span></span> |
| [<span data-ttu-id="4b19d-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4b19d-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="4b19d-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-149">Member</span></span> |
| [<span data-ttu-id="4b19d-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4b19d-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4b19d-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-151">Member</span></span> |
| [<span data-ttu-id="4b19d-152">organizer</span><span class="sxs-lookup"><span data-stu-id="4b19d-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="4b19d-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-153">Member</span></span> |
| [<span data-ttu-id="4b19d-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4b19d-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4b19d-155">Member</span><span class="sxs-lookup"><span data-stu-id="4b19d-155">Member</span></span> |
| [<span data-ttu-id="4b19d-156">sender</span><span class="sxs-lookup"><span data-stu-id="4b19d-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="4b19d-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-157">Member</span></span> |
| [<span data-ttu-id="4b19d-158">start</span><span class="sxs-lookup"><span data-stu-id="4b19d-158">start</span></span>](#start-datetime) | <span data-ttu-id="4b19d-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-159">Member</span></span> |
| [<span data-ttu-id="4b19d-160">subject</span><span class="sxs-lookup"><span data-stu-id="4b19d-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="4b19d-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-161">Member</span></span> |
| [<span data-ttu-id="4b19d-162">to</span><span class="sxs-lookup"><span data-stu-id="4b19d-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4b19d-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="4b19d-163">Member</span></span> |
| [<span data-ttu-id="4b19d-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4b19d-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4b19d-165">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-165">Method</span></span> |
| [<span data-ttu-id="4b19d-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4b19d-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4b19d-167">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-167">Method</span></span> |
| [<span data-ttu-id="4b19d-168">close</span><span class="sxs-lookup"><span data-stu-id="4b19d-168">close</span></span>](#close) | <span data-ttu-id="4b19d-169">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-169">Method</span></span> |
| [<span data-ttu-id="4b19d-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4b19d-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="4b19d-171">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-171">Method</span></span> |
| [<span data-ttu-id="4b19d-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4b19d-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="4b19d-173">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-173">Method</span></span> |
| [<span data-ttu-id="4b19d-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="4b19d-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="4b19d-175">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-175">Method</span></span> |
| [<span data-ttu-id="4b19d-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4b19d-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4b19d-177">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-177">Method</span></span> |
| [<span data-ttu-id="4b19d-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4b19d-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4b19d-179">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-179">Method</span></span> |
| [<span data-ttu-id="4b19d-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4b19d-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4b19d-181">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-181">Method</span></span> |
| [<span data-ttu-id="4b19d-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4b19d-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4b19d-183">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-183">Method</span></span> |
| [<span data-ttu-id="4b19d-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4b19d-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4b19d-185">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-185">Method</span></span> |
| [<span data-ttu-id="4b19d-186">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="4b19d-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="4b19d-187">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-187">Method</span></span> |
| [<span data-ttu-id="4b19d-188">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="4b19d-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="4b19d-189">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-189">Method</span></span> |
| [<span data-ttu-id="4b19d-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4b19d-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4b19d-191">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-191">Method</span></span> |
| [<span data-ttu-id="4b19d-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4b19d-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4b19d-193">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-193">Method</span></span> |
| [<span data-ttu-id="4b19d-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4b19d-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4b19d-195">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-195">Method</span></span> |
| [<span data-ttu-id="4b19d-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4b19d-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4b19d-197">Метод</span><span class="sxs-lookup"><span data-stu-id="4b19d-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="4b19d-198">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-198">Example</span></span>

<span data-ttu-id="4b19d-199">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="4b19d-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="4b19d-200">Members</span><span class="sxs-lookup"><span data-stu-id="4b19d-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="4b19d-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="4b19d-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="4b19d-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-204">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="4b19d-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4b19d-205">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="4b19d-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-206">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-206">Type</span></span>

*   <span data-ttu-id="4b19d-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="4b19d-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-208">Requirements</span></span>

|<span data-ttu-id="4b19d-209">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-209">Requirement</span></span>| <span data-ttu-id="4b19d-210">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-211">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-212">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-212">1.0</span></span>|
|[<span data-ttu-id="4b19d-213">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-214">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-215">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-216">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-217">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-217">Example</span></span>

<span data-ttu-id="4b19d-218">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4b19d-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-220">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4b19d-221">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-221">Compose mode only.</span></span>

<span data-ttu-id="4b19d-222">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-222">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4b19d-223">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-223">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4b19d-224">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-224">Get 500 members maximum.</span></span>
- <span data-ttu-id="4b19d-225">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-225">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-226">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-226">Type</span></span>

*   [<span data-ttu-id="4b19d-227">Получатели</span><span class="sxs-lookup"><span data-stu-id="4b19d-227">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4b19d-228">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-228">Requirements</span></span>

|<span data-ttu-id="4b19d-229">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-229">Requirement</span></span>| <span data-ttu-id="4b19d-230">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-231">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-232">1.1</span><span class="sxs-lookup"><span data-stu-id="4b19d-232">1.1</span></span>|
|[<span data-ttu-id="4b19d-233">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-233">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-234">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-235">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-236">Создание</span><span class="sxs-lookup"><span data-stu-id="4b19d-236">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-237">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-237">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="4b19d-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-239">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-239">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-240">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-240">Type</span></span>

*   [<span data-ttu-id="4b19d-241">Body</span><span class="sxs-lookup"><span data-stu-id="4b19d-241">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4b19d-242">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-242">Requirements</span></span>

|<span data-ttu-id="4b19d-243">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-243">Requirement</span></span>| <span data-ttu-id="4b19d-244">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-245">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-246">1.1</span><span class="sxs-lookup"><span data-stu-id="4b19d-246">1.1</span></span>|
|[<span data-ttu-id="4b19d-247">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-248">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-249">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-250">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-251">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-251">Example</span></span>

<span data-ttu-id="4b19d-252">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="4b19d-252">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="4b19d-253">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b19d-253">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4b19d-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-255">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4b19d-256">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4b19d-257">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4b19d-257">Read mode</span></span>

<span data-ttu-id="4b19d-258">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-258">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="4b19d-259">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4b19d-260">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-260">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="4b19d-261">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4b19d-261">Compose mode</span></span>

<span data-ttu-id="4b19d-262">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="4b19d-263">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-263">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4b19d-264">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-264">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4b19d-265">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-265">Get 500 members maximum.</span></span>
- <span data-ttu-id="4b19d-266">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-266">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4b19d-267">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-267">Type</span></span>

*   <span data-ttu-id="4b19d-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-269">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-269">Requirements</span></span>

|<span data-ttu-id="4b19d-270">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-270">Requirement</span></span>| <span data-ttu-id="4b19d-271">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-272">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-272">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-273">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-273">1.0</span></span>|
|[<span data-ttu-id="4b19d-274">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-274">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-275">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-275">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-276">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-276">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-277">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-277">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="4b19d-278">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="4b19d-278">(nullable) conversationId: String</span></span>

<span data-ttu-id="4b19d-279">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="4b19d-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4b19d-p109">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4b19d-p110">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-284">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-284">Type</span></span>

*   <span data-ttu-id="4b19d-285">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-286">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-286">Requirements</span></span>

|<span data-ttu-id="4b19d-287">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-287">Requirement</span></span>| <span data-ttu-id="4b19d-288">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-289">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-289">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-290">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-290">1.0</span></span>|
|[<span data-ttu-id="4b19d-291">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-291">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-292">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-293">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-293">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-294">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-294">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-295">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-295">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="4b19d-296">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="4b19d-296">dateTimeCreated: Date</span></span>

<span data-ttu-id="4b19d-p111">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-299">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-299">Type</span></span>

*   <span data-ttu-id="4b19d-300">Дата</span><span class="sxs-lookup"><span data-stu-id="4b19d-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-301">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-301">Requirements</span></span>

|<span data-ttu-id="4b19d-302">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-302">Requirement</span></span>| <span data-ttu-id="4b19d-303">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-304">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-305">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-305">1.0</span></span>|
|[<span data-ttu-id="4b19d-306">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-307">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-308">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-309">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-310">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-310">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="4b19d-311">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="4b19d-311">dateTimeModified: Date</span></span>

<span data-ttu-id="4b19d-p112">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-314">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-314">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-315">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-315">Type</span></span>

*   <span data-ttu-id="4b19d-316">Дата</span><span class="sxs-lookup"><span data-stu-id="4b19d-316">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-317">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-317">Requirements</span></span>

|<span data-ttu-id="4b19d-318">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-318">Requirement</span></span>| <span data-ttu-id="4b19d-319">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-320">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-321">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-321">1.0</span></span>|
|[<span data-ttu-id="4b19d-322">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-323">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-324">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-325">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-325">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-326">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-326">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="4b19d-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-328">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-328">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4b19d-p113">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="4b19d-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4b19d-331">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4b19d-331">Read mode</span></span>

<span data-ttu-id="4b19d-332">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-332">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="4b19d-333">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4b19d-333">Compose mode</span></span>

<span data-ttu-id="4b19d-334">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-334">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4b19d-335">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="4b19d-335">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4b19d-336">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-336">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4b19d-337">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-337">Type</span></span>

*   <span data-ttu-id="4b19d-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-339">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-339">Requirements</span></span>

|<span data-ttu-id="4b19d-340">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-340">Requirement</span></span>| <span data-ttu-id="4b19d-341">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-342">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-343">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-343">1.0</span></span>|
|[<span data-ttu-id="4b19d-344">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-344">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-345">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-346">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-346">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-347">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-347">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="4b19d-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-p114">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="4b19d-p115">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-353">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-353">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-354">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-354">Type</span></span>

*   [<span data-ttu-id="4b19d-355">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4b19d-355">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="4b19d-356">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="4b19d-357">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-357">Requirements</span></span>

|<span data-ttu-id="4b19d-358">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-358">Requirement</span></span>| <span data-ttu-id="4b19d-359">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-360">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-361">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-361">1.0</span></span>|
|[<span data-ttu-id="4b19d-362">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-363">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-364">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-365">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-365">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="4b19d-366">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="4b19d-366">internetMessageId: String</span></span>

<span data-ttu-id="4b19d-p116">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-369">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-369">Type</span></span>

*   <span data-ttu-id="4b19d-370">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-371">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-371">Requirements</span></span>

|<span data-ttu-id="4b19d-372">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-372">Requirement</span></span>| <span data-ttu-id="4b19d-373">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-374">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-375">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-375">1.0</span></span>|
|[<span data-ttu-id="4b19d-376">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-376">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-377">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-378">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-378">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-379">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-380">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-380">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="4b19d-381">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="4b19d-381">itemClass: String</span></span>

<span data-ttu-id="4b19d-p117">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4b19d-p118">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="4b19d-386">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-386">Type</span></span> | <span data-ttu-id="4b19d-387">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-387">Description</span></span> | <span data-ttu-id="4b19d-388">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="4b19d-388">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="4b19d-389">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="4b19d-389">Appointment items</span></span> | <span data-ttu-id="4b19d-390">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="4b19d-391">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="4b19d-391">Message items</span></span> | <span data-ttu-id="4b19d-392">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="4b19d-393">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-394">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-394">Type</span></span>

*   <span data-ttu-id="4b19d-395">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-396">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-396">Requirements</span></span>

|<span data-ttu-id="4b19d-397">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-397">Requirement</span></span>| <span data-ttu-id="4b19d-398">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-399">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-400">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-400">1.0</span></span>|
|[<span data-ttu-id="4b19d-401">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-401">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-402">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-403">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-403">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-404">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-405">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-405">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4b19d-406">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="4b19d-406">(nullable) itemId: String</span></span>

<span data-ttu-id="4b19d-p119">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-409">Идентификатор, возвращаемый свойством `itemId`, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="4b19d-409">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="4b19d-410">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="4b19d-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4b19d-411">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4b19d-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4b19d-412">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="4b19d-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4b19d-p121">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-415">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-415">Type</span></span>

*   <span data-ttu-id="4b19d-416">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-417">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-417">Requirements</span></span>

|<span data-ttu-id="4b19d-418">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-418">Requirement</span></span>| <span data-ttu-id="4b19d-419">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-420">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-421">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-421">1.0</span></span>|
|[<span data-ttu-id="4b19d-422">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-422">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-423">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-424">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-424">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-425">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-426">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-426">Example</span></span>

<span data-ttu-id="4b19d-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="4b19d-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-430">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="4b19d-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4b19d-431">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="4b19d-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-432">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-432">Type</span></span>

*   [<span data-ttu-id="4b19d-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4b19d-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4b19d-434">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-434">Requirements</span></span>

|<span data-ttu-id="4b19d-435">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-435">Requirement</span></span>| <span data-ttu-id="4b19d-436">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-437">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-438">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-438">1.0</span></span>|
|[<span data-ttu-id="4b19d-439">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-439">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-440">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-441">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-441">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-442">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-442">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-443">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-443">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="4b19d-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-445">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4b19d-446">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4b19d-446">Read mode</span></span>

<span data-ttu-id="4b19d-447">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-447">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="4b19d-448">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4b19d-448">Compose mode</span></span>

<span data-ttu-id="4b19d-449">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4b19d-450">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-450">Type</span></span>

*   <span data-ttu-id="4b19d-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-452">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-452">Requirements</span></span>

|<span data-ttu-id="4b19d-453">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-453">Requirement</span></span>| <span data-ttu-id="4b19d-454">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-455">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-456">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-456">1.0</span></span>|
|[<span data-ttu-id="4b19d-457">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-457">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-458">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-459">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-459">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-460">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-460">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4b19d-461">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="4b19d-461">normalizedSubject: String</span></span>

<span data-ttu-id="4b19d-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4b19d-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="4b19d-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-466">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-466">Type</span></span>

*   <span data-ttu-id="4b19d-467">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-467">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-468">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-468">Requirements</span></span>

|<span data-ttu-id="4b19d-469">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-469">Requirement</span></span>| <span data-ttu-id="4b19d-470">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-470">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-471">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-471">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-472">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-472">1.0</span></span>|
|[<span data-ttu-id="4b19d-473">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-473">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-474">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-474">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-475">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-475">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-476">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-476">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-477">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-477">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="4b19d-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-479">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-479">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-480">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-480">Type</span></span>

*   [<span data-ttu-id="4b19d-481">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4b19d-481">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4b19d-482">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-482">Requirements</span></span>

|<span data-ttu-id="4b19d-483">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-483">Requirement</span></span>| <span data-ttu-id="4b19d-484">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-485">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-486">1.3</span><span class="sxs-lookup"><span data-stu-id="4b19d-486">1.3</span></span>|
|[<span data-ttu-id="4b19d-487">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-488">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-489">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-490">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-490">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-491">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-491">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4b19d-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-493">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="4b19d-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4b19d-494">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4b19d-495">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4b19d-495">Read mode</span></span>

<span data-ttu-id="4b19d-496">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="4b19d-497">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4b19d-498">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-498">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4b19d-499">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4b19d-499">Compose mode</span></span>

<span data-ttu-id="4b19d-500">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-500">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="4b19d-501">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-501">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4b19d-502">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-502">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4b19d-503">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-503">Get 500 members maximum.</span></span>
- <span data-ttu-id="4b19d-504">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-504">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4b19d-505">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-505">Type</span></span>

*   <span data-ttu-id="4b19d-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-507">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-507">Requirements</span></span>

|<span data-ttu-id="4b19d-508">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-508">Requirement</span></span>| <span data-ttu-id="4b19d-509">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-510">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-511">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-511">1.0</span></span>|
|[<span data-ttu-id="4b19d-512">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-513">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-514">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-514">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-515">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-515">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="4b19d-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-p128">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-519">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-519">Type</span></span>

*   [<span data-ttu-id="4b19d-520">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4b19d-520">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4b19d-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-521">Requirements</span></span>

|<span data-ttu-id="4b19d-522">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-522">Requirement</span></span>| <span data-ttu-id="4b19d-523">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-524">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-525">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-525">1.0</span></span>|
|[<span data-ttu-id="4b19d-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-527">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-529">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-529">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-530">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-530">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4b19d-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-532">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="4b19d-532">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4b19d-533">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-533">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4b19d-534">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4b19d-534">Read mode</span></span>

<span data-ttu-id="4b19d-535">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-535">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="4b19d-536">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4b19d-537">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-537">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4b19d-538">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4b19d-538">Compose mode</span></span>

<span data-ttu-id="4b19d-539">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-539">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="4b19d-540">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-540">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4b19d-541">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-541">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4b19d-542">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-542">Get 500 members maximum.</span></span>
- <span data-ttu-id="4b19d-543">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-543">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="4b19d-544">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-544">Type</span></span>

*   <span data-ttu-id="4b19d-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-546">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-546">Requirements</span></span>

|<span data-ttu-id="4b19d-547">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-547">Requirement</span></span>| <span data-ttu-id="4b19d-548">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-549">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-550">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-550">1.0</span></span>|
|[<span data-ttu-id="4b19d-551">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-552">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-553">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-554">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-554">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="4b19d-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-p132">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4b19d-p133">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-560">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-560">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4b19d-561">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-561">Type</span></span>

*   [<span data-ttu-id="4b19d-562">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4b19d-562">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4b19d-563">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-563">Requirements</span></span>

|<span data-ttu-id="4b19d-564">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-564">Requirement</span></span>| <span data-ttu-id="4b19d-565">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-565">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-566">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-566">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-567">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-567">1.0</span></span>|
|[<span data-ttu-id="4b19d-568">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-568">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-569">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-570">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-570">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-571">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-571">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-572">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-572">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="4b19d-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-574">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-574">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4b19d-p134">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="4b19d-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4b19d-577">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4b19d-577">Read mode</span></span>

<span data-ttu-id="4b19d-578">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-578">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="4b19d-579">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4b19d-579">Compose mode</span></span>

<span data-ttu-id="4b19d-580">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-580">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4b19d-581">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="4b19d-581">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4b19d-582">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-582">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4b19d-583">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-583">Type</span></span>

*   <span data-ttu-id="4b19d-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-585">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-585">Requirements</span></span>

|<span data-ttu-id="4b19d-586">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-586">Requirement</span></span>| <span data-ttu-id="4b19d-587">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-588">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-589">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-589">1.0</span></span>|
|[<span data-ttu-id="4b19d-590">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-591">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-592">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-593">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-593">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="4b19d-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-595">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-595">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4b19d-596">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="4b19d-596">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4b19d-597">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4b19d-597">Read mode</span></span>

<span data-ttu-id="4b19d-p135">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="4b19d-600">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4b19d-600">Compose mode</span></span>

<span data-ttu-id="4b19d-601">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="4b19d-601">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="4b19d-602">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-602">Type</span></span>

*   <span data-ttu-id="4b19d-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-604">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-604">Requirements</span></span>

|<span data-ttu-id="4b19d-605">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-605">Requirement</span></span>| <span data-ttu-id="4b19d-606">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-607">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-608">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-608">1.0</span></span>|
|[<span data-ttu-id="4b19d-609">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-610">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-611">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-612">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-612">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4b19d-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4b19d-614">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-614">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4b19d-615">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-615">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4b19d-616">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4b19d-616">Read mode</span></span>

<span data-ttu-id="4b19d-617">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-617">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="4b19d-618">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4b19d-619">Тем не менее, в Windows и Mac вы можете настроить максимальную длину участников 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-619">However, on Windows and Mac, you can set up to get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="4b19d-620">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4b19d-620">Compose mode</span></span>

<span data-ttu-id="4b19d-621">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-621">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="4b19d-622">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-622">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4b19d-623">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-623">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4b19d-624">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-624">Get 500 members maximum.</span></span>
- <span data-ttu-id="4b19d-625">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4b19d-625">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4b19d-626">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-626">Type</span></span>

*   <span data-ttu-id="4b19d-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-628">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-628">Requirements</span></span>

|<span data-ttu-id="4b19d-629">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-629">Requirement</span></span>| <span data-ttu-id="4b19d-630">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-630">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-631">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-631">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-632">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-632">1.0</span></span>|
|[<span data-ttu-id="4b19d-633">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-633">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-634">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-634">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-635">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-635">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-636">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-636">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4b19d-637">Методы</span><span class="sxs-lookup"><span data-stu-id="4b19d-637">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4b19d-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4b19d-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4b19d-639">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-639">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4b19d-640">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-640">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4b19d-641">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4b19d-641">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-642">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-642">Parameters</span></span>

|<span data-ttu-id="4b19d-643">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-643">Name</span></span>| <span data-ttu-id="4b19d-644">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-644">Type</span></span>| <span data-ttu-id="4b19d-645">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b19d-645">Attributes</span></span>| <span data-ttu-id="4b19d-646">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-646">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="4b19d-647">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-647">String</span></span>||<span data-ttu-id="4b19d-p139">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4b19d-650">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-650">String</span></span>||<span data-ttu-id="4b19d-p140">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4b19d-653">Объект</span><span class="sxs-lookup"><span data-stu-id="4b19d-653">Object</span></span>| <span data-ttu-id="4b19d-654">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-654">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-655">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4b19d-655">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="4b19d-656">Объект</span><span class="sxs-lookup"><span data-stu-id="4b19d-656">Object</span></span> | <span data-ttu-id="4b19d-657">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-657">&lt;optional&gt;</span></span> | <span data-ttu-id="4b19d-658">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="4b19d-658">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="4b19d-659">Boolean</span><span class="sxs-lookup"><span data-stu-id="4b19d-659">Boolean</span></span> | <span data-ttu-id="4b19d-660">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-660">&lt;optional&gt;</span></span> | <span data-ttu-id="4b19d-661">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4b19d-661">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="4b19d-662">function</span><span class="sxs-lookup"><span data-stu-id="4b19d-662">function</span></span>| <span data-ttu-id="4b19d-663">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-663">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-664">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b19d-664">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4b19d-665">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-665">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4b19d-666">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="4b19d-666">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4b19d-667">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4b19d-667">Errors</span></span>

| <span data-ttu-id="4b19d-668">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4b19d-668">Error code</span></span> | <span data-ttu-id="4b19d-669">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-669">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="4b19d-670">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="4b19d-670">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="4b19d-671">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="4b19d-671">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4b19d-672">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="4b19d-672">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4b19d-673">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-673">Requirements</span></span>

|<span data-ttu-id="4b19d-674">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-674">Requirement</span></span>| <span data-ttu-id="4b19d-675">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-676">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-677">1.1</span><span class="sxs-lookup"><span data-stu-id="4b19d-677">1.1</span></span>|
|[<span data-ttu-id="4b19d-678">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-679">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-679">ReadWriteItem</span></span>|
|[<span data-ttu-id="4b19d-680">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-681">Создание</span><span class="sxs-lookup"><span data-stu-id="4b19d-681">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4b19d-682">Примеры</span><span class="sxs-lookup"><span data-stu-id="4b19d-682">Examples</span></span>

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

<span data-ttu-id="4b19d-683">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-683">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4b19d-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4b19d-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4b19d-685">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-685">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4b19d-p141">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4b19d-689">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4b19d-689">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4b19d-690">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="4b19d-690">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-691">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-691">Parameters</span></span>

|<span data-ttu-id="4b19d-692">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-692">Name</span></span>| <span data-ttu-id="4b19d-693">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-693">Type</span></span>| <span data-ttu-id="4b19d-694">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b19d-694">Attributes</span></span>| <span data-ttu-id="4b19d-695">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-695">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="4b19d-696">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-696">String</span></span>||<span data-ttu-id="4b19d-p142">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4b19d-699">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-699">String</span></span>||<span data-ttu-id="4b19d-700">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-700">The subject of the item to be attached.</span></span> <span data-ttu-id="4b19d-701">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-701">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4b19d-702">Object</span><span class="sxs-lookup"><span data-stu-id="4b19d-702">Object</span></span>| <span data-ttu-id="4b19d-703">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-703">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-704">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4b19d-704">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4b19d-705">Объект</span><span class="sxs-lookup"><span data-stu-id="4b19d-705">Object</span></span>| <span data-ttu-id="4b19d-706">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-706">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-707">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b19d-707">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4b19d-708">функция</span><span class="sxs-lookup"><span data-stu-id="4b19d-708">function</span></span>| <span data-ttu-id="4b19d-709">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-709">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-710">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b19d-710">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4b19d-711">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-711">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4b19d-712">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="4b19d-712">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4b19d-713">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4b19d-713">Errors</span></span>

| <span data-ttu-id="4b19d-714">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4b19d-714">Error code</span></span> | <span data-ttu-id="4b19d-715">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-715">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4b19d-716">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="4b19d-716">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4b19d-717">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-717">Requirements</span></span>

|<span data-ttu-id="4b19d-718">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-718">Requirement</span></span>| <span data-ttu-id="4b19d-719">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-720">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-721">1.1</span><span class="sxs-lookup"><span data-stu-id="4b19d-721">1.1</span></span>|
|[<span data-ttu-id="4b19d-722">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-722">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-723">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-723">ReadWriteItem</span></span>|
|[<span data-ttu-id="4b19d-724">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-724">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-725">Создание</span><span class="sxs-lookup"><span data-stu-id="4b19d-725">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-726">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-726">Example</span></span>

<span data-ttu-id="4b19d-727">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-727">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="4b19d-728">close()</span><span class="sxs-lookup"><span data-stu-id="4b19d-728">close()</span></span>

<span data-ttu-id="4b19d-729">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="4b19d-729">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4b19d-p144">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-732">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-732">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4b19d-733">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="4b19d-733">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-734">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-734">Requirements</span></span>

|<span data-ttu-id="4b19d-735">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-735">Requirement</span></span>| <span data-ttu-id="4b19d-736">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-736">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-737">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-737">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-738">1.3</span><span class="sxs-lookup"><span data-stu-id="4b19d-738">1.3</span></span>|
|[<span data-ttu-id="4b19d-739">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-739">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-740">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4b19d-740">Restricted</span></span>|
|[<span data-ttu-id="4b19d-741">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-741">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-742">Создание</span><span class="sxs-lookup"><span data-stu-id="4b19d-742">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="4b19d-743">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4b19d-743">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="4b19d-744">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-744">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-745">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-745">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4b19d-746">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="4b19d-746">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4b19d-747">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="4b19d-747">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4b19d-p145">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-751">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-751">Parameters</span></span>

| <span data-ttu-id="4b19d-752">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-752">Name</span></span> | <span data-ttu-id="4b19d-753">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-753">Type</span></span> | <span data-ttu-id="4b19d-754">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b19d-754">Attributes</span></span> | <span data-ttu-id="4b19d-755">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-755">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="4b19d-756">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4b19d-756">String &#124; Object</span></span>| |<span data-ttu-id="4b19d-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4b19d-759">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="4b19d-759">**OR**</span></span><br/><span data-ttu-id="4b19d-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4b19d-762">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-762">String</span></span> | <span data-ttu-id="4b19d-763">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-763">&lt;optional&gt;</span></span> | <span data-ttu-id="4b19d-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4b19d-766">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-766">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4b19d-767">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-767">&lt;optional&gt;</span></span> | <span data-ttu-id="4b19d-768">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="4b19d-768">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4b19d-769">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-769">String</span></span> | | <span data-ttu-id="4b19d-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4b19d-772">Строка</span><span class="sxs-lookup"><span data-stu-id="4b19d-772">String</span></span> | | <span data-ttu-id="4b19d-773">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-773">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4b19d-774">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-774">String</span></span> | | <span data-ttu-id="4b19d-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="4b19d-777">Логический</span><span class="sxs-lookup"><span data-stu-id="4b19d-777">Boolean</span></span> | | <span data-ttu-id="4b19d-p151">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4b19d-780">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-780">String</span></span> | | <span data-ttu-id="4b19d-p152">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4b19d-784">function</span><span class="sxs-lookup"><span data-stu-id="4b19d-784">function</span></span> | <span data-ttu-id="4b19d-785">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-785">&lt;optional&gt;</span></span> | <span data-ttu-id="4b19d-786">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b19d-786">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4b19d-787">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-787">Requirements</span></span>

|<span data-ttu-id="4b19d-788">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-788">Requirement</span></span>| <span data-ttu-id="4b19d-789">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-789">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-790">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-790">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-791">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-791">1.0</span></span>|
|[<span data-ttu-id="4b19d-792">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-792">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-793">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-793">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-794">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-794">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-795">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-795">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4b19d-796">Примеры</span><span class="sxs-lookup"><span data-stu-id="4b19d-796">Examples</span></span>

<span data-ttu-id="4b19d-797">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-797">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4b19d-798">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-798">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4b19d-799">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-799">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4b19d-800">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="4b19d-800">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4b19d-801">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="4b19d-801">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4b19d-802">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="4b19d-802">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="4b19d-803">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4b19d-803">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="4b19d-804">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-804">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-805">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-805">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4b19d-806">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="4b19d-806">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4b19d-807">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="4b19d-807">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4b19d-p153">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-811">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-811">Parameters</span></span>

| <span data-ttu-id="4b19d-812">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-812">Name</span></span> | <span data-ttu-id="4b19d-813">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-813">Type</span></span> | <span data-ttu-id="4b19d-814">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b19d-814">Attributes</span></span> | <span data-ttu-id="4b19d-815">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-815">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="4b19d-816">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4b19d-816">String &#124; Object</span></span>| | <span data-ttu-id="4b19d-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4b19d-819">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="4b19d-819">**OR**</span></span><br/><span data-ttu-id="4b19d-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4b19d-822">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-822">String</span></span> | <span data-ttu-id="4b19d-823">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-823">&lt;optional&gt;</span></span> | <span data-ttu-id="4b19d-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4b19d-826">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-826">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4b19d-827">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-827">&lt;optional&gt;</span></span> | <span data-ttu-id="4b19d-828">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="4b19d-828">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4b19d-829">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-829">String</span></span> | | <span data-ttu-id="4b19d-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4b19d-832">Строка</span><span class="sxs-lookup"><span data-stu-id="4b19d-832">String</span></span> | | <span data-ttu-id="4b19d-833">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-833">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4b19d-834">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-834">String</span></span> | | <span data-ttu-id="4b19d-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="4b19d-837">Логический</span><span class="sxs-lookup"><span data-stu-id="4b19d-837">Boolean</span></span> | | <span data-ttu-id="4b19d-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4b19d-840">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-840">String</span></span> | | <span data-ttu-id="4b19d-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4b19d-844">function</span><span class="sxs-lookup"><span data-stu-id="4b19d-844">function</span></span> | <span data-ttu-id="4b19d-845">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-845">&lt;optional&gt;</span></span> | <span data-ttu-id="4b19d-846">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b19d-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4b19d-847">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-847">Requirements</span></span>

|<span data-ttu-id="4b19d-848">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-848">Requirement</span></span>| <span data-ttu-id="4b19d-849">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-850">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-851">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-851">1.0</span></span>|
|[<span data-ttu-id="4b19d-852">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-853">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-854">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-855">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-855">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4b19d-856">Примеры</span><span class="sxs-lookup"><span data-stu-id="4b19d-856">Examples</span></span>

<span data-ttu-id="4b19d-857">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-857">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4b19d-858">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-858">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4b19d-859">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-859">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4b19d-860">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="4b19d-860">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4b19d-861">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="4b19d-861">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4b19d-862">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="4b19d-862">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="4b19d-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="4b19d-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="4b19d-864">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-864">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-865">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-866">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-866">Requirements</span></span>

|<span data-ttu-id="4b19d-867">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-867">Requirement</span></span>| <span data-ttu-id="4b19d-868">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-868">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-869">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-869">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-870">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-870">1.0</span></span>|
|[<span data-ttu-id="4b19d-871">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-871">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-872">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-872">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-873">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-873">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-874">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-874">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b19d-875">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4b19d-875">Returns:</span></span>

<span data-ttu-id="4b19d-876">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-876">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="4b19d-877">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-877">Example</span></span>

<span data-ttu-id="4b19d-878">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-878">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="4b19d-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="4b19d-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="4b19d-880">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-880">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-881">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-881">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-882">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-882">Parameters</span></span>

|<span data-ttu-id="4b19d-883">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-883">Name</span></span>| <span data-ttu-id="4b19d-884">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-884">Type</span></span>| <span data-ttu-id="4b19d-885">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-885">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="4b19d-886">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4b19d-886">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="4b19d-887">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="4b19d-887">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b19d-888">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-888">Requirements</span></span>

|<span data-ttu-id="4b19d-889">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-889">Requirement</span></span>| <span data-ttu-id="4b19d-890">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-891">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-892">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-892">1.0</span></span>|
|[<span data-ttu-id="4b19d-893">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-894">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4b19d-894">Restricted</span></span>|
|[<span data-ttu-id="4b19d-895">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-896">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-896">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b19d-897">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4b19d-897">Returns:</span></span>

<span data-ttu-id="4b19d-898">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="4b19d-898">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4b19d-899">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="4b19d-899">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4b19d-900">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-900">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4b19d-901">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="4b19d-901">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="4b19d-902">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="4b19d-902">Value of `entityType`</span></span> | <span data-ttu-id="4b19d-903">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="4b19d-903">Type of objects in returned array</span></span> | <span data-ttu-id="4b19d-904">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-904">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="4b19d-905">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-905">String</span></span> | <span data-ttu-id="4b19d-906">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4b19d-906">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="4b19d-907">Contact</span><span class="sxs-lookup"><span data-stu-id="4b19d-907">Contact</span></span> | <span data-ttu-id="4b19d-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4b19d-908">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="4b19d-909">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-909">String</span></span> | <span data-ttu-id="4b19d-910">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4b19d-910">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="4b19d-911">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4b19d-911">MeetingSuggestion</span></span> | <span data-ttu-id="4b19d-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4b19d-912">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="4b19d-913">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4b19d-913">PhoneNumber</span></span> | <span data-ttu-id="4b19d-914">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4b19d-914">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="4b19d-915">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4b19d-915">TaskSuggestion</span></span> | <span data-ttu-id="4b19d-916">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4b19d-916">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="4b19d-917">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-917">String</span></span> | <span data-ttu-id="4b19d-918">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4b19d-918">**Restricted**</span></span> |

<span data-ttu-id="4b19d-919">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="4b19d-919">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="4b19d-920">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-920">Example</span></span>

<span data-ttu-id="4b19d-921">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-921">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="4b19d-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="4b19d-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="4b19d-923">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4b19d-923">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-924">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-924">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4b19d-925">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-925">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-926">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-926">Parameters</span></span>

|<span data-ttu-id="4b19d-927">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-927">Name</span></span>| <span data-ttu-id="4b19d-928">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-928">Type</span></span>| <span data-ttu-id="4b19d-929">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-929">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4b19d-930">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-930">String</span></span>|<span data-ttu-id="4b19d-931">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="4b19d-931">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b19d-932">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-932">Requirements</span></span>

|<span data-ttu-id="4b19d-933">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-933">Requirement</span></span>| <span data-ttu-id="4b19d-934">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-935">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-936">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-936">1.0</span></span>|
|[<span data-ttu-id="4b19d-937">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-938">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-939">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-940">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b19d-941">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4b19d-941">Returns:</span></span>

<span data-ttu-id="4b19d-p162">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4b19d-944">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="4b19d-944">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="4b19d-945">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4b19d-945">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4b19d-946">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4b19d-946">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-947">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4b19d-p163">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4b19d-951">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-951">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4b19d-952">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-952">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4b19d-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-956">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-956">Requirements</span></span>

|<span data-ttu-id="4b19d-957">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-957">Requirement</span></span>| <span data-ttu-id="4b19d-958">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-958">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-959">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-959">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-960">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-960">1.0</span></span>|
|[<span data-ttu-id="4b19d-961">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-961">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-962">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-962">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-963">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-963">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-964">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-964">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b19d-965">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4b19d-965">Returns:</span></span>

<span data-ttu-id="4b19d-p165">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="4b19d-968">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="4b19d-968">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="4b19d-969">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-969">Example</span></span>

<span data-ttu-id="4b19d-970">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="4b19d-970">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4b19d-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="4b19d-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4b19d-972">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4b19d-972">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-973">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-973">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4b19d-974">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-974">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4b19d-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-977">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-977">Parameters</span></span>

|<span data-ttu-id="4b19d-978">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-978">Name</span></span>| <span data-ttu-id="4b19d-979">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-979">Type</span></span>| <span data-ttu-id="4b19d-980">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-980">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4b19d-981">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-981">String</span></span>|<span data-ttu-id="4b19d-982">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="4b19d-982">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b19d-983">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-983">Requirements</span></span>

|<span data-ttu-id="4b19d-984">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-984">Requirement</span></span>| <span data-ttu-id="4b19d-985">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-986">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-987">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-987">1.0</span></span>|
|[<span data-ttu-id="4b19d-988">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-988">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-989">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-990">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-990">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-991">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-991">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b19d-992">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4b19d-992">Returns:</span></span>

<span data-ttu-id="4b19d-993">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4b19d-993">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="4b19d-994">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4b19d-994">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="4b19d-995">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-995">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4b19d-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4b19d-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4b19d-997">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-997">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4b19d-p167">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает пустую строку для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p167">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-1000">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-1000">Parameters</span></span>

|<span data-ttu-id="4b19d-1001">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-1001">Name</span></span>| <span data-ttu-id="4b19d-1002">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-1002">Type</span></span>| <span data-ttu-id="4b19d-1003">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b19d-1003">Attributes</span></span>| <span data-ttu-id="4b19d-1004">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1004">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="4b19d-1005">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4b19d-1005">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4b19d-p168">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="4b19d-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="4b19d-1009">Object</span><span class="sxs-lookup"><span data-stu-id="4b19d-1009">Object</span></span>| <span data-ttu-id="4b19d-1010">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1011">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1011">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4b19d-1012">Объект</span><span class="sxs-lookup"><span data-stu-id="4b19d-1012">Object</span></span>| <span data-ttu-id="4b19d-1013">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1014">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1014">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4b19d-1015">функция</span><span class="sxs-lookup"><span data-stu-id="4b19d-1015">function</span></span>||<span data-ttu-id="4b19d-1016">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b19d-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4b19d-1017">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1017">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4b19d-1018">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1018">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b19d-1019">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-1019">Requirements</span></span>

|<span data-ttu-id="4b19d-1020">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-1020">Requirement</span></span>| <span data-ttu-id="4b19d-1021">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1021">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-1022">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-1022">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-1023">1.2</span><span class="sxs-lookup"><span data-stu-id="4b19d-1023">1.2</span></span>|
|[<span data-ttu-id="4b19d-1024">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-1024">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-1025">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-1025">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-1026">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-1026">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-1027">Создание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1027">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b19d-1028">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4b19d-1028">Returns:</span></span>

<span data-ttu-id="4b19d-1029">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1029">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="4b19d-1030">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="4b19d-1030">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4b19d-1031">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-1031">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="4b19d-1032">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="4b19d-1032">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="4b19d-1033">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1033">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="4b19d-1034">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4b19d-1034">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-1035">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1035">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-1036">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-1036">Requirements</span></span>

|<span data-ttu-id="4b19d-1037">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-1037">Requirement</span></span>| <span data-ttu-id="4b19d-1038">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1038">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-1039">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-1039">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-1040">1.6</span><span class="sxs-lookup"><span data-stu-id="4b19d-1040">1.6</span></span> |
|[<span data-ttu-id="4b19d-1041">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-1041">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-1042">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-1042">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-1043">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-1043">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-1044">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1044">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b19d-1045">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4b19d-1045">Returns:</span></span>

<span data-ttu-id="4b19d-1046">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4b19d-1046">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="4b19d-1047">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-1047">Example</span></span>

<span data-ttu-id="4b19d-1048">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1048">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="4b19d-1049">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4b19d-1049">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="4b19d-p171">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4b19d-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-1052">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1052">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4b19d-p172">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4b19d-1056">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1056">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4b19d-1057">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1057">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4b19d-p173">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4b19d-1061">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-1061">Requirements</span></span>

|<span data-ttu-id="4b19d-1062">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-1062">Requirement</span></span>| <span data-ttu-id="4b19d-1063">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1063">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-1064">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-1064">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-1065">1.6</span><span class="sxs-lookup"><span data-stu-id="4b19d-1065">1.6</span></span> |
|[<span data-ttu-id="4b19d-1066">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-1066">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-1067">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-1067">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-1068">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-1068">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-1069">Чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1069">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4b19d-1070">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4b19d-1070">Returns:</span></span>

<span data-ttu-id="4b19d-p174">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="4b19d-1073">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-1073">Example</span></span>

<span data-ttu-id="4b19d-1074">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1074">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4b19d-1075">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4b19d-1075">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4b19d-1076">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1076">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4b19d-p175">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p175">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-1080">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-1080">Parameters</span></span>

|<span data-ttu-id="4b19d-1081">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-1081">Name</span></span>| <span data-ttu-id="4b19d-1082">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-1082">Type</span></span>| <span data-ttu-id="4b19d-1083">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b19d-1083">Attributes</span></span>| <span data-ttu-id="4b19d-1084">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1084">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4b19d-1085">function</span><span class="sxs-lookup"><span data-stu-id="4b19d-1085">function</span></span>||<span data-ttu-id="4b19d-1086">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b19d-1086">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4b19d-1087">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1087">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4b19d-1088">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1088">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="4b19d-1089">Объект</span><span class="sxs-lookup"><span data-stu-id="4b19d-1089">Object</span></span>| <span data-ttu-id="4b19d-1090">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1091">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1091">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4b19d-1092">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1092">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b19d-1093">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-1093">Requirements</span></span>

|<span data-ttu-id="4b19d-1094">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-1094">Requirement</span></span>| <span data-ttu-id="4b19d-1095">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-1096">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-1097">1.0</span><span class="sxs-lookup"><span data-stu-id="4b19d-1097">1.0</span></span>|
|[<span data-ttu-id="4b19d-1098">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-1099">ReadItem</span></span>|
|[<span data-ttu-id="4b19d-1100">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-1101">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1101">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-1102">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-1102">Example</span></span>

<span data-ttu-id="4b19d-p178">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p178">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4b19d-1106">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4b19d-1106">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4b19d-1107">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1107">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4b19d-1108">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1108">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="4b19d-1109">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1109">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="4b19d-1110">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1110">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="4b19d-1111">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1111">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-1112">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-1112">Parameters</span></span>

|<span data-ttu-id="4b19d-1113">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-1113">Name</span></span>| <span data-ttu-id="4b19d-1114">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-1114">Type</span></span>| <span data-ttu-id="4b19d-1115">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b19d-1115">Attributes</span></span>| <span data-ttu-id="4b19d-1116">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1116">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="4b19d-1117">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-1117">String</span></span>||<span data-ttu-id="4b19d-1118">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1118">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="4b19d-1119">Object</span><span class="sxs-lookup"><span data-stu-id="4b19d-1119">Object</span></span>| <span data-ttu-id="4b19d-1120">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1121">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4b19d-1122">Object</span><span class="sxs-lookup"><span data-stu-id="4b19d-1122">Object</span></span>| <span data-ttu-id="4b19d-1123">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1124">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4b19d-1125">функция</span><span class="sxs-lookup"><span data-stu-id="4b19d-1125">function</span></span>| <span data-ttu-id="4b19d-1126">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1126">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1127">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b19d-1127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4b19d-1128">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1128">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4b19d-1129">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4b19d-1129">Errors</span></span>

| <span data-ttu-id="4b19d-1130">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4b19d-1130">Error code</span></span> | <span data-ttu-id="4b19d-1131">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1131">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="4b19d-1132">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1132">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4b19d-1133">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-1133">Requirements</span></span>

|<span data-ttu-id="4b19d-1134">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-1134">Requirement</span></span>| <span data-ttu-id="4b19d-1135">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-1136">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4b19d-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-1137">1.1</span><span class="sxs-lookup"><span data-stu-id="4b19d-1137">1.1</span></span>|
|[<span data-ttu-id="4b19d-1138">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-1139">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-1139">ReadWriteItem</span></span>|
|[<span data-ttu-id="4b19d-1140">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-1141">Создание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1141">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-1142">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-1142">Example</span></span>

<span data-ttu-id="4b19d-1143">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="4b19d-1143">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="4b19d-1144">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4b19d-1144">saveAsync([options], callback)</span></span>

<span data-ttu-id="4b19d-1145">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1145">Asynchronously saves an item.</span></span>

<span data-ttu-id="4b19d-1146">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1146">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="4b19d-1147">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1147">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="4b19d-1148">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1148">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-1149">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1149">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4b19d-1150">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1150">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4b19d-p182">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p182">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4b19d-1154">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="4b19d-1154">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4b19d-1155">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1155">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="4b19d-1156">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1156">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="4b19d-1157">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="4b19d-1157">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="4b19d-1158">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1158">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-1159">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-1159">Parameters</span></span>

|<span data-ttu-id="4b19d-1160">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-1160">Name</span></span>| <span data-ttu-id="4b19d-1161">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-1161">Type</span></span>| <span data-ttu-id="4b19d-1162">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b19d-1162">Attributes</span></span>| <span data-ttu-id="4b19d-1163">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1163">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="4b19d-1164">Объект</span><span class="sxs-lookup"><span data-stu-id="4b19d-1164">Object</span></span>| <span data-ttu-id="4b19d-1165">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1165">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1166">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1166">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4b19d-1167">Объект</span><span class="sxs-lookup"><span data-stu-id="4b19d-1167">Object</span></span>| <span data-ttu-id="4b19d-1168">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1168">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1169">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1169">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4b19d-1170">функция</span><span class="sxs-lookup"><span data-stu-id="4b19d-1170">function</span></span>||<span data-ttu-id="4b19d-1171">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b19d-1171">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4b19d-1172">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1172">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4b19d-1173">Requirements</span><span class="sxs-lookup"><span data-stu-id="4b19d-1173">Requirements</span></span>

|<span data-ttu-id="4b19d-1174">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-1174">Requirement</span></span>| <span data-ttu-id="4b19d-1175">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1175">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-1176">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-1176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-1177">1.3</span><span class="sxs-lookup"><span data-stu-id="4b19d-1177">1.3</span></span>|
|[<span data-ttu-id="4b19d-1178">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-1178">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-1179">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-1179">ReadWriteItem</span></span>|
|[<span data-ttu-id="4b19d-1180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-1180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-1181">Создание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1181">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4b19d-1182">Примеры</span><span class="sxs-lookup"><span data-stu-id="4b19d-1182">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="4b19d-p184">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p184">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4b19d-1185">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4b19d-1185">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4b19d-1186">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1186">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4b19d-p185">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p185">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4b19d-1190">Параметры</span><span class="sxs-lookup"><span data-stu-id="4b19d-1190">Parameters</span></span>

|<span data-ttu-id="4b19d-1191">Имя</span><span class="sxs-lookup"><span data-stu-id="4b19d-1191">Name</span></span>| <span data-ttu-id="4b19d-1192">Тип</span><span class="sxs-lookup"><span data-stu-id="4b19d-1192">Type</span></span>| <span data-ttu-id="4b19d-1193">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4b19d-1193">Attributes</span></span>| <span data-ttu-id="4b19d-1194">Описание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1194">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4b19d-1195">String</span><span class="sxs-lookup"><span data-stu-id="4b19d-1195">String</span></span>||<span data-ttu-id="4b19d-p186">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-p186">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="4b19d-1199">Object</span><span class="sxs-lookup"><span data-stu-id="4b19d-1199">Object</span></span>| <span data-ttu-id="4b19d-1200">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1201">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1201">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4b19d-1202">Объект</span><span class="sxs-lookup"><span data-stu-id="4b19d-1202">Object</span></span>| <span data-ttu-id="4b19d-1203">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1203">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1204">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1204">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="4b19d-1205">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4b19d-1205">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="4b19d-1206">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4b19d-1206">&lt;optional&gt;</span></span>|<span data-ttu-id="4b19d-1207">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1207">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="4b19d-1208">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1208">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4b19d-1209">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1209">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="4b19d-1210">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1210">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4b19d-1211">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="4b19d-1211">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="4b19d-1212">функция</span><span class="sxs-lookup"><span data-stu-id="4b19d-1212">function</span></span>||<span data-ttu-id="4b19d-1213">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4b19d-1213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4b19d-1214">Требования</span><span class="sxs-lookup"><span data-stu-id="4b19d-1214">Requirements</span></span>

|<span data-ttu-id="4b19d-1215">Требование</span><span class="sxs-lookup"><span data-stu-id="4b19d-1215">Requirement</span></span>| <span data-ttu-id="4b19d-1216">Значение</span><span class="sxs-lookup"><span data-stu-id="4b19d-1216">Value</span></span>|
|---|---|
|[<span data-ttu-id="4b19d-1217">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4b19d-1217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4b19d-1218">1.2</span><span class="sxs-lookup"><span data-stu-id="4b19d-1218">1.2</span></span>|
|[<span data-ttu-id="4b19d-1219">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4b19d-1219">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4b19d-1220">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4b19d-1220">ReadWriteItem</span></span>|
|[<span data-ttu-id="4b19d-1221">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4b19d-1221">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4b19d-1222">Создание</span><span class="sxs-lookup"><span data-stu-id="4b19d-1222">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4b19d-1223">Пример</span><span class="sxs-lookup"><span data-stu-id="4b19d-1223">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
