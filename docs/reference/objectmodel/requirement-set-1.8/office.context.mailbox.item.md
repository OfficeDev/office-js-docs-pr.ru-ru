---
title: Office. Context. Mailbox. Item — набор требований 1,8
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 065ea3c74580555c0df1af7b495127a25493b612
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001574"
---
# <a name="item"></a><span data-ttu-id="e7168-102">item</span><span class="sxs-lookup"><span data-stu-id="e7168-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="e7168-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="e7168-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="e7168-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="e7168-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-106">Requirements</span></span>

|<span data-ttu-id="e7168-107">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-107">Requirement</span></span>|<span data-ttu-id="e7168-108">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-110">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-110">1.0</span></span>|
|[<span data-ttu-id="e7168-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e7168-112">Restricted</span></span>|
|[<span data-ttu-id="e7168-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e7168-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="e7168-115">Members and methods</span></span>

| <span data-ttu-id="e7168-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-116">Member</span></span> | <span data-ttu-id="e7168-117">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e7168-118">attachments</span><span class="sxs-lookup"><span data-stu-id="e7168-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="e7168-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-119">Member</span></span> |
| [<span data-ttu-id="e7168-120">bcc</span><span class="sxs-lookup"><span data-stu-id="e7168-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="e7168-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-121">Member</span></span> |
| [<span data-ttu-id="e7168-122">body</span><span class="sxs-lookup"><span data-stu-id="e7168-122">body</span></span>](#body-body) | <span data-ttu-id="e7168-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-123">Member</span></span> |
| [<span data-ttu-id="e7168-124">разделов</span><span class="sxs-lookup"><span data-stu-id="e7168-124">categories</span></span>](#categories-categories) | <span data-ttu-id="e7168-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-125">Member</span></span> |
| [<span data-ttu-id="e7168-126">cc</span><span class="sxs-lookup"><span data-stu-id="e7168-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e7168-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-127">Member</span></span> |
| [<span data-ttu-id="e7168-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="e7168-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="e7168-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-129">Member</span></span> |
| [<span data-ttu-id="e7168-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="e7168-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="e7168-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-131">Member</span></span> |
| [<span data-ttu-id="e7168-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="e7168-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="e7168-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-133">Member</span></span> |
| [<span data-ttu-id="e7168-134">end</span><span class="sxs-lookup"><span data-stu-id="e7168-134">end</span></span>](#end-datetime) | <span data-ttu-id="e7168-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-135">Member</span></span> |
| [<span data-ttu-id="e7168-136">енханцедлокатион</span><span class="sxs-lookup"><span data-stu-id="e7168-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="e7168-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-137">Member</span></span> |
| [<span data-ttu-id="e7168-138">from</span><span class="sxs-lookup"><span data-stu-id="e7168-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="e7168-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-139">Member</span></span> |
| [<span data-ttu-id="e7168-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="e7168-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="e7168-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-141">Member</span></span> |
| [<span data-ttu-id="e7168-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="e7168-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="e7168-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-143">Member</span></span> |
| [<span data-ttu-id="e7168-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="e7168-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="e7168-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-145">Member</span></span> |
| [<span data-ttu-id="e7168-146">itemId</span><span class="sxs-lookup"><span data-stu-id="e7168-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="e7168-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-147">Member</span></span> |
| [<span data-ttu-id="e7168-148">itemType</span><span class="sxs-lookup"><span data-stu-id="e7168-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="e7168-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-149">Member</span></span> |
| [<span data-ttu-id="e7168-150">location</span><span class="sxs-lookup"><span data-stu-id="e7168-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="e7168-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-151">Member</span></span> |
| [<span data-ttu-id="e7168-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="e7168-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="e7168-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-153">Member</span></span> |
| [<span data-ttu-id="e7168-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="e7168-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="e7168-155">Member</span><span class="sxs-lookup"><span data-stu-id="e7168-155">Member</span></span> |
| [<span data-ttu-id="e7168-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="e7168-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e7168-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-157">Member</span></span> |
| [<span data-ttu-id="e7168-158">organizer</span><span class="sxs-lookup"><span data-stu-id="e7168-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="e7168-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-159">Member</span></span> |
| [<span data-ttu-id="e7168-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="e7168-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="e7168-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-161">Member</span></span> |
| [<span data-ttu-id="e7168-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="e7168-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e7168-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-163">Member</span></span> |
| [<span data-ttu-id="e7168-164">sender</span><span class="sxs-lookup"><span data-stu-id="e7168-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="e7168-165">Member</span><span class="sxs-lookup"><span data-stu-id="e7168-165">Member</span></span> |
| [<span data-ttu-id="e7168-166">seriesId</span><span class="sxs-lookup"><span data-stu-id="e7168-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="e7168-167">Member</span><span class="sxs-lookup"><span data-stu-id="e7168-167">Member</span></span> |
| [<span data-ttu-id="e7168-168">start</span><span class="sxs-lookup"><span data-stu-id="e7168-168">start</span></span>](#start-datetime) | <span data-ttu-id="e7168-169">Member</span><span class="sxs-lookup"><span data-stu-id="e7168-169">Member</span></span> |
| [<span data-ttu-id="e7168-170">subject</span><span class="sxs-lookup"><span data-stu-id="e7168-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="e7168-171">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-171">Member</span></span> |
| [<span data-ttu-id="e7168-172">to</span><span class="sxs-lookup"><span data-stu-id="e7168-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e7168-173">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7168-173">Member</span></span> |
| [<span data-ttu-id="e7168-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e7168-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="e7168-175">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-175">Method</span></span> |
| [<span data-ttu-id="e7168-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="e7168-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="e7168-177">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-177">Method</span></span> |
| [<span data-ttu-id="e7168-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e7168-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e7168-179">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-179">Method</span></span> |
| [<span data-ttu-id="e7168-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e7168-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="e7168-181">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-181">Method</span></span> |
| [<span data-ttu-id="e7168-182">close</span><span class="sxs-lookup"><span data-stu-id="e7168-182">close</span></span>](#close) | <span data-ttu-id="e7168-183">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-183">Method</span></span> |
| [<span data-ttu-id="e7168-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="e7168-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="e7168-185">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-185">Method</span></span> |
| [<span data-ttu-id="e7168-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="e7168-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="e7168-187">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-187">Method</span></span> |
| [<span data-ttu-id="e7168-188">жеталлинтернесеадерсасинк</span><span class="sxs-lookup"><span data-stu-id="e7168-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="e7168-189">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-189">Method</span></span> |
| [<span data-ttu-id="e7168-190">жетаттачментконтентасинк</span><span class="sxs-lookup"><span data-stu-id="e7168-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="e7168-191">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-191">Method</span></span> |
| [<span data-ttu-id="e7168-192">жетаттачментсасинк</span><span class="sxs-lookup"><span data-stu-id="e7168-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="e7168-193">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-193">Method</span></span> |
| [<span data-ttu-id="e7168-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="e7168-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="e7168-195">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-195">Method</span></span> |
| [<span data-ttu-id="e7168-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="e7168-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="e7168-197">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-197">Method</span></span> |
| [<span data-ttu-id="e7168-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="e7168-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="e7168-199">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-199">Method</span></span> |
| [<span data-ttu-id="e7168-200">жетитемидасинк</span><span class="sxs-lookup"><span data-stu-id="e7168-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="e7168-201">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-201">Method</span></span> |
| [<span data-ttu-id="e7168-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e7168-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="e7168-203">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-203">Method</span></span> |
| [<span data-ttu-id="e7168-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="e7168-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="e7168-205">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-205">Method</span></span> |
| [<span data-ttu-id="e7168-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e7168-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="e7168-207">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-207">Method</span></span> |
| [<span data-ttu-id="e7168-208">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="e7168-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="e7168-209">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-209">Method</span></span> |
| [<span data-ttu-id="e7168-210">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="e7168-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="e7168-211">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-211">Method</span></span> |
| [<span data-ttu-id="e7168-212">жетшаредпропертиесасинк</span><span class="sxs-lookup"><span data-stu-id="e7168-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="e7168-213">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-213">Method</span></span> |
| [<span data-ttu-id="e7168-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e7168-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="e7168-215">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-215">Method</span></span> |
| [<span data-ttu-id="e7168-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e7168-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="e7168-217">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-217">Method</span></span> |
| [<span data-ttu-id="e7168-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e7168-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="e7168-219">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-219">Method</span></span> |
| [<span data-ttu-id="e7168-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="e7168-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="e7168-221">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-221">Method</span></span> |
| [<span data-ttu-id="e7168-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e7168-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="e7168-223">Метод</span><span class="sxs-lookup"><span data-stu-id="e7168-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="e7168-224">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-224">Example</span></span>

<span data-ttu-id="e7168-225">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="e7168-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="e7168-226">Members</span><span class="sxs-lookup"><span data-stu-id="e7168-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="e7168-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="e7168-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="e7168-228">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="e7168-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="e7168-229">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e7168-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-230">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="e7168-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="e7168-231">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="e7168-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-232">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-232">Type</span></span>

*   <span data-ttu-id="e7168-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="e7168-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-234">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-234">Requirements</span></span>

|<span data-ttu-id="e7168-235">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-235">Requirement</span></span>|<span data-ttu-id="e7168-236">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-238">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-238">1.0</span></span>|
|[<span data-ttu-id="e7168-239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-240">ReadItem</span></span>|
|[<span data-ttu-id="e7168-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-242">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-243">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-243">Example</span></span>

<span data-ttu-id="e7168-244">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="e7168-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-246">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="e7168-247">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e7168-247">Compose mode only.</span></span>

<span data-ttu-id="e7168-248">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-248">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e7168-249">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="e7168-249">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e7168-250">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-250">Get 500 members maximum.</span></span>
- <span data-ttu-id="e7168-251">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-251">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-252">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-252">Type</span></span>

*   [<span data-ttu-id="e7168-253">Получатели</span><span class="sxs-lookup"><span data-stu-id="e7168-253">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="e7168-254">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-254">Requirements</span></span>

|<span data-ttu-id="e7168-255">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-255">Requirement</span></span>|<span data-ttu-id="e7168-256">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-256">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-257">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-257">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-258">1.1</span><span class="sxs-lookup"><span data-stu-id="e7168-258">1.1</span></span>|
|[<span data-ttu-id="e7168-259">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-259">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-260">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-260">ReadItem</span></span>|
|[<span data-ttu-id="e7168-261">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-261">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-262">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-262">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-263">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-263">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-18"></a><span data-ttu-id="e7168-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-265">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-265">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-266">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-266">Type</span></span>

*   [<span data-ttu-id="e7168-267">Body</span><span class="sxs-lookup"><span data-stu-id="e7168-267">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="e7168-268">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-268">Requirements</span></span>

|<span data-ttu-id="e7168-269">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-269">Requirement</span></span>|<span data-ttu-id="e7168-270">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-271">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-272">1.1</span><span class="sxs-lookup"><span data-stu-id="e7168-272">1.1</span></span>|
|[<span data-ttu-id="e7168-273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-274">ReadItem</span></span>|
|[<span data-ttu-id="e7168-275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-277">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-277">Example</span></span>

<span data-ttu-id="e7168-278">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="e7168-278">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="e7168-279">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-279">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategoriesviewoutlook-js-18"></a><span data-ttu-id="e7168-280">Категории: [категории](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-280">categories: [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-281">Получает объект, предоставляющий методы для управления категориями элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-281">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-282">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-282">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-283">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-283">Type</span></span>

*   [<span data-ttu-id="e7168-284">Categories</span><span class="sxs-lookup"><span data-stu-id="e7168-284">Categories</span></span>](/javascript/api/outlook/office.categories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="e7168-285">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-285">Requirements</span></span>

|<span data-ttu-id="e7168-286">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-286">Requirement</span></span>|<span data-ttu-id="e7168-287">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-288">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-289">1.8</span><span class="sxs-lookup"><span data-stu-id="e7168-289">1.8</span></span>|
|[<span data-ttu-id="e7168-290">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-291">ReadItem</span></span>|
|[<span data-ttu-id="e7168-292">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-293">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-294">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-294">Example</span></span>

<span data-ttu-id="e7168-295">В этом примере возвращаются категории элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-295">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="e7168-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-297">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-297">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="e7168-298">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-298">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-299">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-299">Read mode</span></span>

<span data-ttu-id="e7168-300">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-300">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="e7168-301">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-301">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e7168-302">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="e7168-302">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-303">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-303">Compose mode</span></span>

<span data-ttu-id="e7168-304">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-304">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="e7168-305">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-305">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e7168-306">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="e7168-306">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e7168-307">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-307">Get 500 members maximum.</span></span>
- <span data-ttu-id="e7168-308">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-308">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e7168-309">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-309">Type</span></span>

*   <span data-ttu-id="e7168-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-311">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-311">Requirements</span></span>

|<span data-ttu-id="e7168-312">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-312">Requirement</span></span>|<span data-ttu-id="e7168-313">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-314">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-315">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-315">1.0</span></span>|
|[<span data-ttu-id="e7168-316">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-316">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-317">ReadItem</span></span>|
|[<span data-ttu-id="e7168-318">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-318">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-319">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-319">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="e7168-320">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="e7168-320">(nullable) conversationId: String</span></span>

<span data-ttu-id="e7168-321">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="e7168-321">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="e7168-p109">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="e7168-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="e7168-p110">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="e7168-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-326">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-326">Type</span></span>

*   <span data-ttu-id="e7168-327">String</span><span class="sxs-lookup"><span data-stu-id="e7168-327">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-328">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-328">Requirements</span></span>

|<span data-ttu-id="e7168-329">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-329">Requirement</span></span>|<span data-ttu-id="e7168-330">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-330">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-331">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-331">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-332">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-332">1.0</span></span>|
|[<span data-ttu-id="e7168-333">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-333">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-334">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-334">ReadItem</span></span>|
|[<span data-ttu-id="e7168-335">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-335">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-336">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-336">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-337">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-337">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="e7168-338">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="e7168-338">dateTimeCreated: Date</span></span>

<span data-ttu-id="e7168-p111">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e7168-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-341">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-341">Type</span></span>

*   <span data-ttu-id="e7168-342">Дата</span><span class="sxs-lookup"><span data-stu-id="e7168-342">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-343">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-343">Requirements</span></span>

|<span data-ttu-id="e7168-344">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-344">Requirement</span></span>|<span data-ttu-id="e7168-345">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-346">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-347">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-347">1.0</span></span>|
|[<span data-ttu-id="e7168-348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-349">ReadItem</span></span>|
|[<span data-ttu-id="e7168-350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-351">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-352">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-352">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="e7168-353">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="e7168-353">dateTimeModified: Date</span></span>

<span data-ttu-id="e7168-p112">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e7168-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-356">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-356">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-357">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-357">Type</span></span>

*   <span data-ttu-id="e7168-358">Дата</span><span class="sxs-lookup"><span data-stu-id="e7168-358">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-359">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-359">Requirements</span></span>

|<span data-ttu-id="e7168-360">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-360">Requirement</span></span>|<span data-ttu-id="e7168-361">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-362">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-363">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-363">1.0</span></span>|
|[<span data-ttu-id="e7168-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-365">ReadItem</span></span>|
|[<span data-ttu-id="e7168-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-367">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-368">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-368">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="e7168-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-370">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-370">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="e7168-p113">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="e7168-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-373">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-373">Read mode</span></span>

<span data-ttu-id="e7168-374">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="e7168-374">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-375">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-375">Compose mode</span></span>

<span data-ttu-id="e7168-376">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="e7168-376">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="e7168-377">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="e7168-377">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="e7168-378">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-378">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e7168-379">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-379">Type</span></span>

*   <span data-ttu-id="e7168-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-381">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-381">Requirements</span></span>

|<span data-ttu-id="e7168-382">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-382">Requirement</span></span>|<span data-ttu-id="e7168-383">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-383">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-384">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-384">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-385">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-385">1.0</span></span>|
|[<span data-ttu-id="e7168-386">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-386">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-387">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-387">ReadItem</span></span>|
|[<span data-ttu-id="e7168-388">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-388">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-389">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-389">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocationviewoutlook-js-18"></a><span data-ttu-id="e7168-390">Енханцедлокатион: [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-390">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-391">Получает или задает расположение встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-391">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-392">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-392">Read mode</span></span>

<span data-ttu-id="e7168-393">Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) , который позволяет получить набор расположений (каждый, представленный объектом локатиондетаилс), связанный с встречей. [](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) `enhancedLocation`</span><span class="sxs-lookup"><span data-stu-id="e7168-393">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e7168-394">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-394">Compose mode</span></span>

<span data-ttu-id="e7168-395">`enhancedLocation` Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) , который предоставляет методы для получения, удаления или добавления расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-396">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-396">Type</span></span>

*   [<span data-ttu-id="e7168-397">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e7168-397">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="e7168-398">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-398">Requirements</span></span>

|<span data-ttu-id="e7168-399">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-399">Requirement</span></span>|<span data-ttu-id="e7168-400">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-400">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-401">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-401">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-402">1.8</span><span class="sxs-lookup"><span data-stu-id="e7168-402">1.8</span></span>|
|[<span data-ttu-id="e7168-403">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-403">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-404">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-404">ReadItem</span></span>|
|[<span data-ttu-id="e7168-405">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-405">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-406">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-406">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-407">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-407">Example</span></span>

<span data-ttu-id="e7168-408">В следующем примере показано получение текущих расположений, связанных с встречей.</span><span class="sxs-lookup"><span data-stu-id="e7168-408">The following example gets the current locations associated with the appointment.</span></span>

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18fromjavascriptapioutlookofficefromviewoutlook-js-18"></a><span data-ttu-id="e7168-409">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-409">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-410">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-410">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="e7168-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="e7168-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-413">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e7168-413">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-414">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-414">Read mode</span></span>

<span data-ttu-id="e7168-415">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="e7168-415">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-416">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-416">Compose mode</span></span>

<span data-ttu-id="e7168-417">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="e7168-417">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e7168-418">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-418">Type</span></span>

*   <span data-ttu-id="e7168-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [из](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-420">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-420">Requirements</span></span>

|<span data-ttu-id="e7168-421">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-421">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e7168-422">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-423">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-423">1.0</span></span>|<span data-ttu-id="e7168-424">1.7</span><span class="sxs-lookup"><span data-stu-id="e7168-424">1.7</span></span>|
|[<span data-ttu-id="e7168-425">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-426">ReadItem</span></span>|<span data-ttu-id="e7168-427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e7168-427">ReadWriteItem</span></span>|
|[<span data-ttu-id="e7168-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-429">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-429">Read</span></span>|<span data-ttu-id="e7168-430">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-430">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheadersviewoutlook-js-18"></a><span data-ttu-id="e7168-431">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-431">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-432">Возвращает или задает настраиваемые заголовки Интернета для сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-432">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="e7168-433">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e7168-433">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-434">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-434">Type</span></span>

*   [<span data-ttu-id="e7168-435">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="e7168-435">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="e7168-436">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-436">Requirements</span></span>

|<span data-ttu-id="e7168-437">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-437">Requirement</span></span>|<span data-ttu-id="e7168-438">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-439">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-440">1.8</span><span class="sxs-lookup"><span data-stu-id="e7168-440">1.8</span></span>|
|[<span data-ttu-id="e7168-441">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-442">ReadItem</span></span>|
|[<span data-ttu-id="e7168-443">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-444">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-444">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-445">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-445">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="e7168-446">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="e7168-446">internetMessageId: String</span></span>

<span data-ttu-id="e7168-p116">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e7168-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-449">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-449">Type</span></span>

*   <span data-ttu-id="e7168-450">String</span><span class="sxs-lookup"><span data-stu-id="e7168-450">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-451">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-451">Requirements</span></span>

|<span data-ttu-id="e7168-452">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-452">Requirement</span></span>|<span data-ttu-id="e7168-453">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-453">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-454">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-454">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-455">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-455">1.0</span></span>|
|[<span data-ttu-id="e7168-456">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-456">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-457">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-457">ReadItem</span></span>|
|[<span data-ttu-id="e7168-458">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-458">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-459">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-459">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-460">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-460">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="e7168-461">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="e7168-461">itemClass: String</span></span>

<span data-ttu-id="e7168-p117">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e7168-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="e7168-p118">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="e7168-466">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-466">Type</span></span>|<span data-ttu-id="e7168-467">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-467">Description</span></span>|<span data-ttu-id="e7168-468">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="e7168-468">item class</span></span>|
|---|---|---|
|<span data-ttu-id="e7168-469">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="e7168-469">Appointment items</span></span>|<span data-ttu-id="e7168-470">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="e7168-470">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="e7168-471">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="e7168-471">Message items</span></span>|<span data-ttu-id="e7168-472">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-472">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="e7168-473">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="e7168-473">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-474">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-474">Type</span></span>

*   <span data-ttu-id="e7168-475">String</span><span class="sxs-lookup"><span data-stu-id="e7168-475">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-476">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-476">Requirements</span></span>

|<span data-ttu-id="e7168-477">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-477">Requirement</span></span>|<span data-ttu-id="e7168-478">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-478">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-479">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-479">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-480">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-480">1.0</span></span>|
|[<span data-ttu-id="e7168-481">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-481">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-482">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-482">ReadItem</span></span>|
|[<span data-ttu-id="e7168-483">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-483">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-484">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-484">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-485">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-485">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="e7168-486">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="e7168-486">(nullable) itemId: String</span></span>

<span data-ttu-id="e7168-487">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-487">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="e7168-488">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e7168-488">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-489">Идентификатор, возвращаемый `itemId` свойством, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="e7168-489">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="e7168-490">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="e7168-490">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="e7168-491">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="e7168-491">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e7168-492">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="e7168-492">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="e7168-p121">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-495">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-495">Type</span></span>

*   <span data-ttu-id="e7168-496">String</span><span class="sxs-lookup"><span data-stu-id="e7168-496">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-497">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-497">Requirements</span></span>

|<span data-ttu-id="e7168-498">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-498">Requirement</span></span>|<span data-ttu-id="e7168-499">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-499">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-500">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-500">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-501">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-501">1.0</span></span>|
|[<span data-ttu-id="e7168-502">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-502">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-503">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-503">ReadItem</span></span>|
|[<span data-ttu-id="e7168-504">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-504">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-505">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-505">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-506">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-506">Example</span></span>

<span data-ttu-id="e7168-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-18"></a><span data-ttu-id="e7168-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-510">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="e7168-510">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="e7168-511">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="e7168-511">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-512">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-512">Type</span></span>

*   [<span data-ttu-id="e7168-513">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="e7168-513">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="e7168-514">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-514">Requirements</span></span>

|<span data-ttu-id="e7168-515">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-515">Requirement</span></span>|<span data-ttu-id="e7168-516">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-516">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-517">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-517">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-518">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-518">1.0</span></span>|
|[<span data-ttu-id="e7168-519">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-519">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-520">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-520">ReadItem</span></span>|
|[<span data-ttu-id="e7168-521">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-521">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-522">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-522">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-523">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-523">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-18"></a><span data-ttu-id="e7168-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-525">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-525">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-526">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-526">Read mode</span></span>

<span data-ttu-id="e7168-527">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-527">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-528">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-528">Compose mode</span></span>

<span data-ttu-id="e7168-529">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-529">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e7168-530">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-530">Type</span></span>

*   <span data-ttu-id="e7168-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-532">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-532">Requirements</span></span>

|<span data-ttu-id="e7168-533">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-533">Requirement</span></span>|<span data-ttu-id="e7168-534">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-535">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-536">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-536">1.0</span></span>|
|[<span data-ttu-id="e7168-537">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-537">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-538">ReadItem</span></span>|
|[<span data-ttu-id="e7168-539">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-539">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-540">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-540">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="e7168-541">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="e7168-541">normalizedSubject: String</span></span>

<span data-ttu-id="e7168-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e7168-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="e7168-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="e7168-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-546">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-546">Type</span></span>

*   <span data-ttu-id="e7168-547">String</span><span class="sxs-lookup"><span data-stu-id="e7168-547">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-548">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-548">Requirements</span></span>

|<span data-ttu-id="e7168-549">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-549">Requirement</span></span>|<span data-ttu-id="e7168-550">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-551">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-552">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-552">1.0</span></span>|
|[<span data-ttu-id="e7168-553">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-553">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-554">ReadItem</span></span>|
|[<span data-ttu-id="e7168-555">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-555">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-556">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-556">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-557">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-557">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-18"></a><span data-ttu-id="e7168-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-559">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-559">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-560">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-560">Type</span></span>

*   [<span data-ttu-id="e7168-561">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="e7168-561">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="e7168-562">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-562">Requirements</span></span>

|<span data-ttu-id="e7168-563">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-563">Requirement</span></span>|<span data-ttu-id="e7168-564">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-565">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-566">1.3</span><span class="sxs-lookup"><span data-stu-id="e7168-566">1.3</span></span>|
|[<span data-ttu-id="e7168-567">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-567">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-568">ReadItem</span></span>|
|[<span data-ttu-id="e7168-569">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-569">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-570">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-570">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-571">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-571">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="e7168-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-573">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="e7168-573">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="e7168-574">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-574">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-575">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-575">Read mode</span></span>

<span data-ttu-id="e7168-576">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-576">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="e7168-577">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-577">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e7168-578">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="e7168-578">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-579">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-579">Compose mode</span></span>

<span data-ttu-id="e7168-580">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-580">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="e7168-581">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-581">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e7168-582">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="e7168-582">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e7168-583">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-583">Get 500 members maximum.</span></span>
- <span data-ttu-id="e7168-584">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-584">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e7168-585">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-585">Type</span></span>

*   <span data-ttu-id="e7168-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-587">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-587">Requirements</span></span>

|<span data-ttu-id="e7168-588">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-588">Requirement</span></span>|<span data-ttu-id="e7168-589">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-589">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-590">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-591">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-591">1.0</span></span>|
|[<span data-ttu-id="e7168-592">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-592">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-593">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-593">ReadItem</span></span>|
|[<span data-ttu-id="e7168-594">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-595">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-595">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18organizerjavascriptapioutlookofficeorganizerviewoutlook-js-18"></a><span data-ttu-id="e7168-596">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[Организатор](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-596">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-597">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-597">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-598">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-598">Read mode</span></span>

<span data-ttu-id="e7168-599">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-599">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-600">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-600">Compose mode</span></span>

<span data-ttu-id="e7168-601">`organizer` Свойство возвращает объект [организатора](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) , который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="e7168-601">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="e7168-602">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-602">Type</span></span>

*   <span data-ttu-id="e7168-603">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Организатор](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e7168-603">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-604">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-604">Requirements</span></span>

|<span data-ttu-id="e7168-605">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-605">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e7168-606">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-607">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-607">1.0</span></span>|<span data-ttu-id="e7168-608">1.7</span><span class="sxs-lookup"><span data-stu-id="e7168-608">1.7</span></span>|
|[<span data-ttu-id="e7168-609">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-610">ReadItem</span></span>|<span data-ttu-id="e7168-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e7168-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="e7168-612">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-613">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-613">Read</span></span>|<span data-ttu-id="e7168-614">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-614">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-18"></a><span data-ttu-id="e7168-615">(Nullable) повторение: [повторение](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-615">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-616">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-616">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="e7168-617">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="e7168-617">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="e7168-618">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-618">Read and compose modes for appointment items.</span></span> <span data-ttu-id="e7168-619">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-619">Read mode for meeting request items.</span></span>

<span data-ttu-id="e7168-620">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="e7168-620">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="e7168-621">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="e7168-621">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="e7168-622">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-622">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="e7168-623">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="e7168-623">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="e7168-624">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="e7168-624">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-625">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-625">Read mode</span></span>

<span data-ttu-id="e7168-626">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-626">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that represents the appointment recurrence.</span></span> <span data-ttu-id="e7168-627">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-627">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-628">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-628">Compose mode</span></span>

<span data-ttu-id="e7168-629">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-629">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="e7168-630">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="e7168-630">This is available for appointments.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="e7168-631">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-631">Type</span></span>

* [<span data-ttu-id="e7168-632">Повторения</span><span class="sxs-lookup"><span data-stu-id="e7168-632">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)

|<span data-ttu-id="e7168-633">Requirement</span><span class="sxs-lookup"><span data-stu-id="e7168-633">Requirement</span></span>|<span data-ttu-id="e7168-634">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-635">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-636">1.7</span><span class="sxs-lookup"><span data-stu-id="e7168-636">1.7</span></span>|
|[<span data-ttu-id="e7168-637">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-638">ReadItem</span></span>|
|[<span data-ttu-id="e7168-639">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-640">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-640">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="e7168-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-642">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="e7168-642">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="e7168-643">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-643">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-644">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-644">Read mode</span></span>

<span data-ttu-id="e7168-645">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-645">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="e7168-646">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-646">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e7168-647">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="e7168-647">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-648">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-648">Compose mode</span></span>

<span data-ttu-id="e7168-649">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-649">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="e7168-650">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-650">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e7168-651">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="e7168-651">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e7168-652">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-652">Get 500 members maximum.</span></span>
- <span data-ttu-id="e7168-653">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-653">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="e7168-654">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-654">Type</span></span>

*   <span data-ttu-id="e7168-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-656">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-656">Requirements</span></span>

|<span data-ttu-id="e7168-657">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-657">Requirement</span></span>|<span data-ttu-id="e7168-658">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-658">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-659">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-659">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-660">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-660">1.0</span></span>|
|[<span data-ttu-id="e7168-661">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-661">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-662">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-662">ReadItem</span></span>|
|[<span data-ttu-id="e7168-663">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-663">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-664">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-664">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18"></a><span data-ttu-id="e7168-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-p135">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e7168-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="e7168-p136">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="e7168-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-670">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e7168-670">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-671">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-671">Type</span></span>

*   [<span data-ttu-id="e7168-672">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e7168-672">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="e7168-673">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-673">Requirements</span></span>

|<span data-ttu-id="e7168-674">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-674">Requirement</span></span>|<span data-ttu-id="e7168-675">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-676">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-677">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-677">1.0</span></span>|
|[<span data-ttu-id="e7168-678">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-679">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-679">ReadItem</span></span>|
|[<span data-ttu-id="e7168-680">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-681">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-681">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-682">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-682">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="e7168-683">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="e7168-683">(nullable) seriesId: String</span></span>

<span data-ttu-id="e7168-684">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="e7168-684">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="e7168-685">В Outlook в Интернете и на настольных клиентах `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="e7168-685">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="e7168-686">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-686">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-687">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="e7168-687">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e7168-688">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="e7168-688">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="e7168-689">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="e7168-689">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e7168-690">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="e7168-690">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="e7168-691">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-691">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="e7168-692">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-692">Type</span></span>

* <span data-ttu-id="e7168-693">String</span><span class="sxs-lookup"><span data-stu-id="e7168-693">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-694">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-694">Requirements</span></span>

|<span data-ttu-id="e7168-695">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-695">Requirement</span></span>|<span data-ttu-id="e7168-696">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-697">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-698">1.7</span><span class="sxs-lookup"><span data-stu-id="e7168-698">1.7</span></span>|
|[<span data-ttu-id="e7168-699">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-700">ReadItem</span></span>|
|[<span data-ttu-id="e7168-701">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-702">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-702">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-703">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-703">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="e7168-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-705">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-705">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="e7168-p139">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="e7168-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-708">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-708">Read mode</span></span>

<span data-ttu-id="e7168-709">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="e7168-709">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-710">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-710">Compose mode</span></span>

<span data-ttu-id="e7168-711">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="e7168-711">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="e7168-712">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="e7168-712">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="e7168-713">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e7168-713">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e7168-714">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-714">Type</span></span>

*   <span data-ttu-id="e7168-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-716">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-716">Requirements</span></span>

|<span data-ttu-id="e7168-717">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-717">Requirement</span></span>|<span data-ttu-id="e7168-718">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-719">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-719">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-720">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-720">1.0</span></span>|
|[<span data-ttu-id="e7168-721">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-721">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-722">ReadItem</span></span>|
|[<span data-ttu-id="e7168-723">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-723">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-724">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-724">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-18"></a><span data-ttu-id="e7168-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-726">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-726">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="e7168-727">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="e7168-727">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-728">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-728">Read mode</span></span>

<span data-ttu-id="e7168-p140">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="e7168-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="e7168-731">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="e7168-731">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-732">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-732">Compose mode</span></span>
<span data-ttu-id="e7168-733">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="e7168-733">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="e7168-734">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-734">Type</span></span>

*   <span data-ttu-id="e7168-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-736">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-736">Requirements</span></span>

|<span data-ttu-id="e7168-737">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-737">Requirement</span></span>|<span data-ttu-id="e7168-738">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-738">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-739">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-739">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-740">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-740">1.0</span></span>|
|[<span data-ttu-id="e7168-741">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-741">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-742">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-742">ReadItem</span></span>|
|[<span data-ttu-id="e7168-743">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-743">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-744">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-744">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="e7168-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-746">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-746">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="e7168-747">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-747">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e7168-748">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e7168-748">Read mode</span></span>

<span data-ttu-id="e7168-749">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-749">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="e7168-750">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-750">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e7168-751">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="e7168-751">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="e7168-752">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e7168-752">Compose mode</span></span>

<span data-ttu-id="e7168-753">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-753">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="e7168-754">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-754">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="e7168-755">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="e7168-755">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="e7168-756">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-756">Get 500 members maximum.</span></span>
- <span data-ttu-id="e7168-757">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="e7168-757">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e7168-758">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-758">Type</span></span>

*   <span data-ttu-id="e7168-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-760">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-760">Requirements</span></span>

|<span data-ttu-id="e7168-761">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-761">Requirement</span></span>|<span data-ttu-id="e7168-762">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-762">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-763">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-763">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-764">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-764">1.0</span></span>|
|[<span data-ttu-id="e7168-765">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-765">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-766">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-766">ReadItem</span></span>|
|[<span data-ttu-id="e7168-767">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-767">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-768">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-768">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e7168-769">Методы</span><span class="sxs-lookup"><span data-stu-id="e7168-769">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="e7168-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e7168-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e7168-771">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="e7168-771">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e7168-772">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="e7168-772">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="e7168-773">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e7168-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-774">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-774">Parameters</span></span>
|<span data-ttu-id="e7168-775">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-775">Name</span></span>|<span data-ttu-id="e7168-776">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-776">Type</span></span>|<span data-ttu-id="e7168-777">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-777">Attributes</span></span>|<span data-ttu-id="e7168-778">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-778">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="e7168-779">String</span><span class="sxs-lookup"><span data-stu-id="e7168-779">String</span></span>||<span data-ttu-id="e7168-p144">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="e7168-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e7168-782">String</span><span class="sxs-lookup"><span data-stu-id="e7168-782">String</span></span>||<span data-ttu-id="e7168-p145">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e7168-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e7168-785">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-785">Object</span></span>|<span data-ttu-id="e7168-786">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-786">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-787">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-787">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-788">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-788">Object</span></span>|<span data-ttu-id="e7168-789">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-789">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-790">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="e7168-790">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e7168-791">Boolean</span><span class="sxs-lookup"><span data-stu-id="e7168-791">Boolean</span></span>|<span data-ttu-id="e7168-792">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-792">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-793">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="e7168-793">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e7168-794">function</span><span class="sxs-lookup"><span data-stu-id="e7168-794">function</span></span>|<span data-ttu-id="e7168-795">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-795">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-796">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-796">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e7168-797">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e7168-797">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e7168-798">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="e7168-798">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e7168-799">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-799">Errors</span></span>

|<span data-ttu-id="e7168-800">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-800">Error code</span></span>|<span data-ttu-id="e7168-801">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-801">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e7168-802">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="e7168-802">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e7168-803">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="e7168-803">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e7168-804">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="e7168-804">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-805">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-805">Requirements</span></span>

|<span data-ttu-id="e7168-806">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-806">Requirement</span></span>|<span data-ttu-id="e7168-807">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-808">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-809">1.1</span><span class="sxs-lookup"><span data-stu-id="e7168-809">1.1</span></span>|
|[<span data-ttu-id="e7168-810">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-811">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e7168-811">ReadWriteItem</span></span>|
|[<span data-ttu-id="e7168-812">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-813">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-813">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e7168-814">Примеры</span><span class="sxs-lookup"><span data-stu-id="e7168-814">Examples</span></span>

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

<span data-ttu-id="e7168-815">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-815">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="e7168-816">addFileAttachmentFromBase64Async (base64File, Аттачментнаме, [параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="e7168-816">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e7168-817">Добавляет файл из кодировки Base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="e7168-817">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e7168-818">`addFileAttachmentFromBase64Async` Метод передает файл из кодировки Base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="e7168-818">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="e7168-819">Этот метод возвращает идентификатор вложения в объекте AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="e7168-819">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="e7168-820">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e7168-820">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-821">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-821">Parameters</span></span>

|<span data-ttu-id="e7168-822">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-822">Name</span></span>|<span data-ttu-id="e7168-823">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-823">Type</span></span>|<span data-ttu-id="e7168-824">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-824">Attributes</span></span>|<span data-ttu-id="e7168-825">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-825">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="e7168-826">String</span><span class="sxs-lookup"><span data-stu-id="e7168-826">String</span></span>||<span data-ttu-id="e7168-827">Содержимое изображения или файла в кодировке Base64, которое добавляется в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="e7168-827">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="e7168-828">String</span><span class="sxs-lookup"><span data-stu-id="e7168-828">String</span></span>||<span data-ttu-id="e7168-p147">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e7168-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e7168-831">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-831">Object</span></span>|<span data-ttu-id="e7168-832">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-832">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-833">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-833">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-834">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-834">Object</span></span>|<span data-ttu-id="e7168-835">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-835">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-836">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="e7168-836">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e7168-837">Boolean</span><span class="sxs-lookup"><span data-stu-id="e7168-837">Boolean</span></span>|<span data-ttu-id="e7168-838">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-838">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-839">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="e7168-839">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e7168-840">function</span><span class="sxs-lookup"><span data-stu-id="e7168-840">function</span></span>|<span data-ttu-id="e7168-841">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-841">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-842">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-842">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e7168-843">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e7168-843">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e7168-844">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="e7168-844">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e7168-845">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-845">Errors</span></span>

|<span data-ttu-id="e7168-846">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-846">Error code</span></span>|<span data-ttu-id="e7168-847">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-847">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e7168-848">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="e7168-848">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e7168-849">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="e7168-849">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e7168-850">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="e7168-850">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-851">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-851">Requirements</span></span>

|<span data-ttu-id="e7168-852">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-852">Requirement</span></span>|<span data-ttu-id="e7168-853">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-853">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-854">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-854">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-855">1.8</span><span class="sxs-lookup"><span data-stu-id="e7168-855">1.8</span></span>|
|[<span data-ttu-id="e7168-856">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-856">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-857">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e7168-857">ReadWriteItem</span></span>|
|[<span data-ttu-id="e7168-858">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-858">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-859">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-859">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e7168-860">Примеры</span><span class="sxs-lookup"><span data-stu-id="e7168-860">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e7168-861">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e7168-861">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e7168-862">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="e7168-862">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e7168-863">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="e7168-863">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-864">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-864">Parameters</span></span>

| <span data-ttu-id="e7168-865">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-865">Name</span></span> | <span data-ttu-id="e7168-866">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-866">Type</span></span> | <span data-ttu-id="e7168-867">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-867">Attributes</span></span> | <span data-ttu-id="e7168-868">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-868">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e7168-869">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e7168-869">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e7168-870">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="e7168-870">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e7168-871">Function</span><span class="sxs-lookup"><span data-stu-id="e7168-871">Function</span></span> || <span data-ttu-id="e7168-p148">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="e7168-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e7168-875">Объект</span><span class="sxs-lookup"><span data-stu-id="e7168-875">Object</span></span> | <span data-ttu-id="e7168-876">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-876">&lt;optional&gt;</span></span> | <span data-ttu-id="e7168-877">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-877">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e7168-878">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-878">Object</span></span> | <span data-ttu-id="e7168-879">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-879">&lt;optional&gt;</span></span> | <span data-ttu-id="e7168-880">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-880">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e7168-881">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-881">function</span></span>| <span data-ttu-id="e7168-882">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-882">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-883">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-884">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-884">Requirements</span></span>

|<span data-ttu-id="e7168-885">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-885">Requirement</span></span>| <span data-ttu-id="e7168-886">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-887">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7168-888">1.7</span><span class="sxs-lookup"><span data-stu-id="e7168-888">1.7</span></span> |
|[<span data-ttu-id="e7168-889">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7168-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-890">ReadItem</span></span> |
|[<span data-ttu-id="e7168-891">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7168-892">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-892">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="e7168-893">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-893">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="e7168-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e7168-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e7168-895">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="e7168-895">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="e7168-p149">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="e7168-899">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e7168-899">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="e7168-900">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="e7168-900">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-901">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-901">Parameters</span></span>

|<span data-ttu-id="e7168-902">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-902">Name</span></span>|<span data-ttu-id="e7168-903">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-903">Type</span></span>|<span data-ttu-id="e7168-904">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-904">Attributes</span></span>|<span data-ttu-id="e7168-905">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-905">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="e7168-906">String</span><span class="sxs-lookup"><span data-stu-id="e7168-906">String</span></span>||<span data-ttu-id="e7168-p150">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="e7168-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e7168-909">String</span><span class="sxs-lookup"><span data-stu-id="e7168-909">String</span></span>||<span data-ttu-id="e7168-910">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-910">The subject of the item to be attached.</span></span> <span data-ttu-id="e7168-911">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e7168-911">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e7168-912">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-912">Object</span></span>|<span data-ttu-id="e7168-913">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-913">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-914">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-914">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-915">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-915">Object</span></span>|<span data-ttu-id="e7168-916">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-916">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-917">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-917">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e7168-918">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-918">function</span></span>|<span data-ttu-id="e7168-919">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-919">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-920">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-920">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e7168-921">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e7168-921">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e7168-922">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="e7168-922">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e7168-923">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-923">Errors</span></span>

|<span data-ttu-id="e7168-924">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-924">Error code</span></span>|<span data-ttu-id="e7168-925">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-925">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e7168-926">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="e7168-926">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-927">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-927">Requirements</span></span>

|<span data-ttu-id="e7168-928">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-928">Requirement</span></span>|<span data-ttu-id="e7168-929">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-930">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-931">1.1</span><span class="sxs-lookup"><span data-stu-id="e7168-931">1.1</span></span>|
|[<span data-ttu-id="e7168-932">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-932">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-933">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e7168-933">ReadWriteItem</span></span>|
|[<span data-ttu-id="e7168-934">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-934">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-935">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-935">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-936">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-936">Example</span></span>

<span data-ttu-id="e7168-937">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="e7168-937">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="e7168-938">close()</span><span class="sxs-lookup"><span data-stu-id="e7168-938">close()</span></span>

<span data-ttu-id="e7168-939">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="e7168-939">Closes the current item that is being composed.</span></span>

<span data-ttu-id="e7168-p152">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="e7168-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-942">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="e7168-942">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="e7168-943">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="e7168-943">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-944">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-944">Requirements</span></span>

|<span data-ttu-id="e7168-945">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-945">Requirement</span></span>|<span data-ttu-id="e7168-946">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-947">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-948">1.3</span><span class="sxs-lookup"><span data-stu-id="e7168-948">1.3</span></span>|
|[<span data-ttu-id="e7168-949">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-950">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e7168-950">Restricted</span></span>|
|[<span data-ttu-id="e7168-951">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-952">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-952">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="e7168-953">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="e7168-953">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="e7168-954">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-954">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-955">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-955">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7168-956">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="e7168-956">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e7168-957">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="e7168-957">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="e7168-p153">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="e7168-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-961">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-961">Parameters</span></span>

|<span data-ttu-id="e7168-962">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-962">Name</span></span>|<span data-ttu-id="e7168-963">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-963">Type</span></span>|<span data-ttu-id="e7168-964">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-964">Attributes</span></span>|<span data-ttu-id="e7168-965">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-965">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e7168-966">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e7168-966">String &#124; Object</span></span>||<span data-ttu-id="e7168-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e7168-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e7168-969">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="e7168-969">**OR**</span></span><br/><span data-ttu-id="e7168-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="e7168-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e7168-972">String</span><span class="sxs-lookup"><span data-stu-id="e7168-972">String</span></span>|<span data-ttu-id="e7168-973">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-973">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e7168-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e7168-976">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-976">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e7168-977">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-977">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-978">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="e7168-978">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e7168-979">String</span><span class="sxs-lookup"><span data-stu-id="e7168-979">String</span></span>||<span data-ttu-id="e7168-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e7168-982">Строка</span><span class="sxs-lookup"><span data-stu-id="e7168-982">String</span></span>||<span data-ttu-id="e7168-983">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e7168-983">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e7168-984">Строка</span><span class="sxs-lookup"><span data-stu-id="e7168-984">String</span></span>||<span data-ttu-id="e7168-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="e7168-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e7168-987">Логический</span><span class="sxs-lookup"><span data-stu-id="e7168-987">Boolean</span></span>||<span data-ttu-id="e7168-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="e7168-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e7168-990">String</span><span class="sxs-lookup"><span data-stu-id="e7168-990">String</span></span>||<span data-ttu-id="e7168-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="e7168-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e7168-994">function</span><span class="sxs-lookup"><span data-stu-id="e7168-994">function</span></span>|<span data-ttu-id="e7168-995">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-995">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-996">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-996">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-997">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-997">Requirements</span></span>

|<span data-ttu-id="e7168-998">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-998">Requirement</span></span>|<span data-ttu-id="e7168-999">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1000">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-1001">1.0</span></span>|
|[<span data-ttu-id="e7168-1002">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1003">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1004">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1005">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1005">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e7168-1006">Примеры</span><span class="sxs-lookup"><span data-stu-id="e7168-1006">Examples</span></span>

<span data-ttu-id="e7168-1007">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1007">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="e7168-1008">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-1008">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="e7168-1009">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-1009">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e7168-1010">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="e7168-1010">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e7168-1011">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="e7168-1011">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e7168-1012">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="e7168-1012">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="e7168-1013">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="e7168-1013">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="e7168-1014">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-1014">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1015">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-1015">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7168-1016">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="e7168-1016">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e7168-1017">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="e7168-1017">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="e7168-p161">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="e7168-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1021">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1021">Parameters</span></span>

|<span data-ttu-id="e7168-1022">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1022">Name</span></span>|<span data-ttu-id="e7168-1023">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1023">Type</span></span>|<span data-ttu-id="e7168-1024">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1024">Attributes</span></span>|<span data-ttu-id="e7168-1025">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1025">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e7168-1026">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1026">String &#124; Object</span></span>||<span data-ttu-id="e7168-p162">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e7168-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e7168-1029">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="e7168-1029">**OR**</span></span><br/><span data-ttu-id="e7168-p163">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="e7168-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e7168-1032">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1032">String</span></span>|<span data-ttu-id="e7168-1033">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1033">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-p164">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e7168-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e7168-1036">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1036">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e7168-1037">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1038">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="e7168-1038">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e7168-1039">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1039">String</span></span>||<span data-ttu-id="e7168-p165">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e7168-1042">Строка</span><span class="sxs-lookup"><span data-stu-id="e7168-1042">String</span></span>||<span data-ttu-id="e7168-1043">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e7168-1043">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e7168-1044">Строка</span><span class="sxs-lookup"><span data-stu-id="e7168-1044">String</span></span>||<span data-ttu-id="e7168-p166">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="e7168-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e7168-1047">Логический</span><span class="sxs-lookup"><span data-stu-id="e7168-1047">Boolean</span></span>||<span data-ttu-id="e7168-p167">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="e7168-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e7168-1050">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1050">String</span></span>||<span data-ttu-id="e7168-p168">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="e7168-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e7168-1054">function</span><span class="sxs-lookup"><span data-stu-id="e7168-1054">function</span></span>|<span data-ttu-id="e7168-1055">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1055">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1056">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1056">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1057">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-1057">Requirements</span></span>

|<span data-ttu-id="e7168-1058">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1058">Requirement</span></span>|<span data-ttu-id="e7168-1059">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1059">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1060">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-1060">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1061">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-1061">1.0</span></span>|
|[<span data-ttu-id="e7168-1062">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1062">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1063">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1063">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1064">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1064">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1065">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1065">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e7168-1066">Примеры</span><span class="sxs-lookup"><span data-stu-id="e7168-1066">Examples</span></span>

<span data-ttu-id="e7168-1067">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1067">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="e7168-1068">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-1068">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="e7168-1069">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-1069">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e7168-1070">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="e7168-1070">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e7168-1071">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="e7168-1071">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e7168-1072">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="e7168-1072">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="e7168-1073">Жеталлинтернесеадерсасинк ([параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="e7168-1073">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="e7168-1074">Получает все заголовки Интернета для сообщения в виде строки.</span><span class="sxs-lookup"><span data-stu-id="e7168-1074">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="e7168-1075">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e7168-1075">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1076">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1076">Parameters</span></span>

|<span data-ttu-id="e7168-1077">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1077">Name</span></span>|<span data-ttu-id="e7168-1078">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1078">Type</span></span>|<span data-ttu-id="e7168-1079">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1079">Attributes</span></span>|<span data-ttu-id="e7168-1080">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1080">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e7168-1081">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1081">Object</span></span>|<span data-ttu-id="e7168-1082">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1083">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1083">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-1084">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1084">Object</span></span>|<span data-ttu-id="e7168-1085">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1085">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1086">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1086">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e7168-1087">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1087">function</span></span>|<span data-ttu-id="e7168-1088">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1088">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1089">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="e7168-1090">В случае успешного выполнения данные заголовков Интернета предоставляются в свойстве asyncResult. Value в виде String.</span><span class="sxs-lookup"><span data-stu-id="e7168-1090">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="e7168-1091">Сведения о форматировании возвращаемого строкового значения приведены в [RFC 2183](https://tools.ietf.org/html/rfc2183) .</span><span class="sxs-lookup"><span data-stu-id="e7168-1091">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="e7168-1092">Если происходит сбой вызова, свойство asyncResult. Error будет содержать код ошибки с причиной сбоя.</span><span class="sxs-lookup"><span data-stu-id="e7168-1092">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1093">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1093">Requirements</span></span>

|<span data-ttu-id="e7168-1094">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1094">Requirement</span></span>|<span data-ttu-id="e7168-1095">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1096">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1097">1.8</span><span class="sxs-lookup"><span data-stu-id="e7168-1097">1.8</span></span>|
|[<span data-ttu-id="e7168-1098">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1099">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1100">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1101">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1102">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1102">Returns:</span></span>

<span data-ttu-id="e7168-1103">Данные заголовков Интернета в виде строки, отформатированной в соответствии со [спецификацией RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="e7168-1103">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="e7168-1104">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="e7168-1104">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e7168-1105">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1105">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontentviewoutlook-js-18"></a><span data-ttu-id="e7168-1106">Жетаттачментконтентасинк (attachmentId, [параметры], [callback]) → [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-1106">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

<span data-ttu-id="e7168-1107">Получает указанное вложение из сообщения или встречи и возвращает его в виде `AttachmentContent` объекта.</span><span class="sxs-lookup"><span data-stu-id="e7168-1107">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="e7168-1108">`getAttachmentContentAsync` Метод получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-1108">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="e7168-1109">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, когда Аттачментидс был получен с помощью вызова `getAttachmentsAsync` или. `item.attachments`</span><span class="sxs-lookup"><span data-stu-id="e7168-1109">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="e7168-1110">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="e7168-1110">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="e7168-1111">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="e7168-1111">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1112">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1112">Parameters</span></span>

|<span data-ttu-id="e7168-1113">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1113">Name</span></span>|<span data-ttu-id="e7168-1114">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1114">Type</span></span>|<span data-ttu-id="e7168-1115">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1115">Attributes</span></span>|<span data-ttu-id="e7168-1116">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1116">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="e7168-1117">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1117">String</span></span>||<span data-ttu-id="e7168-1118">Идентификатор вложения, которое требуется получить.</span><span class="sxs-lookup"><span data-stu-id="e7168-1118">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="e7168-1119">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1119">Object</span></span>|<span data-ttu-id="e7168-1120">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1121">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-1122">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1122">Object</span></span>|<span data-ttu-id="e7168-1123">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1124">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e7168-1125">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1125">function</span></span>|<span data-ttu-id="e7168-1126">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1126">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1127">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1128">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-1128">Requirements</span></span>

|<span data-ttu-id="e7168-1129">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1129">Requirement</span></span>|<span data-ttu-id="e7168-1130">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1131">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1132">1.8</span><span class="sxs-lookup"><span data-stu-id="e7168-1132">1.8</span></span>|
|[<span data-ttu-id="e7168-1133">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1133">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1134">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1135">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1136">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1136">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1137">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1137">Returns:</span></span>

<span data-ttu-id="e7168-1138">Тип: [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-1138">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="e7168-1139">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1139">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="e7168-1140">Жетаттачментсасинк ([параметры], [обратный вызов]) → массив. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="e7168-1140">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="e7168-1141">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="e7168-1141">Gets the item's attachments as an array.</span></span> <span data-ttu-id="e7168-1142">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e7168-1142">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1143">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1143">Parameters</span></span>

|<span data-ttu-id="e7168-1144">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1144">Name</span></span>|<span data-ttu-id="e7168-1145">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1145">Type</span></span>|<span data-ttu-id="e7168-1146">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1146">Attributes</span></span>|<span data-ttu-id="e7168-1147">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1147">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e7168-1148">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1148">Object</span></span>|<span data-ttu-id="e7168-1149">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1149">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1150">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1150">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-1151">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1151">Object</span></span>|<span data-ttu-id="e7168-1152">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1152">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1153">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1153">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e7168-1154">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1154">function</span></span>|<span data-ttu-id="e7168-1155">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1155">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1156">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1157">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1157">Requirements</span></span>

|<span data-ttu-id="e7168-1158">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1158">Requirement</span></span>|<span data-ttu-id="e7168-1159">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1159">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1160">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1161">1.8</span><span class="sxs-lookup"><span data-stu-id="e7168-1161">1.8</span></span>|
|[<span data-ttu-id="e7168-1162">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1162">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1163">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1163">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1164">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1165">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-1165">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1166">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1166">Returns:</span></span>

<span data-ttu-id="e7168-1167">Тип: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="e7168-1167">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="example"></a><span data-ttu-id="e7168-1168">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1168">Example</span></span>

<span data-ttu-id="e7168-1169">В приведенном ниже примере создается строка HTML со сведениями обо всех вложениях в текущем элементе.</span><span class="sxs-lookup"><span data-stu-id="e7168-1169">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="e7168-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="e7168-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="e7168-1171">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-1171">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1172">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-1172">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-1173">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-1173">Requirements</span></span>

|<span data-ttu-id="e7168-1174">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1174">Requirement</span></span>|<span data-ttu-id="e7168-1175">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1175">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1176">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-1176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1177">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-1177">1.0</span></span>|
|[<span data-ttu-id="e7168-1178">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1178">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1179">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1181">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1181">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1182">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1182">Returns:</span></span>

<span data-ttu-id="e7168-1183">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-1183">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="e7168-1184">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1184">Example</span></span>

<span data-ttu-id="e7168-1185">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-1185">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="e7168-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="e7168-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="e7168-1187">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-1187">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1188">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-1188">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1189">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1189">Parameters</span></span>

|<span data-ttu-id="e7168-1190">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1190">Name</span></span>|<span data-ttu-id="e7168-1191">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1191">Type</span></span>|<span data-ttu-id="e7168-1192">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1192">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="e7168-1193">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="e7168-1193">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.8)|<span data-ttu-id="e7168-1194">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="e7168-1194">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1195">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1195">Requirements</span></span>

|<span data-ttu-id="e7168-1196">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1196">Requirement</span></span>|<span data-ttu-id="e7168-1197">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1197">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1198">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-1198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1199">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-1199">1.0</span></span>|
|[<span data-ttu-id="e7168-1200">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1201">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e7168-1201">Restricted</span></span>|
|[<span data-ttu-id="e7168-1202">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1203">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1203">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1204">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1204">Returns:</span></span>

<span data-ttu-id="e7168-1205">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="e7168-1205">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="e7168-1206">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="e7168-1206">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="e7168-1207">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1207">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="e7168-1208">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="e7168-1208">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="e7168-1209">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="e7168-1209">Value of `entityType`</span></span>|<span data-ttu-id="e7168-1210">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="e7168-1210">Type of objects in returned array</span></span>|<span data-ttu-id="e7168-1211">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1211">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="e7168-1212">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1212">String</span></span>|<span data-ttu-id="e7168-1213">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="e7168-1213">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="e7168-1214">Contact</span><span class="sxs-lookup"><span data-stu-id="e7168-1214">Contact</span></span>|<span data-ttu-id="e7168-1215">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e7168-1215">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="e7168-1216">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1216">String</span></span>|<span data-ttu-id="e7168-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e7168-1217">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="e7168-1218">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="e7168-1218">MeetingSuggestion</span></span>|<span data-ttu-id="e7168-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e7168-1219">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="e7168-1220">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="e7168-1220">PhoneNumber</span></span>|<span data-ttu-id="e7168-1221">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="e7168-1221">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="e7168-1222">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="e7168-1222">TaskSuggestion</span></span>|<span data-ttu-id="e7168-1223">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e7168-1223">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="e7168-1224">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1224">String</span></span>|<span data-ttu-id="e7168-1225">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="e7168-1225">**Restricted**</span></span>|

<span data-ttu-id="e7168-1226">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="e7168-1226">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

##### <a name="example"></a><span data-ttu-id="e7168-1227">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1227">Example</span></span>

<span data-ttu-id="e7168-1228">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-1228">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="e7168-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="e7168-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="e7168-1230">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e7168-1230">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1231">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-1231">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7168-1232">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1232">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1233">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1233">Parameters</span></span>

|<span data-ttu-id="e7168-1234">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1234">Name</span></span>|<span data-ttu-id="e7168-1235">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1235">Type</span></span>|<span data-ttu-id="e7168-1236">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1236">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e7168-1237">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1237">String</span></span>|<span data-ttu-id="e7168-1238">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="e7168-1238">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1239">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-1239">Requirements</span></span>

|<span data-ttu-id="e7168-1240">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1240">Requirement</span></span>|<span data-ttu-id="e7168-1241">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1241">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1242">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-1242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1243">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-1243">1.0</span></span>|
|[<span data-ttu-id="e7168-1244">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1245">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1245">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1246">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1247">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1247">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1248">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1248">Returns:</span></span>

<span data-ttu-id="e7168-p174">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="e7168-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="e7168-1251">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="e7168-1251">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="e7168-1252">Жетитемидасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="e7168-1252">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="e7168-1253">Асинхронно получает идентификатор сохраненного элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-1253">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="e7168-1254">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e7168-1254">Compose mode only.</span></span>

<span data-ttu-id="e7168-1255">При вызове этот метод возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1255">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1256">Если надстройка вызывает `getItemIdAsync` элемент в режиме создания (например, чтобы получить доступ `itemId` к использованию с помощью EWS или REST API), имейте в виду, что если Outlook находится в режиме кэширования, может потребоваться некоторое время до синхронизации элемента с сервером.</span><span class="sxs-lookup"><span data-stu-id="e7168-1256">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="e7168-1257">Пока элемент не будет синхронизирован, он не `itemId` распознается и не будет использоваться, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="e7168-1257">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1258">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1258">Parameters</span></span>

|<span data-ttu-id="e7168-1259">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1259">Name</span></span>|<span data-ttu-id="e7168-1260">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1260">Type</span></span>|<span data-ttu-id="e7168-1261">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1261">Attributes</span></span>|<span data-ttu-id="e7168-1262">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1262">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e7168-1263">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1263">Object</span></span>|<span data-ttu-id="e7168-1264">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1264">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1265">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1265">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-1266">Объект</span><span class="sxs-lookup"><span data-stu-id="e7168-1266">Object</span></span>|<span data-ttu-id="e7168-1267">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1268">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1268">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e7168-1269">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1269">function</span></span>||<span data-ttu-id="e7168-1270">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e7168-1271">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1271">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e7168-1272">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-1272">Errors</span></span>

|<span data-ttu-id="e7168-1273">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-1273">Error code</span></span>|<span data-ttu-id="e7168-1274">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1274">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="e7168-1275">Идентификатор невозможно извлечь, пока не будет сохранен элемент.</span><span class="sxs-lookup"><span data-stu-id="e7168-1275">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1276">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1276">Requirements</span></span>

|<span data-ttu-id="e7168-1277">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1277">Requirement</span></span>|<span data-ttu-id="e7168-1278">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1279">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1280">1.8</span><span class="sxs-lookup"><span data-stu-id="e7168-1280">1.8</span></span>|
|[<span data-ttu-id="e7168-1281">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1282">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1283">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1284">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e7168-1285">Примеры</span><span class="sxs-lookup"><span data-stu-id="e7168-1285">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="e7168-1286">В следующем примере показана структура `result` параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1286">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="e7168-1287">`value` Свойство содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-1287">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="e7168-1288">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e7168-1288">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="e7168-1289">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e7168-1289">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1290">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-1290">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7168-p178">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="e7168-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e7168-1294">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1294">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e7168-1295">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1295">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e7168-p179">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="e7168-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-1299">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1299">Requirements</span></span>

|<span data-ttu-id="e7168-1300">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1300">Requirement</span></span>|<span data-ttu-id="e7168-1301">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1302">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1303">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-1303">1.0</span></span>|
|[<span data-ttu-id="e7168-1304">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1305">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1306">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1307">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1307">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1308">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1308">Returns:</span></span>

<span data-ttu-id="e7168-p180">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="e7168-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="e7168-1311">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="e7168-1311">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e7168-1312">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1312">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e7168-1313">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1313">Example</span></span>

<span data-ttu-id="e7168-1314">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="e7168-1314">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="e7168-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="e7168-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="e7168-1316">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e7168-1316">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1317">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-1317">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7168-1318">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1318">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="e7168-p181">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="e7168-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1321">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1321">Parameters</span></span>

|<span data-ttu-id="e7168-1322">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1322">Name</span></span>|<span data-ttu-id="e7168-1323">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1323">Type</span></span>|<span data-ttu-id="e7168-1324">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1324">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e7168-1325">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1325">String</span></span>|<span data-ttu-id="e7168-1326">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="e7168-1326">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1327">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-1327">Requirements</span></span>

|<span data-ttu-id="e7168-1328">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1328">Requirement</span></span>|<span data-ttu-id="e7168-1329">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1329">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1330">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-1330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1331">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-1331">1.0</span></span>|
|[<span data-ttu-id="e7168-1332">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1332">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1333">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1334">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1334">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1335">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1335">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1336">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1336">Returns:</span></span>

<span data-ttu-id="e7168-1337">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e7168-1337">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="e7168-1338">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="e7168-1338">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="e7168-1339">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1339">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="e7168-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="e7168-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="e7168-1341">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-1341">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="e7168-p182">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="e7168-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1344">В Outlook в Интернете метод возвращает строку "null", если текст не выбран, но курсор находится в теле.</span><span class="sxs-lookup"><span data-stu-id="e7168-1344">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="e7168-1345">Чтобы проверить эту ситуацию, добавьте код, подобный приведенному ниже:</span><span class="sxs-lookup"><span data-stu-id="e7168-1345">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="e7168-1346">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1346">Parameters</span></span>

|<span data-ttu-id="e7168-1347">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1347">Name</span></span>|<span data-ttu-id="e7168-1348">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1348">Type</span></span>|<span data-ttu-id="e7168-1349">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1349">Attributes</span></span>|<span data-ttu-id="e7168-1350">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1350">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="e7168-1351">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e7168-1351">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="e7168-p184">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="e7168-p184">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="e7168-1355">Объект</span><span class="sxs-lookup"><span data-stu-id="e7168-1355">Object</span></span>|<span data-ttu-id="e7168-1356">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1356">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1357">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1357">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-1358">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1358">Object</span></span>|<span data-ttu-id="e7168-1359">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1359">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1360">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1360">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e7168-1361">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1361">function</span></span>||<span data-ttu-id="e7168-1362">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1362">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e7168-1363">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1363">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="e7168-1364">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1364">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1365">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1365">Requirements</span></span>

|<span data-ttu-id="e7168-1366">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1366">Requirement</span></span>|<span data-ttu-id="e7168-1367">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1367">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1368">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1369">1.2</span><span class="sxs-lookup"><span data-stu-id="e7168-1369">1.2</span></span>|
|[<span data-ttu-id="e7168-1370">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1371">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1372">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1373">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-1373">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1374">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1374">Returns:</span></span>

<span data-ttu-id="e7168-1375">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1375">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="e7168-1376">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="e7168-1376">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e7168-1377">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1377">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="e7168-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="e7168-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="e7168-1379">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="e7168-1379">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="e7168-1380">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="e7168-1380">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1381">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-1381">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-1382">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-1382">Requirements</span></span>

|<span data-ttu-id="e7168-1383">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1383">Requirement</span></span>|<span data-ttu-id="e7168-1384">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1384">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1385">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1386">1.6</span><span class="sxs-lookup"><span data-stu-id="e7168-1386">1.6</span></span>|
|[<span data-ttu-id="e7168-1387">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1387">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1388">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1389">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1389">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1390">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1390">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1391">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1391">Returns:</span></span>

<span data-ttu-id="e7168-1392">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="e7168-1392">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="e7168-1393">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1393">Example</span></span>

<span data-ttu-id="e7168-1394">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="e7168-1394">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="e7168-1395">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e7168-1395">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="e7168-p187">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="e7168-p187">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1398">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e7168-1398">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7168-p188">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="e7168-p188">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e7168-1402">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1402">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e7168-1403">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1403">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e7168-p189">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="e7168-p189">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7168-1407">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1407">Requirements</span></span>

|<span data-ttu-id="e7168-1408">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1408">Requirement</span></span>|<span data-ttu-id="e7168-1409">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1410">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1411">1.6</span><span class="sxs-lookup"><span data-stu-id="e7168-1411">1.6</span></span>|
|[<span data-ttu-id="e7168-1412">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1413">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1414">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1415">Чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1415">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7168-1416">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e7168-1416">Returns:</span></span>

<span data-ttu-id="e7168-p190">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="e7168-p190">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="e7168-1419">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1419">Example</span></span>

<span data-ttu-id="e7168-1420">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="e7168-1420">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="e7168-1421">Жетшаредпропертиесасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="e7168-1421">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="e7168-1422">Получает свойства выбранной встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="e7168-1422">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1423">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1423">Parameters</span></span>

|<span data-ttu-id="e7168-1424">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1424">Name</span></span>|<span data-ttu-id="e7168-1425">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1425">Type</span></span>|<span data-ttu-id="e7168-1426">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1426">Attributes</span></span>|<span data-ttu-id="e7168-1427">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1427">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e7168-1428">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1428">Object</span></span>|<span data-ttu-id="e7168-1429">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1429">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1430">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1430">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-1431">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1431">Object</span></span>|<span data-ttu-id="e7168-1432">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1433">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1433">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e7168-1434">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1434">function</span></span>||<span data-ttu-id="e7168-1435">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e7168-1436">Общие свойства предоставляются в виде [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) объекта в `asyncResult.value` свойстве.</span><span class="sxs-lookup"><span data-stu-id="e7168-1436">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e7168-1437">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-1437">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1438">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1438">Requirements</span></span>

|<span data-ttu-id="e7168-1439">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1439">Requirement</span></span>|<span data-ttu-id="e7168-1440">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1440">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1441">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1442">1.8</span><span class="sxs-lookup"><span data-stu-id="e7168-1442">1.8</span></span>|
|[<span data-ttu-id="e7168-1443">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1444">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1445">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1446">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1446">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-1447">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1447">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="e7168-1448">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e7168-1448">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="e7168-1449">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-1449">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="e7168-p192">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="e7168-p192">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1453">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1453">Parameters</span></span>

|<span data-ttu-id="e7168-1454">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1454">Name</span></span>|<span data-ttu-id="e7168-1455">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1455">Type</span></span>|<span data-ttu-id="e7168-1456">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1456">Attributes</span></span>|<span data-ttu-id="e7168-1457">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1457">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="e7168-1458">function</span><span class="sxs-lookup"><span data-stu-id="e7168-1458">function</span></span>||<span data-ttu-id="e7168-1459">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1459">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e7168-1460">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1460">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e7168-1461">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="e7168-1461">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="e7168-1462">Объект</span><span class="sxs-lookup"><span data-stu-id="e7168-1462">Object</span></span>|<span data-ttu-id="e7168-1463">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1464">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1464">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="e7168-1465">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1465">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1466">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1466">Requirements</span></span>

|<span data-ttu-id="e7168-1467">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1467">Requirement</span></span>|<span data-ttu-id="e7168-1468">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1468">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1469">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1470">1.0</span><span class="sxs-lookup"><span data-stu-id="e7168-1470">1.0</span></span>|
|[<span data-ttu-id="e7168-1471">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1472">ReadItem</span></span>|
|[<span data-ttu-id="e7168-1473">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1474">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1474">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-1475">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1475">Example</span></span>

<span data-ttu-id="e7168-p195">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-p195">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="e7168-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e7168-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="e7168-1480">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="e7168-1480">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="e7168-1481">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="e7168-1481">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="e7168-1482">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e7168-1482">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="e7168-1483">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="e7168-1483">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="e7168-1484">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="e7168-1484">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1485">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1485">Parameters</span></span>

|<span data-ttu-id="e7168-1486">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1486">Name</span></span>|<span data-ttu-id="e7168-1487">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1487">Type</span></span>|<span data-ttu-id="e7168-1488">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1488">Attributes</span></span>|<span data-ttu-id="e7168-1489">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1489">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="e7168-1490">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1490">String</span></span>||<span data-ttu-id="e7168-1491">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="e7168-1491">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="e7168-1492">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1492">Object</span></span>|<span data-ttu-id="e7168-1493">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1493">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1494">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1494">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-1495">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1495">Object</span></span>|<span data-ttu-id="e7168-1496">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1496">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1497">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1497">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e7168-1498">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1498">function</span></span>|<span data-ttu-id="e7168-1499">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1499">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1500">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e7168-1501">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="e7168-1501">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e7168-1502">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-1502">Errors</span></span>

|<span data-ttu-id="e7168-1503">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e7168-1503">Error code</span></span>|<span data-ttu-id="e7168-1504">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1504">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="e7168-1505">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="e7168-1505">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1506">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1506">Requirements</span></span>

|<span data-ttu-id="e7168-1507">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1507">Requirement</span></span>|<span data-ttu-id="e7168-1508">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1508">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1509">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7168-1509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1510">1.1</span><span class="sxs-lookup"><span data-stu-id="e7168-1510">1.1</span></span>|
|[<span data-ttu-id="e7168-1511">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1512">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1512">ReadWriteItem</span></span>|
|[<span data-ttu-id="e7168-1513">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1514">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-1514">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-1515">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1515">Example</span></span>

<span data-ttu-id="e7168-1516">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="e7168-1516">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="e7168-1517">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e7168-1517">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="e7168-1518">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="e7168-1518">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="e7168-1519">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1519">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1520">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1520">Parameters</span></span>

| <span data-ttu-id="e7168-1521">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1521">Name</span></span> | <span data-ttu-id="e7168-1522">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1522">Type</span></span> | <span data-ttu-id="e7168-1523">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1523">Attributes</span></span> | <span data-ttu-id="e7168-1524">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1524">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e7168-1525">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e7168-1525">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e7168-1526">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="e7168-1526">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="e7168-1527">Объект</span><span class="sxs-lookup"><span data-stu-id="e7168-1527">Object</span></span> | <span data-ttu-id="e7168-1528">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1528">&lt;optional&gt;</span></span> | <span data-ttu-id="e7168-1529">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1529">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e7168-1530">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1530">Object</span></span> | <span data-ttu-id="e7168-1531">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1531">&lt;optional&gt;</span></span> | <span data-ttu-id="e7168-1532">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1532">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e7168-1533">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1533">function</span></span>| <span data-ttu-id="e7168-1534">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1534">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1535">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1535">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1536">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-1536">Requirements</span></span>

|<span data-ttu-id="e7168-1537">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1537">Requirement</span></span>| <span data-ttu-id="e7168-1538">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1538">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1539">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7168-1540">1.7</span><span class="sxs-lookup"><span data-stu-id="e7168-1540">1.7</span></span> |
|[<span data-ttu-id="e7168-1541">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7168-1542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1542">ReadItem</span></span> |
|[<span data-ttu-id="e7168-1543">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7168-1544">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7168-1544">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="e7168-1545">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e7168-1545">saveAsync([options], callback)</span></span>

<span data-ttu-id="e7168-1546">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="e7168-1546">Asynchronously saves an item.</span></span>

<span data-ttu-id="e7168-1547">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1547">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="e7168-1548">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="e7168-1548">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="e7168-1549">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="e7168-1549">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1550">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="e7168-1550">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="e7168-1551">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="e7168-1551">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="e7168-p199">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="e7168-p199">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="e7168-1555">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="e7168-1555">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="e7168-1556">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="e7168-1556">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="e7168-1557">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e7168-1557">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="e7168-1558">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="e7168-1558">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="e7168-1559">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e7168-1559">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1560">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1560">Parameters</span></span>

|<span data-ttu-id="e7168-1561">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1561">Name</span></span>|<span data-ttu-id="e7168-1562">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1562">Type</span></span>|<span data-ttu-id="e7168-1563">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1563">Attributes</span></span>|<span data-ttu-id="e7168-1564">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1564">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e7168-1565">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1565">Object</span></span>|<span data-ttu-id="e7168-1566">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1566">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1567">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1567">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-1568">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1568">Object</span></span>|<span data-ttu-id="e7168-1569">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1569">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1570">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7168-1570">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e7168-1571">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1571">function</span></span>||<span data-ttu-id="e7168-1572">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1572">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e7168-1573">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1573">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1574">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7168-1574">Requirements</span></span>

|<span data-ttu-id="e7168-1575">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1575">Requirement</span></span>|<span data-ttu-id="e7168-1576">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1576">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1577">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1578">1.3</span><span class="sxs-lookup"><span data-stu-id="e7168-1578">1.3</span></span>|
|[<span data-ttu-id="e7168-1579">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1580">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1580">ReadWriteItem</span></span>|
|[<span data-ttu-id="e7168-1581">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1582">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-1582">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e7168-1583">Примеры</span><span class="sxs-lookup"><span data-stu-id="e7168-1583">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="e7168-p201">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e7168-p201">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="e7168-1586">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="e7168-1586">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="e7168-1587">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7168-1587">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="e7168-p202">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="e7168-p202">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7168-1591">Параметры</span><span class="sxs-lookup"><span data-stu-id="e7168-1591">Parameters</span></span>

|<span data-ttu-id="e7168-1592">Имя</span><span class="sxs-lookup"><span data-stu-id="e7168-1592">Name</span></span>|<span data-ttu-id="e7168-1593">Тип</span><span class="sxs-lookup"><span data-stu-id="e7168-1593">Type</span></span>|<span data-ttu-id="e7168-1594">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e7168-1594">Attributes</span></span>|<span data-ttu-id="e7168-1595">Описание</span><span class="sxs-lookup"><span data-stu-id="e7168-1595">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="e7168-1596">String</span><span class="sxs-lookup"><span data-stu-id="e7168-1596">String</span></span>||<span data-ttu-id="e7168-p203">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="e7168-p203">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="e7168-1600">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1600">Object</span></span>|<span data-ttu-id="e7168-1601">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1602">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e7168-1602">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e7168-1603">Object</span><span class="sxs-lookup"><span data-stu-id="e7168-1603">Object</span></span>|<span data-ttu-id="e7168-1604">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1604">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1605">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="e7168-1605">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="e7168-1606">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e7168-1606">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="e7168-1607">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e7168-1607">&lt;optional&gt;</span></span>|<span data-ttu-id="e7168-1608">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="e7168-1608">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="e7168-1609">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="e7168-1609">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="e7168-1610">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e7168-1610">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="e7168-1611">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="e7168-1611">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="e7168-1612">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="e7168-1612">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="e7168-1613">функция</span><span class="sxs-lookup"><span data-stu-id="e7168-1613">function</span></span>||<span data-ttu-id="e7168-1614">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e7168-1614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7168-1615">Требования</span><span class="sxs-lookup"><span data-stu-id="e7168-1615">Requirements</span></span>

|<span data-ttu-id="e7168-1616">Требование</span><span class="sxs-lookup"><span data-stu-id="e7168-1616">Requirement</span></span>|<span data-ttu-id="e7168-1617">Значение</span><span class="sxs-lookup"><span data-stu-id="e7168-1617">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7168-1618">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7168-1618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e7168-1619">1.2</span><span class="sxs-lookup"><span data-stu-id="e7168-1619">1.2</span></span>|
|[<span data-ttu-id="e7168-1620">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e7168-1620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e7168-1621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e7168-1621">ReadWriteItem</span></span>|
|[<span data-ttu-id="e7168-1622">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7168-1622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e7168-1623">Создание</span><span class="sxs-lookup"><span data-stu-id="e7168-1623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e7168-1624">Пример</span><span class="sxs-lookup"><span data-stu-id="e7168-1624">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
