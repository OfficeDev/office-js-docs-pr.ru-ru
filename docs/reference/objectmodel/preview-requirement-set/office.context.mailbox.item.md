---
title: Office. Context. Mailbox. Item — Предварительная версия набора требований
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 3ccafccd0c84ab243572421609083f56e3f7dfb1
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902237"
---
# <a name="item"></a><span data-ttu-id="48f42-102">item</span><span class="sxs-lookup"><span data-stu-id="48f42-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="48f42-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="48f42-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="48f42-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="48f42-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="48f42-106">Requirements</span></span>

|<span data-ttu-id="48f42-107">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-107">Requirement</span></span>|<span data-ttu-id="48f42-108">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-110">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-110">1.0</span></span>|
|[<span data-ttu-id="48f42-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="48f42-112">Restricted</span></span>|
|[<span data-ttu-id="48f42-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="48f42-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="48f42-115">Members and methods</span></span>

| <span data-ttu-id="48f42-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-116">Member</span></span> | <span data-ttu-id="48f42-117">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="48f42-118">attachments</span><span class="sxs-lookup"><span data-stu-id="48f42-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="48f42-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-119">Member</span></span> |
| [<span data-ttu-id="48f42-120">bcc</span><span class="sxs-lookup"><span data-stu-id="48f42-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="48f42-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-121">Member</span></span> |
| [<span data-ttu-id="48f42-122">body</span><span class="sxs-lookup"><span data-stu-id="48f42-122">body</span></span>](#body-body) | <span data-ttu-id="48f42-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-123">Member</span></span> |
| [<span data-ttu-id="48f42-124">разделов</span><span class="sxs-lookup"><span data-stu-id="48f42-124">categories</span></span>](#categories-categories) | <span data-ttu-id="48f42-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-125">Member</span></span> |
| [<span data-ttu-id="48f42-126">cc</span><span class="sxs-lookup"><span data-stu-id="48f42-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="48f42-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-127">Member</span></span> |
| [<span data-ttu-id="48f42-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="48f42-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="48f42-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-129">Member</span></span> |
| [<span data-ttu-id="48f42-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="48f42-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="48f42-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-131">Member</span></span> |
| [<span data-ttu-id="48f42-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="48f42-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="48f42-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-133">Member</span></span> |
| [<span data-ttu-id="48f42-134">end</span><span class="sxs-lookup"><span data-stu-id="48f42-134">end</span></span>](#end-datetime) | <span data-ttu-id="48f42-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-135">Member</span></span> |
| [<span data-ttu-id="48f42-136">енханцедлокатион</span><span class="sxs-lookup"><span data-stu-id="48f42-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="48f42-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-137">Member</span></span> |
| [<span data-ttu-id="48f42-138">from</span><span class="sxs-lookup"><span data-stu-id="48f42-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="48f42-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-139">Member</span></span> |
| [<span data-ttu-id="48f42-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="48f42-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="48f42-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-141">Member</span></span> |
| [<span data-ttu-id="48f42-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="48f42-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="48f42-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-143">Member</span></span> |
| [<span data-ttu-id="48f42-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="48f42-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="48f42-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-145">Member</span></span> |
| [<span data-ttu-id="48f42-146">itemId</span><span class="sxs-lookup"><span data-stu-id="48f42-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="48f42-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-147">Member</span></span> |
| [<span data-ttu-id="48f42-148">itemType</span><span class="sxs-lookup"><span data-stu-id="48f42-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="48f42-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-149">Member</span></span> |
| [<span data-ttu-id="48f42-150">location</span><span class="sxs-lookup"><span data-stu-id="48f42-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="48f42-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-151">Member</span></span> |
| [<span data-ttu-id="48f42-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="48f42-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="48f42-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-153">Member</span></span> |
| [<span data-ttu-id="48f42-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="48f42-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="48f42-155">Member</span><span class="sxs-lookup"><span data-stu-id="48f42-155">Member</span></span> |
| [<span data-ttu-id="48f42-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="48f42-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="48f42-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-157">Member</span></span> |
| [<span data-ttu-id="48f42-158">organizer</span><span class="sxs-lookup"><span data-stu-id="48f42-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="48f42-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-159">Member</span></span> |
| [<span data-ttu-id="48f42-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="48f42-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="48f42-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-161">Member</span></span> |
| [<span data-ttu-id="48f42-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="48f42-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="48f42-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-163">Member</span></span> |
| [<span data-ttu-id="48f42-164">sender</span><span class="sxs-lookup"><span data-stu-id="48f42-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="48f42-165">Member</span><span class="sxs-lookup"><span data-stu-id="48f42-165">Member</span></span> |
| [<span data-ttu-id="48f42-166">seriesId</span><span class="sxs-lookup"><span data-stu-id="48f42-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="48f42-167">Member</span><span class="sxs-lookup"><span data-stu-id="48f42-167">Member</span></span> |
| [<span data-ttu-id="48f42-168">start</span><span class="sxs-lookup"><span data-stu-id="48f42-168">start</span></span>](#start-datetime) | <span data-ttu-id="48f42-169">Member</span><span class="sxs-lookup"><span data-stu-id="48f42-169">Member</span></span> |
| [<span data-ttu-id="48f42-170">subject</span><span class="sxs-lookup"><span data-stu-id="48f42-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="48f42-171">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-171">Member</span></span> |
| [<span data-ttu-id="48f42-172">to</span><span class="sxs-lookup"><span data-stu-id="48f42-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="48f42-173">Элемент</span><span class="sxs-lookup"><span data-stu-id="48f42-173">Member</span></span> |
| [<span data-ttu-id="48f42-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="48f42-175">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-175">Method</span></span> |
| [<span data-ttu-id="48f42-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="48f42-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="48f42-177">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-177">Method</span></span> |
| [<span data-ttu-id="48f42-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="48f42-179">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-179">Method</span></span> |
| [<span data-ttu-id="48f42-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="48f42-181">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-181">Method</span></span> |
| [<span data-ttu-id="48f42-182">close</span><span class="sxs-lookup"><span data-stu-id="48f42-182">close</span></span>](#close) | <span data-ttu-id="48f42-183">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-183">Method</span></span> |
| [<span data-ttu-id="48f42-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="48f42-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="48f42-185">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-185">Method</span></span> |
| [<span data-ttu-id="48f42-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="48f42-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="48f42-187">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-187">Method</span></span> |
| [<span data-ttu-id="48f42-188">жеталлинтернесеадерсасинк</span><span class="sxs-lookup"><span data-stu-id="48f42-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="48f42-189">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-189">Method</span></span> |
| [<span data-ttu-id="48f42-190">жетаттачментконтентасинк</span><span class="sxs-lookup"><span data-stu-id="48f42-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="48f42-191">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-191">Method</span></span> |
| [<span data-ttu-id="48f42-192">жетаттачментсасинк</span><span class="sxs-lookup"><span data-stu-id="48f42-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="48f42-193">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-193">Method</span></span> |
| [<span data-ttu-id="48f42-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="48f42-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="48f42-195">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-195">Method</span></span> |
| [<span data-ttu-id="48f42-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="48f42-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="48f42-197">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-197">Method</span></span> |
| [<span data-ttu-id="48f42-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="48f42-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="48f42-199">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-199">Method</span></span> |
| [<span data-ttu-id="48f42-200">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-200">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="48f42-201">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-201">Method</span></span> |
| [<span data-ttu-id="48f42-202">жетитемидасинк</span><span class="sxs-lookup"><span data-stu-id="48f42-202">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="48f42-203">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-203">Method</span></span> |
| [<span data-ttu-id="48f42-204">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="48f42-204">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="48f42-205">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-205">Method</span></span> |
| [<span data-ttu-id="48f42-206">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="48f42-206">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="48f42-207">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-207">Method</span></span> |
| [<span data-ttu-id="48f42-208">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-208">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="48f42-209">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-209">Method</span></span> |
| [<span data-ttu-id="48f42-210">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="48f42-210">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="48f42-211">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-211">Method</span></span> |
| [<span data-ttu-id="48f42-212">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="48f42-212">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="48f42-213">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-213">Method</span></span> |
| [<span data-ttu-id="48f42-214">жетшаредпропертиесасинк</span><span class="sxs-lookup"><span data-stu-id="48f42-214">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="48f42-215">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-215">Method</span></span> |
| [<span data-ttu-id="48f42-216">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-216">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="48f42-217">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-217">Method</span></span> |
| [<span data-ttu-id="48f42-218">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-218">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="48f42-219">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-219">Method</span></span> |
| [<span data-ttu-id="48f42-220">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-220">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="48f42-221">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-221">Method</span></span> |
| [<span data-ttu-id="48f42-222">saveAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-222">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="48f42-223">Method</span><span class="sxs-lookup"><span data-stu-id="48f42-223">Method</span></span> |
| [<span data-ttu-id="48f42-224">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="48f42-224">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="48f42-225">Метод</span><span class="sxs-lookup"><span data-stu-id="48f42-225">Method</span></span> |

### <a name="example"></a><span data-ttu-id="48f42-226">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-226">Example</span></span>

<span data-ttu-id="48f42-227">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="48f42-227">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="48f42-228">Members</span><span class="sxs-lookup"><span data-stu-id="48f42-228">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="48f42-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="48f42-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="48f42-230">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="48f42-230">Gets the item's attachments as an array.</span></span> <span data-ttu-id="48f42-231">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="48f42-231">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-232">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="48f42-232">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="48f42-233">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="48f42-233">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-234">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-234">Type</span></span>

*   <span data-ttu-id="48f42-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="48f42-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-236">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-236">Requirements</span></span>

|<span data-ttu-id="48f42-237">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-237">Requirement</span></span>|<span data-ttu-id="48f42-238">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-239">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-240">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-240">1.0</span></span>|
|[<span data-ttu-id="48f42-241">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-242">ReadItem</span></span>|
|[<span data-ttu-id="48f42-243">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-244">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-244">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-245">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-245">Example</span></span>

<span data-ttu-id="48f42-246">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-246">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48f42-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48f42-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48f42-248">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-248">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="48f42-249">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="48f42-249">Compose mode only.</span></span>

<span data-ttu-id="48f42-250">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48f42-251">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="48f42-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48f42-252">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="48f42-253">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-254">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-254">Type</span></span>

*   [<span data-ttu-id="48f42-255">Получатели</span><span class="sxs-lookup"><span data-stu-id="48f42-255">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="48f42-256">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-256">Requirements</span></span>

|<span data-ttu-id="48f42-257">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-257">Requirement</span></span>|<span data-ttu-id="48f42-258">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-259">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-260">1.1</span><span class="sxs-lookup"><span data-stu-id="48f42-260">1.1</span></span>|
|[<span data-ttu-id="48f42-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-262">ReadItem</span></span>|
|[<span data-ttu-id="48f42-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-264">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-264">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-265">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-265">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="48f42-266">body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="48f42-266">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="48f42-267">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-267">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-268">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-268">Type</span></span>

*   [<span data-ttu-id="48f42-269">Body</span><span class="sxs-lookup"><span data-stu-id="48f42-269">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="48f42-270">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-270">Requirements</span></span>

|<span data-ttu-id="48f42-271">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-271">Requirement</span></span>|<span data-ttu-id="48f42-272">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-273">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-274">1.1</span><span class="sxs-lookup"><span data-stu-id="48f42-274">1.1</span></span>|
|[<span data-ttu-id="48f42-275">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-276">ReadItem</span></span>|
|[<span data-ttu-id="48f42-277">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-278">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-279">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-279">Example</span></span>

<span data-ttu-id="48f42-280">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="48f42-280">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="48f42-281">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-281">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="48f42-282">Категории: [категории](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="48f42-282">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="48f42-283">Получает объект, предоставляющий методы для управления категориями элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-283">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-284">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-284">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-285">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-285">Type</span></span>

*   [<span data-ttu-id="48f42-286">Categories</span><span class="sxs-lookup"><span data-stu-id="48f42-286">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="48f42-287">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-287">Requirements</span></span>

|<span data-ttu-id="48f42-288">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-288">Requirement</span></span>|<span data-ttu-id="48f42-289">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-290">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-291">1.8</span><span class="sxs-lookup"><span data-stu-id="48f42-291">1.8</span></span>|
|[<span data-ttu-id="48f42-292">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-293">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-293">ReadItem</span></span>|
|[<span data-ttu-id="48f42-294">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-295">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-295">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-296">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-296">Example</span></span>

<span data-ttu-id="48f42-297">В этом примере возвращаются категории элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-297">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48f42-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48f42-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48f42-299">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-299">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="48f42-300">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-300">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-301">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-301">Read mode</span></span>

<span data-ttu-id="48f42-302">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-302">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="48f42-303">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-303">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48f42-304">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="48f42-304">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-305">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-305">Compose mode</span></span>

<span data-ttu-id="48f42-306">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-306">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="48f42-307">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-307">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48f42-308">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="48f42-308">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48f42-309">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-309">Get 500 members maximum.</span></span>
- <span data-ttu-id="48f42-310">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-310">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48f42-311">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-311">Type</span></span>

*   <span data-ttu-id="48f42-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48f42-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-313">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-313">Requirements</span></span>

|<span data-ttu-id="48f42-314">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-314">Requirement</span></span>|<span data-ttu-id="48f42-315">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-316">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-317">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-317">1.0</span></span>|
|[<span data-ttu-id="48f42-318">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-319">ReadItem</span></span>|
|[<span data-ttu-id="48f42-320">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-321">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-321">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="48f42-322">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="48f42-322">(nullable) conversationId: String</span></span>

<span data-ttu-id="48f42-323">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="48f42-323">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="48f42-p109">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="48f42-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="48f42-p110">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="48f42-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-328">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-328">Type</span></span>

*   <span data-ttu-id="48f42-329">String</span><span class="sxs-lookup"><span data-stu-id="48f42-329">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-330">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-330">Requirements</span></span>

|<span data-ttu-id="48f42-331">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-331">Requirement</span></span>|<span data-ttu-id="48f42-332">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-333">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-334">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-334">1.0</span></span>|
|[<span data-ttu-id="48f42-335">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-336">ReadItem</span></span>|
|[<span data-ttu-id="48f42-337">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-338">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-338">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-339">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-339">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="48f42-340">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="48f42-340">dateTimeCreated: Date</span></span>

<span data-ttu-id="48f42-p111">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="48f42-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-343">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-343">Type</span></span>

*   <span data-ttu-id="48f42-344">Дата</span><span class="sxs-lookup"><span data-stu-id="48f42-344">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-345">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-345">Requirements</span></span>

|<span data-ttu-id="48f42-346">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-346">Requirement</span></span>|<span data-ttu-id="48f42-347">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-348">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-349">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-349">1.0</span></span>|
|[<span data-ttu-id="48f42-350">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-350">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-351">ReadItem</span></span>|
|[<span data-ttu-id="48f42-352">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-352">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-353">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-353">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-354">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-354">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="48f42-355">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="48f42-355">dateTimeModified: Date</span></span>

<span data-ttu-id="48f42-p112">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="48f42-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-358">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-358">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-359">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-359">Type</span></span>

*   <span data-ttu-id="48f42-360">Дата</span><span class="sxs-lookup"><span data-stu-id="48f42-360">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-361">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-361">Requirements</span></span>

|<span data-ttu-id="48f42-362">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-362">Requirement</span></span>|<span data-ttu-id="48f42-363">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-363">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-364">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-364">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-365">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-365">1.0</span></span>|
|[<span data-ttu-id="48f42-366">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-366">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-367">ReadItem</span></span>|
|[<span data-ttu-id="48f42-368">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-368">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-369">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-369">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-370">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-370">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="48f42-371">end: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="48f42-371">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="48f42-372">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-372">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="48f42-p113">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="48f42-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-375">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-375">Read mode</span></span>

<span data-ttu-id="48f42-376">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="48f42-376">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-377">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-377">Compose mode</span></span>

<span data-ttu-id="48f42-378">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="48f42-378">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="48f42-379">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="48f42-379">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="48f42-380">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-380">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="48f42-381">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-381">Type</span></span>

*   <span data-ttu-id="48f42-382">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="48f42-382">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-383">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-383">Requirements</span></span>

|<span data-ttu-id="48f42-384">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-384">Requirement</span></span>|<span data-ttu-id="48f42-385">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-386">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-387">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-387">1.0</span></span>|
|[<span data-ttu-id="48f42-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-389">ReadItem</span></span>|
|[<span data-ttu-id="48f42-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-391">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-391">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="48f42-392">Енханцедлокатион: [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="48f42-392">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="48f42-393">Получает или задает расположение встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-393">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-394">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-394">Read mode</span></span>

<span data-ttu-id="48f42-395">Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который позволяет получить набор расположений (каждый, представленный объектом локатиондетаилс), связанный с встречей. [](/javascript/api/outlook/office.locationdetails) `enhancedLocation`</span><span class="sxs-lookup"><span data-stu-id="48f42-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="48f42-396">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-396">Compose mode</span></span>

<span data-ttu-id="48f42-397">`enhancedLocation` Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который предоставляет методы для получения, удаления или добавления расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-397">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-398">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-398">Type</span></span>

*   [<span data-ttu-id="48f42-399">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="48f42-399">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="48f42-400">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-400">Requirements</span></span>

|<span data-ttu-id="48f42-401">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-401">Requirement</span></span>|<span data-ttu-id="48f42-402">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-403">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-404">1.8</span><span class="sxs-lookup"><span data-stu-id="48f42-404">1.8</span></span>|
|[<span data-ttu-id="48f42-405">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-405">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-406">ReadItem</span></span>|
|[<span data-ttu-id="48f42-407">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-407">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-408">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-408">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-409">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-409">Example</span></span>

<span data-ttu-id="48f42-410">В следующем примере показано получение текущих расположений, связанных с встречей.</span><span class="sxs-lookup"><span data-stu-id="48f42-410">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="48f42-411">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="48f42-411">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="48f42-412">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-412">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="48f42-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="48f42-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-415">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="48f42-415">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-416">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-416">Read mode</span></span>

<span data-ttu-id="48f42-417">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="48f42-417">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-418">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-418">Compose mode</span></span>

<span data-ttu-id="48f42-419">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="48f42-419">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48f42-420">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-420">Type</span></span>

*   <span data-ttu-id="48f42-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [из](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="48f42-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-422">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-422">Requirements</span></span>

|<span data-ttu-id="48f42-423">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-423">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="48f42-424">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-425">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-425">1.0</span></span>|<span data-ttu-id="48f42-426">1.7</span><span class="sxs-lookup"><span data-stu-id="48f42-426">1.7</span></span>|
|[<span data-ttu-id="48f42-427">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-427">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-428">ReadItem</span></span>|<span data-ttu-id="48f42-429">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48f42-429">ReadWriteItem</span></span>|
|[<span data-ttu-id="48f42-430">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-431">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-431">Read</span></span>|<span data-ttu-id="48f42-432">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-432">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="48f42-433">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="48f42-433">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="48f42-434">Возвращает или задает настраиваемые заголовки Интернета для сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-434">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="48f42-435">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="48f42-435">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-436">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-436">Type</span></span>

*   [<span data-ttu-id="48f42-437">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="48f42-437">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="48f42-438">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-438">Requirements</span></span>

|<span data-ttu-id="48f42-439">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-439">Requirement</span></span>|<span data-ttu-id="48f42-440">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-441">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-442">1.8</span><span class="sxs-lookup"><span data-stu-id="48f42-442">1.8</span></span>|
|[<span data-ttu-id="48f42-443">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-444">ReadItem</span></span>|
|[<span data-ttu-id="48f42-445">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-446">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-446">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-447">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-447">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="48f42-448">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="48f42-448">internetMessageId: String</span></span>

<span data-ttu-id="48f42-p116">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="48f42-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-451">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-451">Type</span></span>

*   <span data-ttu-id="48f42-452">String</span><span class="sxs-lookup"><span data-stu-id="48f42-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-453">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-453">Requirements</span></span>

|<span data-ttu-id="48f42-454">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-454">Requirement</span></span>|<span data-ttu-id="48f42-455">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-456">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-457">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-457">1.0</span></span>|
|[<span data-ttu-id="48f42-458">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-459">ReadItem</span></span>|
|[<span data-ttu-id="48f42-460">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-461">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-462">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-462">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="48f42-463">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="48f42-463">itemClass: String</span></span>

<span data-ttu-id="48f42-p117">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="48f42-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="48f42-p118">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="48f42-468">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-468">Type</span></span>|<span data-ttu-id="48f42-469">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-469">Description</span></span>|<span data-ttu-id="48f42-470">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="48f42-470">item class</span></span>|
|---|---|---|
|<span data-ttu-id="48f42-471">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="48f42-471">Appointment items</span></span>|<span data-ttu-id="48f42-472">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="48f42-472">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="48f42-473">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="48f42-473">Message items</span></span>|<span data-ttu-id="48f42-474">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-474">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="48f42-475">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="48f42-475">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-476">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-476">Type</span></span>

*   <span data-ttu-id="48f42-477">String</span><span class="sxs-lookup"><span data-stu-id="48f42-477">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-478">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-478">Requirements</span></span>

|<span data-ttu-id="48f42-479">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-479">Requirement</span></span>|<span data-ttu-id="48f42-480">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-481">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-482">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-482">1.0</span></span>|
|[<span data-ttu-id="48f42-483">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-484">ReadItem</span></span>|
|[<span data-ttu-id="48f42-485">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-486">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-487">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-487">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="48f42-488">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="48f42-488">(nullable) itemId: String</span></span>

<span data-ttu-id="48f42-p119">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="48f42-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-491">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="48f42-491">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="48f42-492">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="48f42-492">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="48f42-493">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="48f42-493">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="48f42-494">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="48f42-494">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="48f42-p121">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-497">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-497">Type</span></span>

*   <span data-ttu-id="48f42-498">String</span><span class="sxs-lookup"><span data-stu-id="48f42-498">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-499">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-499">Requirements</span></span>

|<span data-ttu-id="48f42-500">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-500">Requirement</span></span>|<span data-ttu-id="48f42-501">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-502">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-503">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-503">1.0</span></span>|
|[<span data-ttu-id="48f42-504">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-504">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-505">ReadItem</span></span>|
|[<span data-ttu-id="48f42-506">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-506">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-507">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-508">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-508">Example</span></span>

<span data-ttu-id="48f42-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="48f42-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="48f42-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="48f42-512">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="48f42-512">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="48f42-513">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="48f42-513">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-514">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-514">Type</span></span>

*   [<span data-ttu-id="48f42-515">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="48f42-515">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="48f42-516">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-516">Requirements</span></span>

|<span data-ttu-id="48f42-517">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-517">Requirement</span></span>|<span data-ttu-id="48f42-518">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-519">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-520">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-520">1.0</span></span>|
|[<span data-ttu-id="48f42-521">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-522">ReadItem</span></span>|
|[<span data-ttu-id="48f42-523">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-524">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-524">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-525">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-525">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="48f42-526">location: String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="48f42-526">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="48f42-527">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-527">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-528">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-528">Read mode</span></span>

<span data-ttu-id="48f42-529">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-529">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-530">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-530">Compose mode</span></span>

<span data-ttu-id="48f42-531">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-531">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48f42-532">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-532">Type</span></span>

*   <span data-ttu-id="48f42-533">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="48f42-533">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-534">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-534">Requirements</span></span>

|<span data-ttu-id="48f42-535">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-535">Requirement</span></span>|<span data-ttu-id="48f42-536">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-537">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-538">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-538">1.0</span></span>|
|[<span data-ttu-id="48f42-539">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-540">ReadItem</span></span>|
|[<span data-ttu-id="48f42-541">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-542">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-542">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="48f42-543">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="48f42-543">normalizedSubject: String</span></span>

<span data-ttu-id="48f42-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="48f42-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="48f42-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="48f42-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-548">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-548">Type</span></span>

*   <span data-ttu-id="48f42-549">String</span><span class="sxs-lookup"><span data-stu-id="48f42-549">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-550">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-550">Requirements</span></span>

|<span data-ttu-id="48f42-551">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-551">Requirement</span></span>|<span data-ttu-id="48f42-552">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-553">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-554">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-554">1.0</span></span>|
|[<span data-ttu-id="48f42-555">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-556">ReadItem</span></span>|
|[<span data-ttu-id="48f42-557">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-558">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-558">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-559">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-559">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="48f42-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="48f42-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="48f42-561">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-561">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-562">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-562">Type</span></span>

*   [<span data-ttu-id="48f42-563">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="48f42-563">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="48f42-564">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-564">Requirements</span></span>

|<span data-ttu-id="48f42-565">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-565">Requirement</span></span>|<span data-ttu-id="48f42-566">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-567">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-568">1.3</span><span class="sxs-lookup"><span data-stu-id="48f42-568">1.3</span></span>|
|[<span data-ttu-id="48f42-569">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-570">ReadItem</span></span>|
|[<span data-ttu-id="48f42-571">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-572">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-572">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-573">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-573">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48f42-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48f42-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48f42-575">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="48f42-575">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="48f42-576">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-576">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-577">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-577">Read mode</span></span>

<span data-ttu-id="48f42-578">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-578">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="48f42-579">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-579">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48f42-580">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="48f42-580">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-581">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-581">Compose mode</span></span>

<span data-ttu-id="48f42-582">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-582">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="48f42-583">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-583">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48f42-584">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="48f42-584">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48f42-585">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-585">Get 500 members maximum.</span></span>
- <span data-ttu-id="48f42-586">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-586">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48f42-587">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-587">Type</span></span>

*   <span data-ttu-id="48f42-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48f42-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-589">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-589">Requirements</span></span>

|<span data-ttu-id="48f42-590">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-590">Requirement</span></span>|<span data-ttu-id="48f42-591">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-592">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-593">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-593">1.0</span></span>|
|[<span data-ttu-id="48f42-594">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-595">ReadItem</span></span>|
|[<span data-ttu-id="48f42-596">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-597">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-597">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="48f42-598">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Организатор](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="48f42-598">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="48f42-599">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-599">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-600">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-600">Read mode</span></span>

<span data-ttu-id="48f42-601">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-601">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-602">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-602">Compose mode</span></span>

<span data-ttu-id="48f42-603">`organizer` Свойство возвращает объект [организатора](/javascript/api/outlook/office.organizer) , который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="48f42-603">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="48f42-604">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-604">Type</span></span>

*   <span data-ttu-id="48f42-605">[](/javascript/api/outlook/office.emailaddressdetails) | [Организатор](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="48f42-605">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-606">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-606">Requirements</span></span>

|<span data-ttu-id="48f42-607">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-607">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="48f42-608">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-609">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-609">1.0</span></span>|<span data-ttu-id="48f42-610">1.7</span><span class="sxs-lookup"><span data-stu-id="48f42-610">1.7</span></span>|
|[<span data-ttu-id="48f42-611">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-611">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-612">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-612">ReadItem</span></span>|<span data-ttu-id="48f42-613">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48f42-613">ReadWriteItem</span></span>|
|[<span data-ttu-id="48f42-614">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-615">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-615">Read</span></span>|<span data-ttu-id="48f42-616">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-616">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="48f42-617">(Nullable) повторение: [повторение](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="48f42-617">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="48f42-618">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-618">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="48f42-619">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="48f42-619">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="48f42-620">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-620">Read and compose modes for appointment items.</span></span> <span data-ttu-id="48f42-621">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-621">Read mode for meeting request items.</span></span>

<span data-ttu-id="48f42-622">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="48f42-622">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="48f42-623">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="48f42-623">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="48f42-624">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-624">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="48f42-625">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="48f42-625">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="48f42-626">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="48f42-626">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-627">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-627">Read mode</span></span>

<span data-ttu-id="48f42-628">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-628">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="48f42-629">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-629">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-630">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-630">Compose mode</span></span>

<span data-ttu-id="48f42-631">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-631">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="48f42-632">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="48f42-632">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="48f42-633">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-633">Type</span></span>

* [<span data-ttu-id="48f42-634">Повторения</span><span class="sxs-lookup"><span data-stu-id="48f42-634">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="48f42-635">Requirement</span><span class="sxs-lookup"><span data-stu-id="48f42-635">Requirement</span></span>|<span data-ttu-id="48f42-636">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-637">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-638">1.7</span><span class="sxs-lookup"><span data-stu-id="48f42-638">1.7</span></span>|
|[<span data-ttu-id="48f42-639">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-640">ReadItem</span></span>|
|[<span data-ttu-id="48f42-641">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-642">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48f42-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48f42-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48f42-644">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="48f42-644">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="48f42-645">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-645">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-646">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-646">Read mode</span></span>

<span data-ttu-id="48f42-647">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-647">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="48f42-648">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-648">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48f42-649">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="48f42-649">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-650">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-650">Compose mode</span></span>

<span data-ttu-id="48f42-651">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-651">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="48f42-652">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-652">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48f42-653">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="48f42-653">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48f42-654">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-654">Get 500 members maximum.</span></span>
- <span data-ttu-id="48f42-655">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-655">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="48f42-656">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-656">Type</span></span>

*   <span data-ttu-id="48f42-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48f42-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-658">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-658">Requirements</span></span>

|<span data-ttu-id="48f42-659">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-659">Requirement</span></span>|<span data-ttu-id="48f42-660">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-661">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-662">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-662">1.0</span></span>|
|[<span data-ttu-id="48f42-663">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-664">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-664">ReadItem</span></span>|
|[<span data-ttu-id="48f42-665">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-666">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-666">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="48f42-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="48f42-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="48f42-p135">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="48f42-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="48f42-p136">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="48f42-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-672">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="48f42-672">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-673">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-673">Type</span></span>

*   [<span data-ttu-id="48f42-674">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="48f42-674">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="48f42-675">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-675">Requirements</span></span>

|<span data-ttu-id="48f42-676">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-676">Requirement</span></span>|<span data-ttu-id="48f42-677">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-677">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-678">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-678">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-679">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-679">1.0</span></span>|
|[<span data-ttu-id="48f42-680">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-680">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-681">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-681">ReadItem</span></span>|
|[<span data-ttu-id="48f42-682">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-682">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-683">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-683">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-684">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-684">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="48f42-685">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="48f42-685">(nullable) seriesId: String</span></span>

<span data-ttu-id="48f42-686">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="48f42-686">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="48f42-687">В Outlook в Интернете и на настольных клиентах `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="48f42-687">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="48f42-688">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-688">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-689">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="48f42-689">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="48f42-690">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="48f42-690">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="48f42-691">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="48f42-691">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="48f42-692">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="48f42-692">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="48f42-693">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-693">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="48f42-694">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-694">Type</span></span>

* <span data-ttu-id="48f42-695">String</span><span class="sxs-lookup"><span data-stu-id="48f42-695">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-696">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-696">Requirements</span></span>

|<span data-ttu-id="48f42-697">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-697">Requirement</span></span>|<span data-ttu-id="48f42-698">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-699">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-700">1.7</span><span class="sxs-lookup"><span data-stu-id="48f42-700">1.7</span></span>|
|[<span data-ttu-id="48f42-701">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-702">ReadItem</span></span>|
|[<span data-ttu-id="48f42-703">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-704">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-704">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-705">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-705">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="48f42-706">start: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="48f42-706">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="48f42-707">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-707">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="48f42-p139">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="48f42-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-710">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-710">Read mode</span></span>

<span data-ttu-id="48f42-711">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="48f42-711">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-712">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-712">Compose mode</span></span>

<span data-ttu-id="48f42-713">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="48f42-713">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="48f42-714">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="48f42-714">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="48f42-715">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="48f42-715">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="48f42-716">Type</span><span class="sxs-lookup"><span data-stu-id="48f42-716">Type</span></span>

*   <span data-ttu-id="48f42-717">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="48f42-717">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-718">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-718">Requirements</span></span>

|<span data-ttu-id="48f42-719">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-719">Requirement</span></span>|<span data-ttu-id="48f42-720">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-720">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-721">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-721">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-722">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-722">1.0</span></span>|
|[<span data-ttu-id="48f42-723">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-723">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-724">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-724">ReadItem</span></span>|
|[<span data-ttu-id="48f42-725">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-725">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-726">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-726">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="48f42-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="48f42-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="48f42-728">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-728">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="48f42-729">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="48f42-729">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-730">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-730">Read mode</span></span>

<span data-ttu-id="48f42-p140">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="48f42-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="48f42-733">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="48f42-733">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-734">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-734">Compose mode</span></span>
<span data-ttu-id="48f42-735">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="48f42-735">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="48f42-736">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-736">Type</span></span>

*   <span data-ttu-id="48f42-737">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="48f42-737">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-738">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-738">Requirements</span></span>

|<span data-ttu-id="48f42-739">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-739">Requirement</span></span>|<span data-ttu-id="48f42-740">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-741">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-742">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-742">1.0</span></span>|
|[<span data-ttu-id="48f42-743">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-743">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-744">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-744">ReadItem</span></span>|
|[<span data-ttu-id="48f42-745">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-745">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-746">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-746">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="48f42-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48f42-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="48f42-748">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-748">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="48f42-749">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-749">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="48f42-750">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="48f42-750">Read mode</span></span>

<span data-ttu-id="48f42-751">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-751">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="48f42-752">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-752">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48f42-753">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="48f42-753">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="48f42-754">Режим создания</span><span class="sxs-lookup"><span data-stu-id="48f42-754">Compose mode</span></span>

<span data-ttu-id="48f42-755">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-755">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="48f42-756">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-756">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="48f42-757">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="48f42-757">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="48f42-758">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-758">Get 500 members maximum.</span></span>
- <span data-ttu-id="48f42-759">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="48f42-759">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="48f42-760">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-760">Type</span></span>

*   <span data-ttu-id="48f42-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="48f42-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-762">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-762">Requirements</span></span>

|<span data-ttu-id="48f42-763">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-763">Requirement</span></span>|<span data-ttu-id="48f42-764">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-765">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-766">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-766">1.0</span></span>|
|[<span data-ttu-id="48f42-767">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-767">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-768">ReadItem</span></span>|
|[<span data-ttu-id="48f42-769">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-769">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-770">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-770">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="48f42-771">Методы</span><span class="sxs-lookup"><span data-stu-id="48f42-771">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="48f42-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48f42-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="48f42-773">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="48f42-773">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="48f42-774">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="48f42-774">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="48f42-775">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="48f42-775">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-776">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-776">Parameters</span></span>
|<span data-ttu-id="48f42-777">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-777">Name</span></span>|<span data-ttu-id="48f42-778">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-778">Type</span></span>|<span data-ttu-id="48f42-779">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-779">Attributes</span></span>|<span data-ttu-id="48f42-780">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-780">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="48f42-781">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-781">String</span></span>||<span data-ttu-id="48f42-p144">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="48f42-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="48f42-784">String</span><span class="sxs-lookup"><span data-stu-id="48f42-784">String</span></span>||<span data-ttu-id="48f42-p145">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="48f42-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="48f42-787">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-787">Object</span></span>|<span data-ttu-id="48f42-788">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-788">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-789">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-789">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-790">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-790">Object</span></span>|<span data-ttu-id="48f42-791">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-791">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-792">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="48f42-792">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="48f42-793">Boolean</span><span class="sxs-lookup"><span data-stu-id="48f42-793">Boolean</span></span>|<span data-ttu-id="48f42-794">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-794">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-795">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="48f42-795">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="48f42-796">function</span><span class="sxs-lookup"><span data-stu-id="48f42-796">function</span></span>|<span data-ttu-id="48f42-797">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-797">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-798">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-798">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48f42-799">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48f42-799">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="48f42-800">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="48f42-800">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48f42-801">Ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-801">Errors</span></span>

|<span data-ttu-id="48f42-802">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-802">Error code</span></span>|<span data-ttu-id="48f42-803">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-803">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="48f42-804">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="48f42-804">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="48f42-805">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="48f42-805">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="48f42-806">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="48f42-806">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-807">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-807">Requirements</span></span>

|<span data-ttu-id="48f42-808">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-808">Requirement</span></span>|<span data-ttu-id="48f42-809">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-809">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-810">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-810">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-811">1.1</span><span class="sxs-lookup"><span data-stu-id="48f42-811">1.1</span></span>|
|[<span data-ttu-id="48f42-812">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-812">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-813">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48f42-813">ReadWriteItem</span></span>|
|[<span data-ttu-id="48f42-814">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-814">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-815">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-815">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="48f42-816">Примеры</span><span class="sxs-lookup"><span data-stu-id="48f42-816">Examples</span></span>

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

<span data-ttu-id="48f42-817">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-817">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="48f42-818">addFileAttachmentFromBase64Async (base64File, Аттачментнаме, [параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="48f42-818">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="48f42-819">Добавляет файл из кодировки Base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="48f42-819">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="48f42-820">`addFileAttachmentFromBase64Async` Метод передает файл из кодировки Base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="48f42-820">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="48f42-821">Этот метод возвращает идентификатор вложения в объекте AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="48f42-821">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="48f42-822">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="48f42-822">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-823">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-823">Parameters</span></span>

|<span data-ttu-id="48f42-824">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-824">Name</span></span>|<span data-ttu-id="48f42-825">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-825">Type</span></span>|<span data-ttu-id="48f42-826">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-826">Attributes</span></span>|<span data-ttu-id="48f42-827">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-827">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="48f42-828">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-828">String</span></span>||<span data-ttu-id="48f42-829">Содержимое изображения или файла в кодировке Base64, которое добавляется в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="48f42-829">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="48f42-830">String</span><span class="sxs-lookup"><span data-stu-id="48f42-830">String</span></span>||<span data-ttu-id="48f42-p147">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="48f42-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="48f42-833">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-833">Object</span></span>|<span data-ttu-id="48f42-834">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-834">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-835">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-835">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-836">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-836">Object</span></span>|<span data-ttu-id="48f42-837">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-837">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-838">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="48f42-838">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="48f42-839">Boolean</span><span class="sxs-lookup"><span data-stu-id="48f42-839">Boolean</span></span>|<span data-ttu-id="48f42-840">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-840">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-841">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="48f42-841">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="48f42-842">function</span><span class="sxs-lookup"><span data-stu-id="48f42-842">function</span></span>|<span data-ttu-id="48f42-843">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-843">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-844">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-844">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48f42-845">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48f42-845">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="48f42-846">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="48f42-846">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48f42-847">Ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-847">Errors</span></span>

|<span data-ttu-id="48f42-848">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-848">Error code</span></span>|<span data-ttu-id="48f42-849">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-849">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="48f42-850">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="48f42-850">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="48f42-851">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="48f42-851">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="48f42-852">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="48f42-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-853">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-853">Requirements</span></span>

|<span data-ttu-id="48f42-854">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-854">Requirement</span></span>|<span data-ttu-id="48f42-855">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-856">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-857">1.8</span><span class="sxs-lookup"><span data-stu-id="48f42-857">1.8</span></span>|
|[<span data-ttu-id="48f42-858">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-858">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48f42-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="48f42-860">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-860">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-861">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-861">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="48f42-862">Примеры</span><span class="sxs-lookup"><span data-stu-id="48f42-862">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="48f42-863">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48f42-863">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="48f42-864">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="48f42-864">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="48f42-865">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="48f42-865">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-866">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-866">Parameters</span></span>

| <span data-ttu-id="48f42-867">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-867">Name</span></span> | <span data-ttu-id="48f42-868">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-868">Type</span></span> | <span data-ttu-id="48f42-869">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-869">Attributes</span></span> | <span data-ttu-id="48f42-870">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-870">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="48f42-871">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="48f42-871">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="48f42-872">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="48f42-872">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="48f42-873">Function</span><span class="sxs-lookup"><span data-stu-id="48f42-873">Function</span></span> || <span data-ttu-id="48f42-p148">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="48f42-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="48f42-877">Объект</span><span class="sxs-lookup"><span data-stu-id="48f42-877">Object</span></span> | <span data-ttu-id="48f42-878">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-878">&lt;optional&gt;</span></span> | <span data-ttu-id="48f42-879">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-879">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="48f42-880">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-880">Object</span></span> | <span data-ttu-id="48f42-881">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-881">&lt;optional&gt;</span></span> | <span data-ttu-id="48f42-882">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-882">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="48f42-883">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-883">function</span></span>| <span data-ttu-id="48f42-884">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-884">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-885">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-885">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-886">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-886">Requirements</span></span>

|<span data-ttu-id="48f42-887">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-887">Requirement</span></span>| <span data-ttu-id="48f42-888">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-888">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-889">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-889">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48f42-890">1.7</span><span class="sxs-lookup"><span data-stu-id="48f42-890">1.7</span></span> |
|[<span data-ttu-id="48f42-891">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-891">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48f42-892">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-892">ReadItem</span></span> |
|[<span data-ttu-id="48f42-893">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-893">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48f42-894">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-894">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="48f42-895">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-895">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="48f42-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48f42-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="48f42-897">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="48f42-897">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="48f42-p149">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="48f42-901">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="48f42-901">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="48f42-902">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="48f42-902">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-903">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-903">Parameters</span></span>

|<span data-ttu-id="48f42-904">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-904">Name</span></span>|<span data-ttu-id="48f42-905">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-905">Type</span></span>|<span data-ttu-id="48f42-906">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-906">Attributes</span></span>|<span data-ttu-id="48f42-907">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-907">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="48f42-908">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-908">String</span></span>||<span data-ttu-id="48f42-p150">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="48f42-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="48f42-911">String</span><span class="sxs-lookup"><span data-stu-id="48f42-911">String</span></span>||<span data-ttu-id="48f42-912">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-912">The subject of the item to be attached.</span></span> <span data-ttu-id="48f42-913">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="48f42-913">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="48f42-914">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-914">Object</span></span>|<span data-ttu-id="48f42-915">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-915">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-916">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-916">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-917">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-917">Object</span></span>|<span data-ttu-id="48f42-918">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-918">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-919">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-919">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-920">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-920">function</span></span>|<span data-ttu-id="48f42-921">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-921">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-922">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48f42-923">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48f42-923">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="48f42-924">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="48f42-924">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48f42-925">Ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-925">Errors</span></span>

|<span data-ttu-id="48f42-926">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-926">Error code</span></span>|<span data-ttu-id="48f42-927">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-927">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="48f42-928">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="48f42-928">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-929">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-929">Requirements</span></span>

|<span data-ttu-id="48f42-930">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-930">Requirement</span></span>|<span data-ttu-id="48f42-931">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-931">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-932">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-932">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-933">1.1</span><span class="sxs-lookup"><span data-stu-id="48f42-933">1.1</span></span>|
|[<span data-ttu-id="48f42-934">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-934">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-935">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48f42-935">ReadWriteItem</span></span>|
|[<span data-ttu-id="48f42-936">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-936">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-937">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-937">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-938">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-938">Example</span></span>

<span data-ttu-id="48f42-939">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="48f42-939">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="48f42-940">close()</span><span class="sxs-lookup"><span data-stu-id="48f42-940">close()</span></span>

<span data-ttu-id="48f42-941">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="48f42-941">Closes the current item that is being composed.</span></span>

<span data-ttu-id="48f42-p152">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="48f42-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-944">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="48f42-944">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="48f42-945">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="48f42-945">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-946">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-946">Requirements</span></span>

|<span data-ttu-id="48f42-947">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-947">Requirement</span></span>|<span data-ttu-id="48f42-948">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-948">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-949">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-949">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-950">1.3</span><span class="sxs-lookup"><span data-stu-id="48f42-950">1.3</span></span>|
|[<span data-ttu-id="48f42-951">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-951">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-952">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="48f42-952">Restricted</span></span>|
|[<span data-ttu-id="48f42-953">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-953">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-954">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-954">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="48f42-955">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="48f42-955">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="48f42-956">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-956">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-957">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48f42-958">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="48f42-958">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="48f42-959">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="48f42-959">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="48f42-p153">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="48f42-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-963">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-963">Parameters</span></span>

|<span data-ttu-id="48f42-964">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-964">Name</span></span>|<span data-ttu-id="48f42-965">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-965">Type</span></span>|<span data-ttu-id="48f42-966">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-966">Attributes</span></span>|<span data-ttu-id="48f42-967">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-967">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="48f42-968">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="48f42-968">String &#124; Object</span></span>||<span data-ttu-id="48f42-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="48f42-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="48f42-971">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="48f42-971">**OR**</span></span><br/><span data-ttu-id="48f42-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="48f42-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="48f42-974">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-974">String</span></span>|<span data-ttu-id="48f42-975">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-975">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="48f42-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="48f42-978">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-978">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="48f42-979">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-979">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-980">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="48f42-980">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="48f42-981">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-981">String</span></span>||<span data-ttu-id="48f42-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="48f42-984">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-984">String</span></span>||<span data-ttu-id="48f42-985">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="48f42-985">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="48f42-986">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-986">String</span></span>||<span data-ttu-id="48f42-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="48f42-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="48f42-989">Логический</span><span class="sxs-lookup"><span data-stu-id="48f42-989">Boolean</span></span>||<span data-ttu-id="48f42-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="48f42-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="48f42-992">String</span><span class="sxs-lookup"><span data-stu-id="48f42-992">String</span></span>||<span data-ttu-id="48f42-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="48f42-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="48f42-996">function</span><span class="sxs-lookup"><span data-stu-id="48f42-996">function</span></span>|<span data-ttu-id="48f42-997">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-997">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-998">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-998">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-999">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-999">Requirements</span></span>

|<span data-ttu-id="48f42-1000">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1000">Requirement</span></span>|<span data-ttu-id="48f42-1001">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1002">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-1003">1.0</span></span>|
|[<span data-ttu-id="48f42-1004">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1005">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1006">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1007">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1007">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="48f42-1008">Примеры</span><span class="sxs-lookup"><span data-stu-id="48f42-1008">Examples</span></span>

<span data-ttu-id="48f42-1009">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1009">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="48f42-1010">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-1010">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="48f42-1011">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-1011">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="48f42-1012">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="48f42-1012">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="48f42-1013">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="48f42-1013">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="48f42-1014">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="48f42-1014">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="48f42-1015">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="48f42-1015">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="48f42-1016">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-1016">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1017">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-1017">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48f42-1018">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="48f42-1018">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="48f42-1019">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="48f42-1019">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="48f42-p161">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="48f42-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1023">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1023">Parameters</span></span>

|<span data-ttu-id="48f42-1024">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1024">Name</span></span>|<span data-ttu-id="48f42-1025">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1025">Type</span></span>|<span data-ttu-id="48f42-1026">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1026">Attributes</span></span>|<span data-ttu-id="48f42-1027">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1027">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="48f42-1028">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1028">String &#124; Object</span></span>||<span data-ttu-id="48f42-p162">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="48f42-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="48f42-1031">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="48f42-1031">**OR**</span></span><br/><span data-ttu-id="48f42-p163">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="48f42-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="48f42-1034">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-1034">String</span></span>|<span data-ttu-id="48f42-1035">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-p164">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="48f42-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="48f42-1038">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1038">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="48f42-1039">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1040">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="48f42-1040">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="48f42-1041">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-1041">String</span></span>||<span data-ttu-id="48f42-p165">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="48f42-1044">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-1044">String</span></span>||<span data-ttu-id="48f42-1045">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="48f42-1045">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="48f42-1046">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-1046">String</span></span>||<span data-ttu-id="48f42-p166">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="48f42-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="48f42-1049">Логический</span><span class="sxs-lookup"><span data-stu-id="48f42-1049">Boolean</span></span>||<span data-ttu-id="48f42-p167">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="48f42-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="48f42-1052">String</span><span class="sxs-lookup"><span data-stu-id="48f42-1052">String</span></span>||<span data-ttu-id="48f42-p168">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="48f42-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="48f42-1056">function</span><span class="sxs-lookup"><span data-stu-id="48f42-1056">function</span></span>|<span data-ttu-id="48f42-1057">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1058">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1059">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1059">Requirements</span></span>

|<span data-ttu-id="48f42-1060">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1060">Requirement</span></span>|<span data-ttu-id="48f42-1061">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1061">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1062">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-1062">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1063">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-1063">1.0</span></span>|
|[<span data-ttu-id="48f42-1064">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1064">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1065">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1065">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1066">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1066">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1067">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1067">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="48f42-1068">Примеры</span><span class="sxs-lookup"><span data-stu-id="48f42-1068">Examples</span></span>

<span data-ttu-id="48f42-1069">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1069">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="48f42-1070">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-1070">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="48f42-1071">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-1071">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="48f42-1072">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="48f42-1072">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="48f42-1073">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="48f42-1073">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="48f42-1074">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="48f42-1074">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="48f42-1075">Жеталлинтернесеадерсасинк ([параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="48f42-1075">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="48f42-1076">Получает все заголовки Интернета для сообщения в виде строки.</span><span class="sxs-lookup"><span data-stu-id="48f42-1076">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="48f42-1077">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="48f42-1077">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1078">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1078">Parameters</span></span>

|<span data-ttu-id="48f42-1079">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1079">Name</span></span>|<span data-ttu-id="48f42-1080">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1080">Type</span></span>|<span data-ttu-id="48f42-1081">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1081">Attributes</span></span>|<span data-ttu-id="48f42-1082">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1082">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48f42-1083">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1083">Object</span></span>|<span data-ttu-id="48f42-1084">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1084">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1085">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1085">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1086">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1086">Object</span></span>|<span data-ttu-id="48f42-1087">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1088">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1088">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-1089">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1089">function</span></span>|<span data-ttu-id="48f42-1090">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1091">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1091">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="48f42-1092">В случае успешного выполнения данные заголовков Интернета предоставляются в свойстве asyncResult. Value в виде String.</span><span class="sxs-lookup"><span data-stu-id="48f42-1092">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="48f42-1093">Сведения о форматировании возвращаемого строкового значения приведены в [RFC 2183](https://tools.ietf.org/html/rfc2183) .</span><span class="sxs-lookup"><span data-stu-id="48f42-1093">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="48f42-1094">Если происходит сбой вызова, свойство asyncResult. Error будет содержать код ошибки с причиной сбоя.</span><span class="sxs-lookup"><span data-stu-id="48f42-1094">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1095">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1095">Requirements</span></span>

|<span data-ttu-id="48f42-1096">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1096">Requirement</span></span>|<span data-ttu-id="48f42-1097">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1098">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1099">1.8</span><span class="sxs-lookup"><span data-stu-id="48f42-1099">1.8</span></span>|
|[<span data-ttu-id="48f42-1100">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1101">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1102">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1103">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1103">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1104">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1104">Returns:</span></span>

<span data-ttu-id="48f42-1105">Данные заголовков Интернета в виде строки, отформатированной в соответствии со [спецификацией RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="48f42-1105">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="48f42-1106">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="48f42-1106">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="48f42-1107">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1107">Example</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="48f42-1108">Жетаттачментконтентасинк (attachmentId, [параметры], [callback]) → [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="48f42-1108">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="48f42-1109">Получает указанное вложение из сообщения или встречи и возвращает его в виде `AttachmentContent` объекта.</span><span class="sxs-lookup"><span data-stu-id="48f42-1109">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="48f42-1110">`getAttachmentContentAsync` Метод получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-1110">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="48f42-1111">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, когда Аттачментидс был получен с помощью вызова `getAttachmentsAsync` или. `item.attachments`</span><span class="sxs-lookup"><span data-stu-id="48f42-1111">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="48f42-1112">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="48f42-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="48f42-1113">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="48f42-1113">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1114">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1114">Parameters</span></span>

|<span data-ttu-id="48f42-1115">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1115">Name</span></span>|<span data-ttu-id="48f42-1116">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1116">Type</span></span>|<span data-ttu-id="48f42-1117">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1117">Attributes</span></span>|<span data-ttu-id="48f42-1118">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="48f42-1119">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-1119">String</span></span>||<span data-ttu-id="48f42-1120">Идентификатор вложения, которое требуется получить.</span><span class="sxs-lookup"><span data-stu-id="48f42-1120">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="48f42-1121">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1121">Object</span></span>|<span data-ttu-id="48f42-1122">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1123">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1124">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1124">Object</span></span>|<span data-ttu-id="48f42-1125">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1126">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-1127">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1127">function</span></span>|<span data-ttu-id="48f42-1128">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1129">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1130">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1130">Requirements</span></span>

|<span data-ttu-id="48f42-1131">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1131">Requirement</span></span>|<span data-ttu-id="48f42-1132">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1133">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1134">1.8</span><span class="sxs-lookup"><span data-stu-id="48f42-1134">1.8</span></span>|
|[<span data-ttu-id="48f42-1135">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1136">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1137">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1138">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1138">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1139">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1139">Returns:</span></span>

<span data-ttu-id="48f42-1140">Тип: [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="48f42-1140">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="48f42-1141">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1141">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="48f42-1142">Жетаттачментсасинк ([параметры], [обратный вызов]) → массив. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="48f42-1142">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="48f42-1143">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="48f42-1143">Gets the item's attachments as an array.</span></span> <span data-ttu-id="48f42-1144">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="48f42-1144">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1145">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1145">Parameters</span></span>

|<span data-ttu-id="48f42-1146">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1146">Name</span></span>|<span data-ttu-id="48f42-1147">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1147">Type</span></span>|<span data-ttu-id="48f42-1148">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1148">Attributes</span></span>|<span data-ttu-id="48f42-1149">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1149">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48f42-1150">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1150">Object</span></span>|<span data-ttu-id="48f42-1151">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1151">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1152">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1152">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1153">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1153">Object</span></span>|<span data-ttu-id="48f42-1154">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1155">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1155">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-1156">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1156">function</span></span>|<span data-ttu-id="48f42-1157">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1158">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1158">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1159">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1159">Requirements</span></span>

|<span data-ttu-id="48f42-1160">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1160">Requirement</span></span>|<span data-ttu-id="48f42-1161">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1162">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1163">1.8</span><span class="sxs-lookup"><span data-stu-id="48f42-1163">1.8</span></span>|
|[<span data-ttu-id="48f42-1164">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1165">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1166">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1167">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-1167">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1168">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1168">Returns:</span></span>

<span data-ttu-id="48f42-1169">Тип: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="48f42-1169">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="48f42-1170">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1170">Example</span></span>

<span data-ttu-id="48f42-1171">В приведенном ниже примере создается строка HTML со сведениями обо всех вложениях в текущем элементе.</span><span class="sxs-lookup"><span data-stu-id="48f42-1171">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="48f42-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="48f42-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="48f42-1173">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-1173">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1174">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-1174">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-1175">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1175">Requirements</span></span>

|<span data-ttu-id="48f42-1176">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1176">Requirement</span></span>|<span data-ttu-id="48f42-1177">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1178">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1179">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-1179">1.0</span></span>|
|[<span data-ttu-id="48f42-1180">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1181">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1182">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1183">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1183">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1184">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1184">Returns:</span></span>

<span data-ttu-id="48f42-1185">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="48f42-1185">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="48f42-1186">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1186">Example</span></span>

<span data-ttu-id="48f42-1187">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-1187">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="48f42-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="48f42-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="48f42-1189">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-1189">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1190">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-1190">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1191">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1191">Parameters</span></span>

|<span data-ttu-id="48f42-1192">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1192">Name</span></span>|<span data-ttu-id="48f42-1193">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1193">Type</span></span>|<span data-ttu-id="48f42-1194">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1194">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="48f42-1195">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="48f42-1195">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="48f42-1196">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="48f42-1196">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1197">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1197">Requirements</span></span>

|<span data-ttu-id="48f42-1198">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1198">Requirement</span></span>|<span data-ttu-id="48f42-1199">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1199">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1200">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-1200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1201">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-1201">1.0</span></span>|
|[<span data-ttu-id="48f42-1202">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1203">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="48f42-1203">Restricted</span></span>|
|[<span data-ttu-id="48f42-1204">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1205">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1205">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1206">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1206">Returns:</span></span>

<span data-ttu-id="48f42-1207">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="48f42-1207">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="48f42-1208">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="48f42-1208">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="48f42-1209">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1209">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="48f42-1210">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="48f42-1210">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="48f42-1211">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="48f42-1211">Value of `entityType`</span></span>|<span data-ttu-id="48f42-1212">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="48f42-1212">Type of objects in returned array</span></span>|<span data-ttu-id="48f42-1213">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1213">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="48f42-1214">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-1214">String</span></span>|<span data-ttu-id="48f42-1215">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="48f42-1215">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="48f42-1216">Contact</span><span class="sxs-lookup"><span data-stu-id="48f42-1216">Contact</span></span>|<span data-ttu-id="48f42-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="48f42-1217">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="48f42-1218">String</span><span class="sxs-lookup"><span data-stu-id="48f42-1218">String</span></span>|<span data-ttu-id="48f42-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="48f42-1219">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="48f42-1220">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="48f42-1220">MeetingSuggestion</span></span>|<span data-ttu-id="48f42-1221">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="48f42-1221">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="48f42-1222">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="48f42-1222">PhoneNumber</span></span>|<span data-ttu-id="48f42-1223">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="48f42-1223">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="48f42-1224">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="48f42-1224">TaskSuggestion</span></span>|<span data-ttu-id="48f42-1225">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="48f42-1225">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="48f42-1226">String</span><span class="sxs-lookup"><span data-stu-id="48f42-1226">String</span></span>|<span data-ttu-id="48f42-1227">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="48f42-1227">**Restricted**</span></span>|

<span data-ttu-id="48f42-1228">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="48f42-1228">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="48f42-1229">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1229">Example</span></span>

<span data-ttu-id="48f42-1230">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-1230">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="48f42-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="48f42-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="48f42-1232">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="48f42-1232">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1233">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-1233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48f42-1234">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1234">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1235">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1235">Parameters</span></span>

|<span data-ttu-id="48f42-1236">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1236">Name</span></span>|<span data-ttu-id="48f42-1237">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1237">Type</span></span>|<span data-ttu-id="48f42-1238">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1238">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="48f42-1239">Строка</span><span class="sxs-lookup"><span data-stu-id="48f42-1239">String</span></span>|<span data-ttu-id="48f42-1240">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="48f42-1240">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1241">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1241">Requirements</span></span>

|<span data-ttu-id="48f42-1242">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1242">Requirement</span></span>|<span data-ttu-id="48f42-1243">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1243">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1244">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-1244">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1245">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-1245">1.0</span></span>|
|[<span data-ttu-id="48f42-1246">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1246">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1247">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1247">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1248">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1248">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1249">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1249">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1250">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1250">Returns:</span></span>

<span data-ttu-id="48f42-p174">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="48f42-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="48f42-1253">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="48f42-1253">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="48f42-1254">getInitializationContextAsync ([параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="48f42-1254">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="48f42-1255">Получает данные инициализации, передаваемые при активации надстройки [сообщением с действиями](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="48f42-1255">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1256">Этот метод поддерживается только в Outlook 2016 или более поздней версии для Windows ("нажми и работай" более поздней версии, чем 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="48f42-1256">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1257">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1257">Parameters</span></span>

|<span data-ttu-id="48f42-1258">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1258">Name</span></span>|<span data-ttu-id="48f42-1259">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1259">Type</span></span>|<span data-ttu-id="48f42-1260">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1260">Attributes</span></span>|<span data-ttu-id="48f42-1261">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48f42-1262">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1262">Object</span></span>|<span data-ttu-id="48f42-1263">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1264">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1265">Объект</span><span class="sxs-lookup"><span data-stu-id="48f42-1265">Object</span></span>|<span data-ttu-id="48f42-1266">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1267">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-1268">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1268">function</span></span>|<span data-ttu-id="48f42-1269">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1269">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1270">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48f42-1271">При успешном выполнении данные инициализации предоставляются в `asyncResult.value` свойстве в виде строки.</span><span class="sxs-lookup"><span data-stu-id="48f42-1271">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="48f42-1272">Если `asyncResult` контекст инициализации отсутствует, объект будет содержать `Error` объект со `code` свойством, `9020` `name` для свойства которого задано значение. `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="48f42-1272">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1273">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1273">Requirements</span></span>

|<span data-ttu-id="48f42-1274">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1274">Requirement</span></span>|<span data-ttu-id="48f42-1275">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1275">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1276">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1277">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="48f42-1277">Preview</span></span>|
|[<span data-ttu-id="48f42-1278">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1279">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1280">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1281">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1281">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-1282">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1282">Example</span></span>

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="48f42-1283">Жетитемидасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="48f42-1283">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="48f42-1284">Асинхронно получает идентификатор сохраненного элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-1284">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="48f42-1285">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="48f42-1285">Compose mode only.</span></span>

<span data-ttu-id="48f42-1286">При вызове этот метод возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1286">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1287">Если надстройка вызывает `getItemIdAsync` элемент в режиме создания (например, чтобы получить доступ `itemId` к использованию с помощью EWS или REST API), имейте в виду, что если Outlook находится в режиме кэширования, может потребоваться некоторое время до синхронизации элемента с сервером.</span><span class="sxs-lookup"><span data-stu-id="48f42-1287">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="48f42-1288">Пока элемент не будет синхронизирован, он не `itemId` распознается и не будет использоваться, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="48f42-1288">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1289">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1289">Parameters</span></span>

|<span data-ttu-id="48f42-1290">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1290">Name</span></span>|<span data-ttu-id="48f42-1291">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1291">Type</span></span>|<span data-ttu-id="48f42-1292">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1292">Attributes</span></span>|<span data-ttu-id="48f42-1293">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1293">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48f42-1294">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1294">Object</span></span>|<span data-ttu-id="48f42-1295">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1295">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1296">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1296">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1297">Объект</span><span class="sxs-lookup"><span data-stu-id="48f42-1297">Object</span></span>|<span data-ttu-id="48f42-1298">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1299">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1299">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-1300">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1300">function</span></span>||<span data-ttu-id="48f42-1301">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1301">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48f42-1302">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1302">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48f42-1303">Ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-1303">Errors</span></span>

|<span data-ttu-id="48f42-1304">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-1304">Error code</span></span>|<span data-ttu-id="48f42-1305">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1305">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="48f42-1306">Идентификатор невозможно извлечь, пока не будет сохранен элемент.</span><span class="sxs-lookup"><span data-stu-id="48f42-1306">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1307">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1307">Requirements</span></span>

|<span data-ttu-id="48f42-1308">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1308">Requirement</span></span>|<span data-ttu-id="48f42-1309">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1309">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1310">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1310">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1311">1.8</span><span class="sxs-lookup"><span data-stu-id="48f42-1311">1.8</span></span>|
|[<span data-ttu-id="48f42-1312">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1312">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1313">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1314">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1314">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1315">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-1315">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="48f42-1316">Примеры</span><span class="sxs-lookup"><span data-stu-id="48f42-1316">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="48f42-1317">В следующем примере показана структура `result` параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1317">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="48f42-1318">`value` Свойство содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-1318">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="48f42-1319">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="48f42-1319">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="48f42-1320">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="48f42-1320">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1321">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-1321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48f42-p178">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="48f42-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="48f42-1325">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1325">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="48f42-1326">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1326">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="48f42-p179">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="48f42-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-1330">Requirements</span><span class="sxs-lookup"><span data-stu-id="48f42-1330">Requirements</span></span>

|<span data-ttu-id="48f42-1331">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1331">Requirement</span></span>|<span data-ttu-id="48f42-1332">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1332">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1333">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-1333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1334">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-1334">1.0</span></span>|
|[<span data-ttu-id="48f42-1335">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1336">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1337">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1338">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1338">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1339">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1339">Returns:</span></span>

<span data-ttu-id="48f42-p180">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="48f42-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="48f42-1342">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="48f42-1342">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="48f42-1343">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1343">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="48f42-1344">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1344">Example</span></span>

<span data-ttu-id="48f42-1345">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="48f42-1345">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="48f42-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="48f42-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="48f42-1347">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="48f42-1347">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1348">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-1348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48f42-1349">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1349">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="48f42-p181">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="48f42-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1352">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1352">Parameters</span></span>

|<span data-ttu-id="48f42-1353">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1353">Name</span></span>|<span data-ttu-id="48f42-1354">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1354">Type</span></span>|<span data-ttu-id="48f42-1355">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1355">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="48f42-1356">String</span><span class="sxs-lookup"><span data-stu-id="48f42-1356">String</span></span>|<span data-ttu-id="48f42-1357">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="48f42-1357">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1358">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1358">Requirements</span></span>

|<span data-ttu-id="48f42-1359">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1359">Requirement</span></span>|<span data-ttu-id="48f42-1360">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1360">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1361">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-1361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1362">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-1362">1.0</span></span>|
|[<span data-ttu-id="48f42-1363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1364">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1366">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1366">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1367">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1367">Returns:</span></span>

<span data-ttu-id="48f42-1368">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="48f42-1368">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="48f42-1369">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="48f42-1369">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="48f42-1370">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1370">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="48f42-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="48f42-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="48f42-1372">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-1372">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="48f42-p182">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="48f42-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1375">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1375">Parameters</span></span>

|<span data-ttu-id="48f42-1376">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1376">Name</span></span>|<span data-ttu-id="48f42-1377">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1377">Type</span></span>|<span data-ttu-id="48f42-1378">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1378">Attributes</span></span>|<span data-ttu-id="48f42-1379">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1379">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="48f42-1380">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="48f42-1380">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="48f42-p183">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="48f42-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="48f42-1384">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1384">Object</span></span>|<span data-ttu-id="48f42-1385">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1386">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1386">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1387">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1387">Object</span></span>|<span data-ttu-id="48f42-1388">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1388">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1389">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1389">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-1390">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1390">function</span></span>||<span data-ttu-id="48f42-1391">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1391">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48f42-1392">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1392">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="48f42-1393">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1393">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1394">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1394">Requirements</span></span>

|<span data-ttu-id="48f42-1395">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1395">Requirement</span></span>|<span data-ttu-id="48f42-1396">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1397">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="48f42-1398">1.2</span></span>|
|[<span data-ttu-id="48f42-1399">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1400">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1401">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1402">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-1402">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1403">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1403">Returns:</span></span>

<span data-ttu-id="48f42-1404">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1404">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="48f42-1405">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="48f42-1405">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="48f42-1406">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1406">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="48f42-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="48f42-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="48f42-1408">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="48f42-1408">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="48f42-1409">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="48f42-1409">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1410">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-1410">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-1411">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1411">Requirements</span></span>

|<span data-ttu-id="48f42-1412">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1412">Requirement</span></span>|<span data-ttu-id="48f42-1413">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1413">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1414">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1414">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1415">1.6</span><span class="sxs-lookup"><span data-stu-id="48f42-1415">1.6</span></span>|
|[<span data-ttu-id="48f42-1416">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1417">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1418">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1418">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1419">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1419">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1420">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1420">Returns:</span></span>

<span data-ttu-id="48f42-1421">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="48f42-1421">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="48f42-1422">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1422">Example</span></span>

<span data-ttu-id="48f42-1423">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="48f42-1423">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="48f42-1424">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="48f42-1424">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="48f42-p186">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="48f42-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1427">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="48f42-1427">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="48f42-p187">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="48f42-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="48f42-1431">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1431">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="48f42-1432">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1432">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="48f42-p188">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="48f42-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f42-1436">Requirements</span><span class="sxs-lookup"><span data-stu-id="48f42-1436">Requirements</span></span>

|<span data-ttu-id="48f42-1437">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1437">Requirement</span></span>|<span data-ttu-id="48f42-1438">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1438">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1439">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1440">1.6</span><span class="sxs-lookup"><span data-stu-id="48f42-1440">1.6</span></span>|
|[<span data-ttu-id="48f42-1441">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1442">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1443">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1444">Чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1444">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48f42-1445">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="48f42-1445">Returns:</span></span>

<span data-ttu-id="48f42-p189">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="48f42-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="48f42-1448">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1448">Example</span></span>

<span data-ttu-id="48f42-1449">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="48f42-1449">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="48f42-1450">Жетшаредпропертиесасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="48f42-1450">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="48f42-1451">Получает свойства выбранной встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="48f42-1451">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1452">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1452">Parameters</span></span>

|<span data-ttu-id="48f42-1453">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1453">Name</span></span>|<span data-ttu-id="48f42-1454">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1454">Type</span></span>|<span data-ttu-id="48f42-1455">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1455">Attributes</span></span>|<span data-ttu-id="48f42-1456">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1456">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48f42-1457">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1457">Object</span></span>|<span data-ttu-id="48f42-1458">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1458">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1459">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1459">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1460">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1460">Object</span></span>|<span data-ttu-id="48f42-1461">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1461">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1462">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1462">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-1463">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1463">function</span></span>||<span data-ttu-id="48f42-1464">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1464">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48f42-1465">Общие свойства предоставляются в виде [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) объекта в `asyncResult.value` свойстве.</span><span class="sxs-lookup"><span data-stu-id="48f42-1465">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="48f42-1466">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-1466">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1467">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1467">Requirements</span></span>

|<span data-ttu-id="48f42-1468">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1468">Requirement</span></span>|<span data-ttu-id="48f42-1469">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1470">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1471">1.8</span><span class="sxs-lookup"><span data-stu-id="48f42-1471">1.8</span></span>|
|[<span data-ttu-id="48f42-1472">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1473">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1474">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1475">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-1476">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1476">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="48f42-1477">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="48f42-1477">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="48f42-1478">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-1478">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="48f42-p191">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="48f42-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1482">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1482">Parameters</span></span>

|<span data-ttu-id="48f42-1483">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1483">Name</span></span>|<span data-ttu-id="48f42-1484">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1484">Type</span></span>|<span data-ttu-id="48f42-1485">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1485">Attributes</span></span>|<span data-ttu-id="48f42-1486">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1486">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="48f42-1487">function</span><span class="sxs-lookup"><span data-stu-id="48f42-1487">function</span></span>||<span data-ttu-id="48f42-1488">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1488">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48f42-1489">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1489">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="48f42-1490">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="48f42-1490">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="48f42-1491">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1491">Object</span></span>|<span data-ttu-id="48f42-1492">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1493">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1493">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="48f42-1494">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1494">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1495">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1495">Requirements</span></span>

|<span data-ttu-id="48f42-1496">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1496">Requirement</span></span>|<span data-ttu-id="48f42-1497">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1497">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1498">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1499">1.0</span><span class="sxs-lookup"><span data-stu-id="48f42-1499">1.0</span></span>|
|[<span data-ttu-id="48f42-1500">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1500">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1501">ReadItem</span></span>|
|[<span data-ttu-id="48f42-1502">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1502">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1503">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1503">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-1504">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1504">Example</span></span>

<span data-ttu-id="48f42-p194">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="48f42-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48f42-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="48f42-1509">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="48f42-1509">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="48f42-1510">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="48f42-1510">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="48f42-1511">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="48f42-1511">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="48f42-1512">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="48f42-1512">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="48f42-1513">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="48f42-1513">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1514">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1514">Parameters</span></span>

|<span data-ttu-id="48f42-1515">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1515">Name</span></span>|<span data-ttu-id="48f42-1516">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1516">Type</span></span>|<span data-ttu-id="48f42-1517">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1517">Attributes</span></span>|<span data-ttu-id="48f42-1518">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1518">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="48f42-1519">String</span><span class="sxs-lookup"><span data-stu-id="48f42-1519">String</span></span>||<span data-ttu-id="48f42-1520">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="48f42-1520">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="48f42-1521">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1521">Object</span></span>|<span data-ttu-id="48f42-1522">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1522">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1523">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1523">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1524">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1524">Object</span></span>|<span data-ttu-id="48f42-1525">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1525">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1526">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1526">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-1527">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1527">function</span></span>|<span data-ttu-id="48f42-1528">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1528">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1529">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1529">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="48f42-1530">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="48f42-1530">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="48f42-1531">Ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-1531">Errors</span></span>

|<span data-ttu-id="48f42-1532">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="48f42-1532">Error code</span></span>|<span data-ttu-id="48f42-1533">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1533">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="48f42-1534">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="48f42-1534">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1535">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1535">Requirements</span></span>

|<span data-ttu-id="48f42-1536">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1536">Requirement</span></span>|<span data-ttu-id="48f42-1537">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1537">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1538">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48f42-1538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1539">1.1</span><span class="sxs-lookup"><span data-stu-id="48f42-1539">1.1</span></span>|
|[<span data-ttu-id="48f42-1540">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1541">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1541">ReadWriteItem</span></span>|
|[<span data-ttu-id="48f42-1542">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1543">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-1543">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-1544">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1544">Example</span></span>

<span data-ttu-id="48f42-1545">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="48f42-1545">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="48f42-1546">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48f42-1546">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="48f42-1547">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="48f42-1547">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="48f42-1548">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1548">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1549">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1549">Parameters</span></span>

| <span data-ttu-id="48f42-1550">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1550">Name</span></span> | <span data-ttu-id="48f42-1551">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1551">Type</span></span> | <span data-ttu-id="48f42-1552">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1552">Attributes</span></span> | <span data-ttu-id="48f42-1553">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1553">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="48f42-1554">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="48f42-1554">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="48f42-1555">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="48f42-1555">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="48f42-1556">Объект</span><span class="sxs-lookup"><span data-stu-id="48f42-1556">Object</span></span> | <span data-ttu-id="48f42-1557">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1557">&lt;optional&gt;</span></span> | <span data-ttu-id="48f42-1558">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1558">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="48f42-1559">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1559">Object</span></span> | <span data-ttu-id="48f42-1560">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1560">&lt;optional&gt;</span></span> | <span data-ttu-id="48f42-1561">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1561">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="48f42-1562">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1562">function</span></span>| <span data-ttu-id="48f42-1563">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1563">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1564">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1564">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1565">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1565">Requirements</span></span>

|<span data-ttu-id="48f42-1566">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1566">Requirement</span></span>| <span data-ttu-id="48f42-1567">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1567">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1568">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48f42-1569">1.7</span><span class="sxs-lookup"><span data-stu-id="48f42-1569">1.7</span></span> |
|[<span data-ttu-id="48f42-1570">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1570">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48f42-1571">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1571">ReadItem</span></span> |
|[<span data-ttu-id="48f42-1572">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1572">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48f42-1573">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48f42-1573">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="48f42-1574">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="48f42-1574">saveAsync([options], callback)</span></span>

<span data-ttu-id="48f42-1575">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="48f42-1575">Asynchronously saves an item.</span></span>

<span data-ttu-id="48f42-1576">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1576">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="48f42-1577">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="48f42-1577">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="48f42-1578">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="48f42-1578">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1579">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="48f42-1579">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="48f42-1580">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="48f42-1580">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="48f42-p198">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="48f42-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="48f42-1584">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="48f42-1584">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="48f42-1585">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="48f42-1585">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="48f42-1586">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="48f42-1586">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="48f42-1587">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="48f42-1587">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="48f42-1588">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="48f42-1588">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1589">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1589">Parameters</span></span>

|<span data-ttu-id="48f42-1590">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1590">Name</span></span>|<span data-ttu-id="48f42-1591">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1591">Type</span></span>|<span data-ttu-id="48f42-1592">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1592">Attributes</span></span>|<span data-ttu-id="48f42-1593">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1593">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="48f42-1594">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1594">Object</span></span>|<span data-ttu-id="48f42-1595">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1595">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1596">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1596">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1597">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1597">Object</span></span>|<span data-ttu-id="48f42-1598">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1599">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="48f42-1599">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="48f42-1600">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1600">function</span></span>||<span data-ttu-id="48f42-1601">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48f42-1602">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1602">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1603">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1603">Requirements</span></span>

|<span data-ttu-id="48f42-1604">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1604">Requirement</span></span>|<span data-ttu-id="48f42-1605">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1605">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1606">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1607">1.3</span><span class="sxs-lookup"><span data-stu-id="48f42-1607">1.3</span></span>|
|[<span data-ttu-id="48f42-1608">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1609">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1609">ReadWriteItem</span></span>|
|[<span data-ttu-id="48f42-1610">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1611">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-1611">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="48f42-1612">Примеры</span><span class="sxs-lookup"><span data-stu-id="48f42-1612">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="48f42-p200">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="48f42-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="48f42-1615">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="48f42-1615">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="48f42-1616">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="48f42-1616">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="48f42-p201">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="48f42-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48f42-1620">Параметры</span><span class="sxs-lookup"><span data-stu-id="48f42-1620">Parameters</span></span>

|<span data-ttu-id="48f42-1621">Имя</span><span class="sxs-lookup"><span data-stu-id="48f42-1621">Name</span></span>|<span data-ttu-id="48f42-1622">Тип</span><span class="sxs-lookup"><span data-stu-id="48f42-1622">Type</span></span>|<span data-ttu-id="48f42-1623">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="48f42-1623">Attributes</span></span>|<span data-ttu-id="48f42-1624">Описание</span><span class="sxs-lookup"><span data-stu-id="48f42-1624">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="48f42-1625">String</span><span class="sxs-lookup"><span data-stu-id="48f42-1625">String</span></span>||<span data-ttu-id="48f42-p202">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="48f42-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="48f42-1629">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1629">Object</span></span>|<span data-ttu-id="48f42-1630">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1630">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1631">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="48f42-1631">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="48f42-1632">Object</span><span class="sxs-lookup"><span data-stu-id="48f42-1632">Object</span></span>|<span data-ttu-id="48f42-1633">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1633">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1634">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="48f42-1634">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="48f42-1635">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="48f42-1635">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="48f42-1636">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="48f42-1636">&lt;optional&gt;</span></span>|<span data-ttu-id="48f42-1637">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="48f42-1637">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="48f42-1638">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="48f42-1638">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="48f42-1639">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48f42-1639">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="48f42-1640">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="48f42-1640">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="48f42-1641">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="48f42-1641">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="48f42-1642">функция</span><span class="sxs-lookup"><span data-stu-id="48f42-1642">function</span></span>||<span data-ttu-id="48f42-1643">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48f42-1643">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48f42-1644">Требования</span><span class="sxs-lookup"><span data-stu-id="48f42-1644">Requirements</span></span>

|<span data-ttu-id="48f42-1645">Требование</span><span class="sxs-lookup"><span data-stu-id="48f42-1645">Requirement</span></span>|<span data-ttu-id="48f42-1646">Значение</span><span class="sxs-lookup"><span data-stu-id="48f42-1646">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f42-1647">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="48f42-1647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="48f42-1648">1.2</span><span class="sxs-lookup"><span data-stu-id="48f42-1648">1.2</span></span>|
|[<span data-ttu-id="48f42-1649">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48f42-1649">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="48f42-1650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="48f42-1650">ReadWriteItem</span></span>|
|[<span data-ttu-id="48f42-1651">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48f42-1651">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="48f42-1652">Создание</span><span class="sxs-lookup"><span data-stu-id="48f42-1652">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="48f42-1653">Пример</span><span class="sxs-lookup"><span data-stu-id="48f42-1653">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
