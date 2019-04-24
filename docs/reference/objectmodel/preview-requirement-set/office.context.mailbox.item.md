---
title: Office. Context. Mailbox. Item — Предварительная версия набора требований
description: ''
ms.date: 04/17/2019
localization_priority: Normal
ms.openlocfilehash: cb9c298302bf0df9d7842fde4706d9d0c9710ae4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450396"
---
# <a name="item"></a><span data-ttu-id="fff2d-102">item</span><span class="sxs-lookup"><span data-stu-id="fff2d-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="fff2d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="fff2d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="fff2d-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="fff2d-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="fff2d-106">Requirements</span></span>

|<span data-ttu-id="fff2d-107">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-107">Requirement</span></span>|<span data-ttu-id="fff2d-108">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-110">1.0</span></span>|
|[<span data-ttu-id="fff2d-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="fff2d-112">Restricted</span></span>|
|[<span data-ttu-id="fff2d-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fff2d-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="fff2d-115">Members and methods</span></span>

| <span data-ttu-id="fff2d-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-116">Member</span></span> | <span data-ttu-id="fff2d-117">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fff2d-118">attachments</span><span class="sxs-lookup"><span data-stu-id="fff2d-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="fff2d-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-119">Member</span></span> |
| [<span data-ttu-id="fff2d-120">bcc</span><span class="sxs-lookup"><span data-stu-id="fff2d-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="fff2d-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-121">Member</span></span> |
| [<span data-ttu-id="fff2d-122">body</span><span class="sxs-lookup"><span data-stu-id="fff2d-122">body</span></span>](#body-body) | <span data-ttu-id="fff2d-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-123">Member</span></span> |
| [<span data-ttu-id="fff2d-124">разделов</span><span class="sxs-lookup"><span data-stu-id="fff2d-124">categories</span></span>](#categories-categories) | <span data-ttu-id="fff2d-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-125">Member</span></span> |
| [<span data-ttu-id="fff2d-126">cc</span><span class="sxs-lookup"><span data-stu-id="fff2d-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="fff2d-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-127">Member</span></span> |
| [<span data-ttu-id="fff2d-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="fff2d-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="fff2d-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-129">Member</span></span> |
| [<span data-ttu-id="fff2d-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="fff2d-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="fff2d-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-131">Member</span></span> |
| [<span data-ttu-id="fff2d-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="fff2d-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="fff2d-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-133">Member</span></span> |
| [<span data-ttu-id="fff2d-134">end</span><span class="sxs-lookup"><span data-stu-id="fff2d-134">end</span></span>](#end-datetime) | <span data-ttu-id="fff2d-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-135">Member</span></span> |
| [<span data-ttu-id="fff2d-136">Енханцедлокатион</span><span class="sxs-lookup"><span data-stu-id="fff2d-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="fff2d-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-137">Member</span></span> |
| [<span data-ttu-id="fff2d-138">from</span><span class="sxs-lookup"><span data-stu-id="fff2d-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="fff2d-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-139">Member</span></span> |
| [<span data-ttu-id="fff2d-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="fff2d-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="fff2d-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-141">Member</span></span> |
| [<span data-ttu-id="fff2d-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="fff2d-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="fff2d-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-143">Member</span></span> |
| [<span data-ttu-id="fff2d-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="fff2d-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="fff2d-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-145">Member</span></span> |
| [<span data-ttu-id="fff2d-146">itemId</span><span class="sxs-lookup"><span data-stu-id="fff2d-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="fff2d-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-147">Member</span></span> |
| [<span data-ttu-id="fff2d-148">itemType</span><span class="sxs-lookup"><span data-stu-id="fff2d-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="fff2d-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-149">Member</span></span> |
| [<span data-ttu-id="fff2d-150">location</span><span class="sxs-lookup"><span data-stu-id="fff2d-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="fff2d-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-151">Member</span></span> |
| [<span data-ttu-id="fff2d-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="fff2d-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="fff2d-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-153">Member</span></span> |
| [<span data-ttu-id="fff2d-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="fff2d-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="fff2d-155">Member</span><span class="sxs-lookup"><span data-stu-id="fff2d-155">Member</span></span> |
| [<span data-ttu-id="fff2d-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="fff2d-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="fff2d-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-157">Member</span></span> |
| [<span data-ttu-id="fff2d-158">organizer</span><span class="sxs-lookup"><span data-stu-id="fff2d-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="fff2d-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-159">Member</span></span> |
| [<span data-ttu-id="fff2d-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="fff2d-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="fff2d-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-161">Member</span></span> |
| [<span data-ttu-id="fff2d-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="fff2d-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="fff2d-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-163">Member</span></span> |
| [<span data-ttu-id="fff2d-164">sender</span><span class="sxs-lookup"><span data-stu-id="fff2d-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="fff2d-165">Member</span><span class="sxs-lookup"><span data-stu-id="fff2d-165">Member</span></span> |
| [<span data-ttu-id="fff2d-166">seriesId</span><span class="sxs-lookup"><span data-stu-id="fff2d-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="fff2d-167">Member</span><span class="sxs-lookup"><span data-stu-id="fff2d-167">Member</span></span> |
| [<span data-ttu-id="fff2d-168">start</span><span class="sxs-lookup"><span data-stu-id="fff2d-168">start</span></span>](#start-datetime) | <span data-ttu-id="fff2d-169">Member</span><span class="sxs-lookup"><span data-stu-id="fff2d-169">Member</span></span> |
| [<span data-ttu-id="fff2d-170">subject</span><span class="sxs-lookup"><span data-stu-id="fff2d-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="fff2d-171">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-171">Member</span></span> |
| [<span data-ttu-id="fff2d-172">to</span><span class="sxs-lookup"><span data-stu-id="fff2d-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="fff2d-173">Элемент</span><span class="sxs-lookup"><span data-stu-id="fff2d-173">Member</span></span> |
| [<span data-ttu-id="fff2d-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="fff2d-175">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-175">Method</span></span> |
| [<span data-ttu-id="fff2d-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="fff2d-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="fff2d-177">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-177">Method</span></span> |
| [<span data-ttu-id="fff2d-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="fff2d-179">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-179">Method</span></span> |
| [<span data-ttu-id="fff2d-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="fff2d-181">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-181">Method</span></span> |
| [<span data-ttu-id="fff2d-182">close</span><span class="sxs-lookup"><span data-stu-id="fff2d-182">close</span></span>](#close) | <span data-ttu-id="fff2d-183">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-183">Method</span></span> |
| [<span data-ttu-id="fff2d-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="fff2d-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="fff2d-185">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-185">Method</span></span> |
| [<span data-ttu-id="fff2d-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="fff2d-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="fff2d-187">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-187">Method</span></span> |
| [<span data-ttu-id="fff2d-188">Жетаттачментконтентасинк</span><span class="sxs-lookup"><span data-stu-id="fff2d-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="fff2d-189">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-189">Method</span></span> |
| [<span data-ttu-id="fff2d-190">Жетаттачментсасинк</span><span class="sxs-lookup"><span data-stu-id="fff2d-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="fff2d-191">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-191">Method</span></span> |
| [<span data-ttu-id="fff2d-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="fff2d-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="fff2d-193">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-193">Method</span></span> |
| [<span data-ttu-id="fff2d-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="fff2d-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="fff2d-195">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-195">Method</span></span> |
| [<span data-ttu-id="fff2d-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="fff2d-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="fff2d-197">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-197">Method</span></span> |
| [<span data-ttu-id="fff2d-198">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="fff2d-199">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-199">Method</span></span> |
| [<span data-ttu-id="fff2d-200">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="fff2d-200">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="fff2d-201">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-201">Method</span></span> |
| [<span data-ttu-id="fff2d-202">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="fff2d-202">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="fff2d-203">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-203">Method</span></span> |
| [<span data-ttu-id="fff2d-204">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-204">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="fff2d-205">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-205">Method</span></span> |
| [<span data-ttu-id="fff2d-206">Жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="fff2d-206">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="fff2d-207">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-207">Method</span></span> |
| [<span data-ttu-id="fff2d-208">Жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="fff2d-208">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="fff2d-209">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-209">Method</span></span> |
| [<span data-ttu-id="fff2d-210">Жетшаредпропертиесасинк</span><span class="sxs-lookup"><span data-stu-id="fff2d-210">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="fff2d-211">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-211">Method</span></span> |
| [<span data-ttu-id="fff2d-212">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-212">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="fff2d-213">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-213">Method</span></span> |
| [<span data-ttu-id="fff2d-214">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-214">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="fff2d-215">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-215">Method</span></span> |
| [<span data-ttu-id="fff2d-216">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-216">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="fff2d-217">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-217">Method</span></span> |
| [<span data-ttu-id="fff2d-218">saveAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-218">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="fff2d-219">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-219">Method</span></span> |
| [<span data-ttu-id="fff2d-220">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="fff2d-220">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="fff2d-221">Метод</span><span class="sxs-lookup"><span data-stu-id="fff2d-221">Method</span></span> |

### <a name="example"></a><span data-ttu-id="fff2d-222">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-222">Example</span></span>

<span data-ttu-id="fff2d-223">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="fff2d-223">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="fff2d-224">Элементы</span><span class="sxs-lookup"><span data-stu-id="fff2d-224">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="fff2d-225">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fff2d-225">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="fff2d-226">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="fff2d-226">Gets the item's attachments as an array.</span></span> <span data-ttu-id="fff2d-227">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-227">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-228">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="fff2d-228">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="fff2d-229">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="fff2d-229">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-230">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-230">Type</span></span>

*   <span data-ttu-id="fff2d-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fff2d-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-232">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-232">Requirements</span></span>

|<span data-ttu-id="fff2d-233">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-233">Requirement</span></span>|<span data-ttu-id="fff2d-234">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-235">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-236">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-236">1.0</span></span>|
|[<span data-ttu-id="fff2d-237">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-238">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-239">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-240">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-240">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-241">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-241">Example</span></span>

<span data-ttu-id="fff2d-242">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-242">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

---
---

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fff2d-243">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fff2d-243">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fff2d-244">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-244">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="fff2d-245">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-245">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-246">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-246">Type</span></span>

*   [<span data-ttu-id="fff2d-247">Получатели</span><span class="sxs-lookup"><span data-stu-id="fff2d-247">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="fff2d-248">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-248">Requirements</span></span>

|<span data-ttu-id="fff2d-249">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-249">Requirement</span></span>|<span data-ttu-id="fff2d-250">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-251">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-252">1.1</span><span class="sxs-lookup"><span data-stu-id="fff2d-252">1.1</span></span>|
|[<span data-ttu-id="fff2d-253">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-254">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-255">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-256">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-256">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-257">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-257">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

---
---

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="fff2d-258">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="fff2d-258">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="fff2d-259">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-259">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-260">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-260">Type</span></span>

*   [<span data-ttu-id="fff2d-261">Body</span><span class="sxs-lookup"><span data-stu-id="fff2d-261">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="fff2d-262">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-262">Requirements</span></span>

|<span data-ttu-id="fff2d-263">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-263">Requirement</span></span>|<span data-ttu-id="fff2d-264">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-265">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-266">1.1</span><span class="sxs-lookup"><span data-stu-id="fff2d-266">1.1</span></span>|
|[<span data-ttu-id="fff2d-267">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-267">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-268">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-269">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-269">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-270">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-270">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-271">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-271">Example</span></span>

<span data-ttu-id="fff2d-272">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="fff2d-272">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="fff2d-273">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-273">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

####  <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="fff2d-274">Категории:[категории](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="fff2d-274">categories :[Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="fff2d-275">Получает объект, предоставляющий методы для управления категориями элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-275">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-276">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-276">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-277">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-277">Type</span></span>

*   [<span data-ttu-id="fff2d-278">Categories</span><span class="sxs-lookup"><span data-stu-id="fff2d-278">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="fff2d-279">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-279">Requirements</span></span>

|<span data-ttu-id="fff2d-280">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-280">Requirement</span></span>|<span data-ttu-id="fff2d-281">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-282">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-283">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fff2d-283">Preview</span></span>|
|[<span data-ttu-id="fff2d-284">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-285">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-286">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-287">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-288">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-288">Example</span></span>

<span data-ttu-id="fff2d-289">В этом примере возвращаются категории элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-289">This example gets the item's categories.</span></span>

```javascript
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fff2d-290">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fff2d-290">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fff2d-291">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-291">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="fff2d-292">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-292">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-293">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-293">Read mode</span></span>

<span data-ttu-id="fff2d-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-296">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-296">Compose mode</span></span>

<span data-ttu-id="fff2d-297">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-297">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fff2d-298">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-298">Type</span></span>

*   <span data-ttu-id="fff2d-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fff2d-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-300">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-300">Requirements</span></span>

|<span data-ttu-id="fff2d-301">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-301">Requirement</span></span>|<span data-ttu-id="fff2d-302">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-303">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-304">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-304">1.0</span></span>|
|[<span data-ttu-id="fff2d-305">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-306">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-307">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-308">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-308">Compose or Read</span></span>|

---
---

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="fff2d-309">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="fff2d-309">(nullable) conversationId :String</span></span>

<span data-ttu-id="fff2d-310">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="fff2d-310">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="fff2d-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="fff2d-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-315">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-315">Type</span></span>

*   <span data-ttu-id="fff2d-316">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-316">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-317">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-317">Requirements</span></span>

|<span data-ttu-id="fff2d-318">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-318">Requirement</span></span>|<span data-ttu-id="fff2d-319">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-320">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-321">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-321">1.0</span></span>|
|[<span data-ttu-id="fff2d-322">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-323">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-324">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-325">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-325">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-326">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-326">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="fff2d-327">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="fff2d-327">dateTimeCreated :Date</span></span>

<span data-ttu-id="fff2d-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-330">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-330">Type</span></span>

*   <span data-ttu-id="fff2d-331">Дата</span><span class="sxs-lookup"><span data-stu-id="fff2d-331">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-332">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-332">Requirements</span></span>

|<span data-ttu-id="fff2d-333">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-333">Requirement</span></span>|<span data-ttu-id="fff2d-334">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-335">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-336">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-336">1.0</span></span>|
|[<span data-ttu-id="fff2d-337">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-338">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-339">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-340">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-340">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-341">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-341">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="fff2d-342">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="fff2d-342">dateTimeModified :Date</span></span>

<span data-ttu-id="fff2d-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-345">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-345">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-346">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-346">Type</span></span>

*   <span data-ttu-id="fff2d-347">Дата</span><span class="sxs-lookup"><span data-stu-id="fff2d-347">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-348">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-348">Requirements</span></span>

|<span data-ttu-id="fff2d-349">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-349">Requirement</span></span>|<span data-ttu-id="fff2d-350">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-351">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-352">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-352">1.0</span></span>|
|[<span data-ttu-id="fff2d-353">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-354">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-355">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-356">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-357">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-357">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="fff2d-358">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="fff2d-358">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="fff2d-359">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-359">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="fff2d-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="fff2d-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-362">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-362">Read mode</span></span>

<span data-ttu-id="fff2d-363">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-363">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-364">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-364">Compose mode</span></span>

<span data-ttu-id="fff2d-365">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-365">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="fff2d-366">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="fff2d-366">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="fff2d-367">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-367">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="fff2d-368">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-368">Type</span></span>

*   <span data-ttu-id="fff2d-369">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="fff2d-369">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-370">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-370">Requirements</span></span>

|<span data-ttu-id="fff2d-371">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-371">Requirement</span></span>|<span data-ttu-id="fff2d-372">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-372">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-373">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-373">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-374">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-374">1.0</span></span>|
|[<span data-ttu-id="fff2d-375">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-375">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-376">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-376">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-377">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-377">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-378">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-378">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="fff2d-379">Енханцедлокатион:[енханцедлокатион](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="fff2d-379">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="fff2d-380">Получает или задает расположение встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-380">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-381">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-381">Read mode</span></span>

<span data-ttu-id="fff2d-382">Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который позволяет получить набор расположений (каждый, представленный объектом локатиондетаилс), связанный с встречей. [](/javascript/api/outlook/office.locationdetails) `enhancedLocation`</span><span class="sxs-lookup"><span data-stu-id="fff2d-382">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-383">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-383">Compose mode</span></span>

<span data-ttu-id="fff2d-384">`enhancedLocation` Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который предоставляет методы для получения, удаления или добавления расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-385">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-385">Type</span></span>

*   [<span data-ttu-id="fff2d-386">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="fff2d-386">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="fff2d-387">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-387">Requirements</span></span>

|<span data-ttu-id="fff2d-388">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-388">Requirement</span></span>|<span data-ttu-id="fff2d-389">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-390">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-391">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fff2d-391">Preview</span></span>|
|[<span data-ttu-id="fff2d-392">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-393">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-394">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-395">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-395">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-396">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-396">Example</span></span>

<span data-ttu-id="fff2d-397">В следующем примере показано получение текущих расположений, связанных с встречей.</span><span class="sxs-lookup"><span data-stu-id="fff2d-397">The following example gets the current locations associated with the appointment.</span></span>

```javascript
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

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="fff2d-398">от:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="fff2d-398">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="fff2d-399">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-399">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="fff2d-p112">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-402">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-402">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-403">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-403">Read mode</span></span>

<span data-ttu-id="fff2d-404">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="fff2d-404">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-405">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-405">Compose mode</span></span>

<span data-ttu-id="fff2d-406">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="fff2d-406">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fff2d-407">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-407">Type</span></span>

*   <span data-ttu-id="fff2d-408">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [из](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="fff2d-408">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-409">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-409">Requirements</span></span>

|<span data-ttu-id="fff2d-410">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-410">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="fff2d-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-412">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-412">1.0</span></span>|<span data-ttu-id="fff2d-413">1.7</span><span class="sxs-lookup"><span data-stu-id="fff2d-413">1.7</span></span>|
|[<span data-ttu-id="fff2d-414">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-415">ReadItem</span></span>|<span data-ttu-id="fff2d-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-416">ReadWriteItem</span></span>|
|[<span data-ttu-id="fff2d-417">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-418">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-418">Read</span></span>|<span data-ttu-id="fff2d-419">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-419">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="fff2d-420">Internetheaders::[internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="fff2d-420">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="fff2d-421">Возвращает или задает заголовки Интернета сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-421">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-422">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-422">Type</span></span>

*   [<span data-ttu-id="fff2d-423">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="fff2d-423">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="fff2d-424">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-424">Requirements</span></span>

|<span data-ttu-id="fff2d-425">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-425">Requirement</span></span>|<span data-ttu-id="fff2d-426">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-427">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-428">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fff2d-428">Preview</span></span>|
|[<span data-ttu-id="fff2d-429">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-430">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-431">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-432">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-432">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-433">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-433">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="fff2d-434">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="fff2d-434">internetMessageId :String</span></span>

<span data-ttu-id="fff2d-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-437">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-437">Type</span></span>

*   <span data-ttu-id="fff2d-438">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-438">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-439">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-439">Requirements</span></span>

|<span data-ttu-id="fff2d-440">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-440">Requirement</span></span>|<span data-ttu-id="fff2d-441">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-442">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-443">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-443">1.0</span></span>|
|[<span data-ttu-id="fff2d-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-445">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-447">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-447">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-448">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-448">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="fff2d-449">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="fff2d-449">itemClass :String</span></span>

<span data-ttu-id="fff2d-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="fff2d-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="fff2d-454">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-454">Type</span></span>|<span data-ttu-id="fff2d-455">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-455">Description</span></span>|<span data-ttu-id="fff2d-456">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="fff2d-456">item class</span></span>|
|---|---|---|
|<span data-ttu-id="fff2d-457">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="fff2d-457">Appointment items</span></span>|<span data-ttu-id="fff2d-458">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-458">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="fff2d-459">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="fff2d-459">Message items</span></span>|<span data-ttu-id="fff2d-460">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-460">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="fff2d-461">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-461">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-462">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-462">Type</span></span>

*   <span data-ttu-id="fff2d-463">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-464">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-464">Requirements</span></span>

|<span data-ttu-id="fff2d-465">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-465">Requirement</span></span>|<span data-ttu-id="fff2d-466">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-467">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-468">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-468">1.0</span></span>|
|[<span data-ttu-id="fff2d-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-470">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-472">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-473">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-473">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="fff2d-474">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="fff2d-474">(nullable) itemId :String</span></span>

<span data-ttu-id="fff2d-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-477">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="fff2d-477">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="fff2d-478">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="fff2d-478">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="fff2d-479">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="fff2d-479">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="fff2d-480">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="fff2d-480">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="fff2d-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-483">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-483">Type</span></span>

*   <span data-ttu-id="fff2d-484">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-484">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-485">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-485">Requirements</span></span>

|<span data-ttu-id="fff2d-486">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-486">Requirement</span></span>|<span data-ttu-id="fff2d-487">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-488">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-489">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-489">1.0</span></span>|
|[<span data-ttu-id="fff2d-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-491">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-493">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-493">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-494">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-494">Example</span></span>

<span data-ttu-id="fff2d-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

---
---

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="fff2d-497">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="fff2d-497">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="fff2d-498">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="fff2d-498">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="fff2d-499">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="fff2d-499">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-500">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-500">Type</span></span>

*   [<span data-ttu-id="fff2d-501">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="fff2d-501">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="fff2d-502">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-502">Requirements</span></span>

|<span data-ttu-id="fff2d-503">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-503">Requirement</span></span>|<span data-ttu-id="fff2d-504">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-505">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-506">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-506">1.0</span></span>|
|[<span data-ttu-id="fff2d-507">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-508">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-509">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-510">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-510">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-511">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-511">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="fff2d-512">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="fff2d-512">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="fff2d-513">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-513">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-514">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-514">Read mode</span></span>

<span data-ttu-id="fff2d-515">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-515">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-516">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-516">Compose mode</span></span>

<span data-ttu-id="fff2d-517">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-517">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fff2d-518">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-518">Type</span></span>

*   <span data-ttu-id="fff2d-519">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="fff2d-519">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-520">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-520">Requirements</span></span>

|<span data-ttu-id="fff2d-521">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-521">Requirement</span></span>|<span data-ttu-id="fff2d-522">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-522">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-523">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-524">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-524">1.0</span></span>|
|[<span data-ttu-id="fff2d-525">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-526">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-527">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-528">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-528">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="fff2d-529">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="fff2d-529">normalizedSubject :String</span></span>

<span data-ttu-id="fff2d-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="fff2d-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="fff2d-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-534">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-534">Type</span></span>

*   <span data-ttu-id="fff2d-535">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-535">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-536">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-536">Requirements</span></span>

|<span data-ttu-id="fff2d-537">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-537">Requirement</span></span>|<span data-ttu-id="fff2d-538">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-539">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-540">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-540">1.0</span></span>|
|[<span data-ttu-id="fff2d-541">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-542">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-543">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-544">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-545">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-545">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="fff2d-546">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="fff2d-546">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="fff2d-547">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-547">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-548">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-548">Type</span></span>

*   [<span data-ttu-id="fff2d-549">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="fff2d-549">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="fff2d-550">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-550">Requirements</span></span>

|<span data-ttu-id="fff2d-551">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-551">Requirement</span></span>|<span data-ttu-id="fff2d-552">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-553">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-554">1.3</span><span class="sxs-lookup"><span data-stu-id="fff2d-554">1.3</span></span>|
|[<span data-ttu-id="fff2d-555">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-556">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-557">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-558">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-558">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-559">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-559">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

---
---

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fff2d-560">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fff2d-560">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fff2d-561">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="fff2d-561">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="fff2d-562">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-562">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-563">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-563">Read mode</span></span>

<span data-ttu-id="fff2d-564">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-564">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-565">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-565">Compose mode</span></span>

<span data-ttu-id="fff2d-566">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-566">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fff2d-567">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-567">Type</span></span>

*   <span data-ttu-id="fff2d-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fff2d-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-569">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-569">Requirements</span></span>

|<span data-ttu-id="fff2d-570">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-570">Requirement</span></span>|<span data-ttu-id="fff2d-571">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-571">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-572">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-572">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-573">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-573">1.0</span></span>|
|[<span data-ttu-id="fff2d-574">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-574">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-575">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-575">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-576">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-576">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-577">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-577">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="fff2d-578">Организатор:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Организатор](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="fff2d-578">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="fff2d-579">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-579">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-580">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-580">Read mode</span></span>

<span data-ttu-id="fff2d-581">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-581">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-582">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-582">Compose mode</span></span>

<span data-ttu-id="fff2d-583">Свойство возвращает объект организатора, который предоставляет метод для получения значения организатора. [](/javascript/api/outlook/office.organizer) `organizer`</span><span class="sxs-lookup"><span data-stu-id="fff2d-583">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="fff2d-584">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-584">Type</span></span>

*   <span data-ttu-id="fff2d-585">[](/javascript/api/outlook/office.emailaddressdetails) | [Организатор](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fff2d-585">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-586">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-586">Requirements</span></span>

|<span data-ttu-id="fff2d-587">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-587">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="fff2d-588">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-589">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-589">1.0</span></span>|<span data-ttu-id="fff2d-590">1.7</span><span class="sxs-lookup"><span data-stu-id="fff2d-590">1.7</span></span>|
|[<span data-ttu-id="fff2d-591">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-591">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-592">ReadItem</span></span>|<span data-ttu-id="fff2d-593">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-593">ReadWriteItem</span></span>|
|[<span data-ttu-id="fff2d-594">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-595">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-595">Read</span></span>|<span data-ttu-id="fff2d-596">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-596">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="fff2d-597">(Nullable) повторение[](/javascript/api/outlook/office.recurrence) : повторение</span><span class="sxs-lookup"><span data-stu-id="fff2d-597">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="fff2d-598">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-598">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="fff2d-599">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="fff2d-599">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="fff2d-600">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-600">Read and compose modes for appointment items.</span></span> <span data-ttu-id="fff2d-601">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-601">Read mode for meeting request items.</span></span>

<span data-ttu-id="fff2d-602">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="fff2d-602">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="fff2d-603">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="fff2d-603">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="fff2d-604">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-604">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="fff2d-605">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="fff2d-605">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="fff2d-606">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="fff2d-606">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-607">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-607">Read mode</span></span>

<span data-ttu-id="fff2d-608">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-608">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="fff2d-609">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-609">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-610">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-610">Compose mode</span></span>

<span data-ttu-id="fff2d-611">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-611">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="fff2d-612">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="fff2d-612">This is available for appointments.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="fff2d-613">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-613">Type</span></span>

* [<span data-ttu-id="fff2d-614">Повторения</span><span class="sxs-lookup"><span data-stu-id="fff2d-614">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="fff2d-615">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-615">Requirement</span></span>|<span data-ttu-id="fff2d-616">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-617">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-618">1.7</span><span class="sxs-lookup"><span data-stu-id="fff2d-618">1.7</span></span>|
|[<span data-ttu-id="fff2d-619">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-619">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-620">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-621">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-621">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-622">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-622">Compose or Read</span></span>|

---
---

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fff2d-623">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fff2d-623">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fff2d-624">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="fff2d-624">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="fff2d-625">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-625">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-626">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-626">Read mode</span></span>

<span data-ttu-id="fff2d-627">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-627">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-628">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-628">Compose mode</span></span>

<span data-ttu-id="fff2d-629">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-629">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="fff2d-630">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-630">Type</span></span>

*   <span data-ttu-id="fff2d-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fff2d-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-632">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-632">Requirements</span></span>

|<span data-ttu-id="fff2d-633">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-633">Requirement</span></span>|<span data-ttu-id="fff2d-634">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-635">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-636">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-636">1.0</span></span>|
|[<span data-ttu-id="fff2d-637">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-638">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-639">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-640">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-640">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="fff2d-641">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fff2d-641">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="fff2d-p128">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="fff2d-p129">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-646">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-646">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-647">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-647">Type</span></span>

*   [<span data-ttu-id="fff2d-648">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fff2d-648">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fff2d-649">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-649">Requirements</span></span>

|<span data-ttu-id="fff2d-650">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-650">Requirement</span></span>|<span data-ttu-id="fff2d-651">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-651">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-652">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-653">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-653">1.0</span></span>|
|[<span data-ttu-id="fff2d-654">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-655">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-656">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-657">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-657">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-658">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-658">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="fff2d-659">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-659">(nullable) seriesId :String</span></span>

<span data-ttu-id="fff2d-660">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="fff2d-660">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="fff2d-661">В OWA и Outlook `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="fff2d-661">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="fff2d-662">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-662">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-663">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="fff2d-663">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="fff2d-664">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="fff2d-664">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="fff2d-665">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="fff2d-665">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="fff2d-666">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="fff2d-666">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="fff2d-667">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-667">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="fff2d-668">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-668">Type</span></span>

* <span data-ttu-id="fff2d-669">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-669">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-670">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-670">Requirements</span></span>

|<span data-ttu-id="fff2d-671">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-671">Requirement</span></span>|<span data-ttu-id="fff2d-672">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-673">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-674">1.7</span><span class="sxs-lookup"><span data-stu-id="fff2d-674">1.7</span></span>|
|[<span data-ttu-id="fff2d-675">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-676">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-677">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-678">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-678">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-679">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-679">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

---
---

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="fff2d-680">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="fff2d-680">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="fff2d-681">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-681">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="fff2d-p132">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="fff2d-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-684">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-684">Read mode</span></span>

<span data-ttu-id="fff2d-685">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-685">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-686">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-686">Compose mode</span></span>

<span data-ttu-id="fff2d-687">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-687">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="fff2d-688">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="fff2d-688">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="fff2d-689">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-689">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="fff2d-690">Type</span><span class="sxs-lookup"><span data-stu-id="fff2d-690">Type</span></span>

*   <span data-ttu-id="fff2d-691">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="fff2d-691">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-692">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-692">Requirements</span></span>

|<span data-ttu-id="fff2d-693">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-693">Requirement</span></span>|<span data-ttu-id="fff2d-694">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-695">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-696">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-696">1.0</span></span>|
|[<span data-ttu-id="fff2d-697">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-698">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-699">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-700">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-700">Compose or Read</span></span>|

---
---

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="fff2d-701">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fff2d-701">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="fff2d-702">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-702">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="fff2d-703">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="fff2d-703">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-704">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-704">Read mode</span></span>

<span data-ttu-id="fff2d-p133">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="fff2d-707">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="fff2d-707">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-708">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-708">Compose mode</span></span>
<span data-ttu-id="fff2d-709">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="fff2d-709">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="fff2d-710">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-710">Type</span></span>

*   <span data-ttu-id="fff2d-711">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fff2d-711">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-712">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-712">Requirements</span></span>

|<span data-ttu-id="fff2d-713">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-713">Requirement</span></span>|<span data-ttu-id="fff2d-714">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-715">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-716">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-716">1.0</span></span>|
|[<span data-ttu-id="fff2d-717">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-718">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-719">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-720">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-720">Compose or Read</span></span>|

---
---

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fff2d-721">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fff2d-721">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fff2d-722">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-722">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="fff2d-723">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-723">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fff2d-724">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fff2d-724">Read mode</span></span>

<span data-ttu-id="fff2d-p135">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="fff2d-727">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fff2d-727">Compose mode</span></span>

<span data-ttu-id="fff2d-728">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-728">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fff2d-729">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-729">Type</span></span>

*   <span data-ttu-id="fff2d-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fff2d-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-731">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-731">Requirements</span></span>

|<span data-ttu-id="fff2d-732">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-732">Requirement</span></span>|<span data-ttu-id="fff2d-733">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-734">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-735">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-735">1.0</span></span>|
|[<span data-ttu-id="fff2d-736">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-737">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-738">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-739">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-739">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="fff2d-740">Методы</span><span class="sxs-lookup"><span data-stu-id="fff2d-740">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="fff2d-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fff2d-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fff2d-742">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-742">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="fff2d-743">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-743">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="fff2d-744">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="fff2d-744">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-745">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-745">Parameters</span></span>
|<span data-ttu-id="fff2d-746">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-746">Name</span></span>|<span data-ttu-id="fff2d-747">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-747">Type</span></span>|<span data-ttu-id="fff2d-748">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-748">Attributes</span></span>|<span data-ttu-id="fff2d-749">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-749">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="fff2d-750">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-750">String</span></span>||<span data-ttu-id="fff2d-p136">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="fff2d-753">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-753">String</span></span>||<span data-ttu-id="fff2d-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="fff2d-756">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-756">Object</span></span>|<span data-ttu-id="fff2d-757">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-757">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-758">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-758">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-759">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-759">Object</span></span>|<span data-ttu-id="fff2d-760">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-760">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-761">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="fff2d-761">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="fff2d-762">Boolean</span><span class="sxs-lookup"><span data-stu-id="fff2d-762">Boolean</span></span>|<span data-ttu-id="fff2d-763">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-763">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-764">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="fff2d-764">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="fff2d-765">function</span><span class="sxs-lookup"><span data-stu-id="fff2d-765">function</span></span>|<span data-ttu-id="fff2d-766">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-766">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-767">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-767">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fff2d-768">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-768">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fff2d-769">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="fff2d-769">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fff2d-770">Ошибки</span><span class="sxs-lookup"><span data-stu-id="fff2d-770">Errors</span></span>

|<span data-ttu-id="fff2d-771">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="fff2d-771">Error code</span></span>|<span data-ttu-id="fff2d-772">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-772">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="fff2d-773">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="fff2d-773">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="fff2d-774">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="fff2d-774">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="fff2d-775">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="fff2d-775">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-776">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-776">Requirements</span></span>

|<span data-ttu-id="fff2d-777">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-777">Requirement</span></span>|<span data-ttu-id="fff2d-778">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-778">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-779">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-779">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-780">1.1</span><span class="sxs-lookup"><span data-stu-id="fff2d-780">1.1</span></span>|
|[<span data-ttu-id="fff2d-781">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-781">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-782">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-782">ReadWriteItem</span></span>|
|[<span data-ttu-id="fff2d-783">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-783">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-784">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-784">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fff2d-785">Примеры</span><span class="sxs-lookup"><span data-stu-id="fff2d-785">Examples</span></span>

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

<span data-ttu-id="fff2d-786">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-786">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```javascript
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

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="fff2d-787">addFileAttachmentFromBase64Async (base64File, Аттачментнаме, [параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="fff2d-787">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fff2d-788">Добавляет файл из кодировки Base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-788">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="fff2d-789">`addFileAttachmentFromBase64Async` Метод передает файл из кодировки Base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-789">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="fff2d-790">Этот метод возвращает идентификатор вложения в объекте AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="fff2d-790">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="fff2d-791">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="fff2d-791">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-792">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-792">Parameters</span></span>

|<span data-ttu-id="fff2d-793">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-793">Name</span></span>|<span data-ttu-id="fff2d-794">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-794">Type</span></span>|<span data-ttu-id="fff2d-795">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-795">Attributes</span></span>|<span data-ttu-id="fff2d-796">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-796">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="fff2d-797">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-797">String</span></span>||<span data-ttu-id="fff2d-798">Содержимое изображения или файла в кодировке Base64, которое добавляется в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="fff2d-798">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="fff2d-799">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-799">String</span></span>||<span data-ttu-id="fff2d-p139">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="fff2d-802">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-802">Object</span></span>|<span data-ttu-id="fff2d-803">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-803">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-804">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-804">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-805">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-805">Object</span></span>|<span data-ttu-id="fff2d-806">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-806">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-807">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="fff2d-807">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="fff2d-808">Boolean</span><span class="sxs-lookup"><span data-stu-id="fff2d-808">Boolean</span></span>|<span data-ttu-id="fff2d-809">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-809">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-810">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="fff2d-810">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="fff2d-811">function</span><span class="sxs-lookup"><span data-stu-id="fff2d-811">function</span></span>|<span data-ttu-id="fff2d-812">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-812">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-813">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-813">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fff2d-814">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-814">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fff2d-815">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="fff2d-815">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fff2d-816">Ошибки</span><span class="sxs-lookup"><span data-stu-id="fff2d-816">Errors</span></span>

|<span data-ttu-id="fff2d-817">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="fff2d-817">Error code</span></span>|<span data-ttu-id="fff2d-818">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-818">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="fff2d-819">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="fff2d-819">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="fff2d-820">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="fff2d-820">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="fff2d-821">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="fff2d-821">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-822">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-822">Requirements</span></span>

|<span data-ttu-id="fff2d-823">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-823">Requirement</span></span>|<span data-ttu-id="fff2d-824">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-825">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-826">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fff2d-826">Preview</span></span>|
|[<span data-ttu-id="fff2d-827">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-827">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-828">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-828">ReadWriteItem</span></span>|
|[<span data-ttu-id="fff2d-829">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-829">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-830">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-830">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fff2d-831">Примеры</span><span class="sxs-lookup"><span data-stu-id="fff2d-831">Examples</span></span>

```javascript
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

---
---

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="fff2d-832">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fff2d-832">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="fff2d-833">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="fff2d-833">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="fff2d-834">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-834">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-835">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-835">Parameters</span></span>

| <span data-ttu-id="fff2d-836">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-836">Name</span></span> | <span data-ttu-id="fff2d-837">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-837">Type</span></span> | <span data-ttu-id="fff2d-838">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-838">Attributes</span></span> | <span data-ttu-id="fff2d-839">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-839">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="fff2d-840">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="fff2d-840">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="fff2d-841">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="fff2d-841">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="fff2d-842">Function</span><span class="sxs-lookup"><span data-stu-id="fff2d-842">Function</span></span> || <span data-ttu-id="fff2d-p140">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="fff2d-846">Объект</span><span class="sxs-lookup"><span data-stu-id="fff2d-846">Object</span></span> | <span data-ttu-id="fff2d-847">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-847">&lt;optional&gt;</span></span> | <span data-ttu-id="fff2d-848">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-848">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="fff2d-849">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-849">Object</span></span> | <span data-ttu-id="fff2d-850">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-850">&lt;optional&gt;</span></span> | <span data-ttu-id="fff2d-851">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-851">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="fff2d-852">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-852">function</span></span>| <span data-ttu-id="fff2d-853">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-853">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-854">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-854">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-855">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-855">Requirements</span></span>

|<span data-ttu-id="fff2d-856">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-856">Requirement</span></span>| <span data-ttu-id="fff2d-857">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-857">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-858">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-858">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fff2d-859">1.7</span><span class="sxs-lookup"><span data-stu-id="fff2d-859">1.7</span></span> |
|[<span data-ttu-id="fff2d-860">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-860">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fff2d-861">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-861">ReadItem</span></span> |
|[<span data-ttu-id="fff2d-862">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-862">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fff2d-863">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-863">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="fff2d-864">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-864">Example</span></span>

```javascript
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

---
---

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="fff2d-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fff2d-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fff2d-866">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-866">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="fff2d-p141">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="fff2d-870">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="fff2d-870">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="fff2d-871">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="fff2d-871">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-872">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-872">Parameters</span></span>

|<span data-ttu-id="fff2d-873">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-873">Name</span></span>|<span data-ttu-id="fff2d-874">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-874">Type</span></span>|<span data-ttu-id="fff2d-875">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-875">Attributes</span></span>|<span data-ttu-id="fff2d-876">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-876">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="fff2d-877">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-877">String</span></span>||<span data-ttu-id="fff2d-p142">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="fff2d-880">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-880">String</span></span>||<span data-ttu-id="fff2d-881">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-881">The subject of the item to be attached.</span></span> <span data-ttu-id="fff2d-882">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-882">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="fff2d-883">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-883">Object</span></span>|<span data-ttu-id="fff2d-884">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-884">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-885">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-885">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-886">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-886">Object</span></span>|<span data-ttu-id="fff2d-887">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-887">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-888">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-888">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fff2d-889">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-889">function</span></span>|<span data-ttu-id="fff2d-890">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-890">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-891">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-891">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fff2d-892">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-892">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fff2d-893">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="fff2d-893">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fff2d-894">Ошибки</span><span class="sxs-lookup"><span data-stu-id="fff2d-894">Errors</span></span>

|<span data-ttu-id="fff2d-895">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="fff2d-895">Error code</span></span>|<span data-ttu-id="fff2d-896">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-896">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="fff2d-897">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="fff2d-897">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-898">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-898">Requirements</span></span>

|<span data-ttu-id="fff2d-899">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-899">Requirement</span></span>|<span data-ttu-id="fff2d-900">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-900">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-901">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-901">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-902">1.1</span><span class="sxs-lookup"><span data-stu-id="fff2d-902">1.1</span></span>|
|[<span data-ttu-id="fff2d-903">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-903">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-904">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-904">ReadWriteItem</span></span>|
|[<span data-ttu-id="fff2d-905">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-905">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-906">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-906">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-907">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-907">Example</span></span>

<span data-ttu-id="fff2d-908">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-908">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

---
---

####  <a name="close"></a><span data-ttu-id="fff2d-909">close()</span><span class="sxs-lookup"><span data-stu-id="fff2d-909">close()</span></span>

<span data-ttu-id="fff2d-910">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="fff2d-910">Closes the current item that is being composed.</span></span>

<span data-ttu-id="fff2d-p144">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-913">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-913">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="fff2d-914">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="fff2d-914">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-915">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-915">Requirements</span></span>

|<span data-ttu-id="fff2d-916">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-916">Requirement</span></span>|<span data-ttu-id="fff2d-917">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-917">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-918">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-918">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-919">1.3</span><span class="sxs-lookup"><span data-stu-id="fff2d-919">1.3</span></span>|
|[<span data-ttu-id="fff2d-920">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-920">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-921">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="fff2d-921">Restricted</span></span>|
|[<span data-ttu-id="fff2d-922">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-922">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-923">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-923">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="fff2d-924">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="fff2d-924">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="fff2d-925">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-925">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-926">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-926">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fff2d-927">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="fff2d-927">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fff2d-928">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="fff2d-928">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="fff2d-p145">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-932">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-932">Parameters</span></span>

|<span data-ttu-id="fff2d-933">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-933">Name</span></span>|<span data-ttu-id="fff2d-934">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-934">Type</span></span>|<span data-ttu-id="fff2d-935">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-935">Attributes</span></span>|<span data-ttu-id="fff2d-936">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-936">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="fff2d-937">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-937">String &#124; Object</span></span>||<span data-ttu-id="fff2d-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fff2d-940">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="fff2d-940">**OR**</span></span><br/><span data-ttu-id="fff2d-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="fff2d-943">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-943">String</span></span>|<span data-ttu-id="fff2d-944">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-944">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="fff2d-947">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-947">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="fff2d-948">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-948">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-949">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="fff2d-949">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="fff2d-950">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-950">String</span></span>||<span data-ttu-id="fff2d-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="fff2d-953">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-953">String</span></span>||<span data-ttu-id="fff2d-954">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-954">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="fff2d-955">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-955">String</span></span>||<span data-ttu-id="fff2d-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="fff2d-958">Логический</span><span class="sxs-lookup"><span data-stu-id="fff2d-958">Boolean</span></span>||<span data-ttu-id="fff2d-p151">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="fff2d-961">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-961">String</span></span>||<span data-ttu-id="fff2d-p152">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="fff2d-965">function</span><span class="sxs-lookup"><span data-stu-id="fff2d-965">function</span></span>|<span data-ttu-id="fff2d-966">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-966">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-967">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-968">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-968">Requirements</span></span>

|<span data-ttu-id="fff2d-969">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-969">Requirement</span></span>|<span data-ttu-id="fff2d-970">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-970">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-971">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-971">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-972">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-972">1.0</span></span>|
|[<span data-ttu-id="fff2d-973">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-973">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-974">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-974">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-975">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-975">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-976">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-976">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fff2d-977">Примеры</span><span class="sxs-lookup"><span data-stu-id="fff2d-977">Examples</span></span>

<span data-ttu-id="fff2d-978">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-978">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="fff2d-979">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-979">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="fff2d-980">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-980">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fff2d-981">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="fff2d-981">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fff2d-982">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="fff2d-982">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fff2d-983">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="fff2d-983">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="fff2d-984">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="fff2d-984">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="fff2d-985">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-985">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-986">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-986">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fff2d-987">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="fff2d-987">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fff2d-988">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="fff2d-988">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="fff2d-p153">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-992">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-992">Parameters</span></span>

|<span data-ttu-id="fff2d-993">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-993">Name</span></span>|<span data-ttu-id="fff2d-994">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-994">Type</span></span>|<span data-ttu-id="fff2d-995">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-995">Attributes</span></span>|<span data-ttu-id="fff2d-996">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-996">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="fff2d-997">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-997">String &#124; Object</span></span>||<span data-ttu-id="fff2d-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fff2d-1000">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="fff2d-1000">**OR**</span></span><br/><span data-ttu-id="fff2d-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="fff2d-1003">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-1003">String</span></span>|<span data-ttu-id="fff2d-1004">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="fff2d-1007">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1007">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="fff2d-1008">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1009">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1009">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="fff2d-1010">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-1010">String</span></span>||<span data-ttu-id="fff2d-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="fff2d-1013">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-1013">String</span></span>||<span data-ttu-id="fff2d-1014">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1014">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="fff2d-1015">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-1015">String</span></span>||<span data-ttu-id="fff2d-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="fff2d-1018">Логический</span><span class="sxs-lookup"><span data-stu-id="fff2d-1018">Boolean</span></span>||<span data-ttu-id="fff2d-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="fff2d-1021">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-1021">String</span></span>||<span data-ttu-id="fff2d-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="fff2d-1025">function</span><span class="sxs-lookup"><span data-stu-id="fff2d-1025">function</span></span>|<span data-ttu-id="fff2d-1026">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1026">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1027">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1027">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1028">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1028">Requirements</span></span>

|<span data-ttu-id="fff2d-1029">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1029">Requirement</span></span>|<span data-ttu-id="fff2d-1030">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1030">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1031">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-1031">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1032">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-1032">1.0</span></span>|
|[<span data-ttu-id="fff2d-1033">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1033">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1034">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1034">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1035">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1035">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1036">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1036">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fff2d-1037">Примеры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1037">Examples</span></span>

<span data-ttu-id="fff2d-1038">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1038">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="fff2d-1039">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1039">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="fff2d-1040">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1040">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fff2d-1041">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1041">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fff2d-1042">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1042">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fff2d-1043">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1043">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="fff2d-1044">Жетаттачментконтентасинк (attachmentId, [параметры], [callback]) → [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="fff2d-1044">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="fff2d-1045">Получает указанное вложение из сообщения или встречи и возвращает его в виде `AttachmentContent` объекта.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1045">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="fff2d-1046">`getAttachmentContentAsync` Метод получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1046">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="fff2d-1047">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, когда Аттачментидс был получен с помощью вызова `getAttachmentsAsync` или. `item.attachments`</span><span class="sxs-lookup"><span data-stu-id="fff2d-1047">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="fff2d-1048">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1048">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="fff2d-1049">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1049">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1050">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1050">Parameters</span></span>

|<span data-ttu-id="fff2d-1051">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1051">Name</span></span>|<span data-ttu-id="fff2d-1052">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1052">Type</span></span>|<span data-ttu-id="fff2d-1053">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1053">Attributes</span></span>|<span data-ttu-id="fff2d-1054">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1054">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="fff2d-1055">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-1055">String</span></span>||<span data-ttu-id="fff2d-1056">Идентификатор вложения, которое требуется получить.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1056">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="fff2d-1057">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1057">Object</span></span>|<span data-ttu-id="fff2d-1058">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1058">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1059">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1059">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-1060">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1060">Object</span></span>|<span data-ttu-id="fff2d-1061">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1061">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1062">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1062">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fff2d-1063">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-1063">function</span></span>|<span data-ttu-id="fff2d-1064">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1065">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1065">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1066">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1066">Requirements</span></span>

|<span data-ttu-id="fff2d-1067">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1067">Requirement</span></span>|<span data-ttu-id="fff2d-1068">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1069">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1070">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fff2d-1070">Preview</span></span>|
|[<span data-ttu-id="fff2d-1071">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1072">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1073">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1074">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1074">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1075">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1075">Returns:</span></span>

<span data-ttu-id="fff2d-1076">Тип: [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="fff2d-1076">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="fff2d-1077">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1077">Example</span></span>

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var options = {asyncContext: {type: result.value[i].attachmentType}};
      getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  if (result.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="fff2d-1078">Жетаттачментсасинк ([параметры], [обратный вызов]) → Array. _Лт_[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fff2d-1078">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="fff2d-1079">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1079">Gets the item's attachments as an array.</span></span> <span data-ttu-id="fff2d-1080">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1080">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1081">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1081">Parameters</span></span>

|<span data-ttu-id="fff2d-1082">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1082">Name</span></span>|<span data-ttu-id="fff2d-1083">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1083">Type</span></span>|<span data-ttu-id="fff2d-1084">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1084">Attributes</span></span>|<span data-ttu-id="fff2d-1085">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1085">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="fff2d-1086">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1086">Object</span></span>|<span data-ttu-id="fff2d-1087">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1088">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1088">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-1089">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1089">Object</span></span>|<span data-ttu-id="fff2d-1090">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1091">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1091">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fff2d-1092">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-1092">function</span></span>|<span data-ttu-id="fff2d-1093">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1093">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1094">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1094">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1095">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1095">Requirements</span></span>

|<span data-ttu-id="fff2d-1096">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1096">Requirement</span></span>|<span data-ttu-id="fff2d-1097">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1098">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1099">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fff2d-1099">Preview</span></span>|
|[<span data-ttu-id="fff2d-1100">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1101">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1102">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1103">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1103">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1104">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1104">Returns:</span></span>

<span data-ttu-id="fff2d-1105">Тип: Array. _Лт_[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fff2d-1105">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="fff2d-1106">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1106">Example</span></span>

<span data-ttu-id="fff2d-1107">В приведенном ниже примере создается строка HTML со сведениями обо всех вложениях в текущем элементе.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1107">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="fff2d-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="fff2d-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="fff2d-1109">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1109">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1110">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1110">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-1111">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1111">Requirements</span></span>

|<span data-ttu-id="fff2d-1112">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1112">Requirement</span></span>|<span data-ttu-id="fff2d-1113">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1114">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1115">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-1115">1.0</span></span>|
|[<span data-ttu-id="fff2d-1116">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1116">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1117">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1117">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1118">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1118">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1119">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1119">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1120">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1120">Returns:</span></span>

<span data-ttu-id="fff2d-1121">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="fff2d-1121">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="fff2d-1122">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1122">Example</span></span>

<span data-ttu-id="fff2d-1123">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1123">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="fff2d-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fff2d-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fff2d-1125">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1125">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1126">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1127">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1127">Parameters</span></span>

|<span data-ttu-id="fff2d-1128">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1128">Name</span></span>|<span data-ttu-id="fff2d-1129">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1129">Type</span></span>|<span data-ttu-id="fff2d-1130">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1130">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="fff2d-1131">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="fff2d-1131">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="fff2d-1132">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1132">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1133">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1133">Requirements</span></span>

|<span data-ttu-id="fff2d-1134">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1134">Requirement</span></span>|<span data-ttu-id="fff2d-1135">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1136">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1137">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-1137">1.0</span></span>|
|[<span data-ttu-id="fff2d-1138">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1139">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="fff2d-1139">Restricted</span></span>|
|[<span data-ttu-id="fff2d-1140">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1141">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1142">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1142">Returns:</span></span>

<span data-ttu-id="fff2d-1143">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1143">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="fff2d-1144">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1144">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="fff2d-1145">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1145">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="fff2d-1146">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1146">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="fff2d-1147">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="fff2d-1147">Value of `entityType`</span></span>|<span data-ttu-id="fff2d-1148">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="fff2d-1148">Type of objects in returned array</span></span>|<span data-ttu-id="fff2d-1149">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1149">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="fff2d-1150">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-1150">String</span></span>|<span data-ttu-id="fff2d-1151">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="fff2d-1151">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="fff2d-1152">Contact</span><span class="sxs-lookup"><span data-stu-id="fff2d-1152">Contact</span></span>|<span data-ttu-id="fff2d-1153">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fff2d-1153">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="fff2d-1154">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-1154">String</span></span>|<span data-ttu-id="fff2d-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fff2d-1155">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="fff2d-1156">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="fff2d-1156">MeetingSuggestion</span></span>|<span data-ttu-id="fff2d-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fff2d-1157">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="fff2d-1158">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="fff2d-1158">PhoneNumber</span></span>|<span data-ttu-id="fff2d-1159">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="fff2d-1159">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="fff2d-1160">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="fff2d-1160">TaskSuggestion</span></span>|<span data-ttu-id="fff2d-1161">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fff2d-1161">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="fff2d-1162">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-1162">String</span></span>|<span data-ttu-id="fff2d-1163">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="fff2d-1163">**Restricted**</span></span>|

<span data-ttu-id="fff2d-1164">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fff2d-1164">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="fff2d-1165">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1165">Example</span></span>

<span data-ttu-id="fff2d-1166">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1166">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="fff2d-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fff2d-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fff2d-1168">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1168">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1169">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1169">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fff2d-1170">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1170">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1171">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1171">Parameters</span></span>

|<span data-ttu-id="fff2d-1172">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1172">Name</span></span>|<span data-ttu-id="fff2d-1173">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1173">Type</span></span>|<span data-ttu-id="fff2d-1174">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1174">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="fff2d-1175">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-1175">String</span></span>|<span data-ttu-id="fff2d-1176">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1176">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1177">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1177">Requirements</span></span>

|<span data-ttu-id="fff2d-1178">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1178">Requirement</span></span>|<span data-ttu-id="fff2d-1179">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1179">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1180">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-1180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1181">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-1181">1.0</span></span>|
|[<span data-ttu-id="fff2d-1182">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1183">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1185">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1185">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1186">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1186">Returns:</span></span>

<span data-ttu-id="fff2d-p164">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="fff2d-1189">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fff2d-1189">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="fff2d-1190">getInitializationContextAsync ([параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="fff2d-1190">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="fff2d-1191">Получает данные инициализации, передаваемые при активации надстройки [сообщением с действиями](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1191">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1192">Этот метод поддерживается только Outlook 2016 или более поздней версии для Windows (версии "нажми и работай" более поздней версии, чем 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1192">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1193">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1193">Parameters</span></span>

|<span data-ttu-id="fff2d-1194">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1194">Name</span></span>|<span data-ttu-id="fff2d-1195">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1195">Type</span></span>|<span data-ttu-id="fff2d-1196">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1196">Attributes</span></span>|<span data-ttu-id="fff2d-1197">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1197">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="fff2d-1198">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1198">Object</span></span>|<span data-ttu-id="fff2d-1199">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1199">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1200">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1200">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-1201">Объект</span><span class="sxs-lookup"><span data-stu-id="fff2d-1201">Object</span></span>|<span data-ttu-id="fff2d-1202">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1202">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1203">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1203">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fff2d-1204">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-1204">function</span></span>|<span data-ttu-id="fff2d-1205">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1205">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1206">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1206">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fff2d-1207">При успешном выполнении данные инициализации предоставляются в `asyncResult.value` свойстве в виде строки.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1207">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="fff2d-1208">Если `asyncResult` контекст инициализации отсутствует, объект будет содержать `Error` объект со `code` свойством, `9020` `name` для свойства которого задано значение. `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="fff2d-1208">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1209">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1209">Requirements</span></span>

|<span data-ttu-id="fff2d-1210">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1210">Requirement</span></span>|<span data-ttu-id="fff2d-1211">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1212">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1213">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fff2d-1213">Preview</span></span>|
|[<span data-ttu-id="fff2d-1214">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1215">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1216">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1217">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-1218">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1218">Example</span></span>

```javascript
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

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="fff2d-1219">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="fff2d-1219">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="fff2d-1220">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1220">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1221">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1221">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fff2d-p165">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="fff2d-1225">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1225">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="fff2d-1226">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1226">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="fff2d-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-1230">Requirements</span><span class="sxs-lookup"><span data-stu-id="fff2d-1230">Requirements</span></span>

|<span data-ttu-id="fff2d-1231">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1231">Requirement</span></span>|<span data-ttu-id="fff2d-1232">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1233">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1234">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-1234">1.0</span></span>|
|[<span data-ttu-id="fff2d-1235">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1236">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1238">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1238">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1239">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1239">Returns:</span></span>

<span data-ttu-id="fff2d-p167">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="fff2d-1242">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="fff2d-1242">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fff2d-1243">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1243">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fff2d-1244">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1244">Example</span></span>

<span data-ttu-id="fff2d-1245">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1245">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="fff2d-1246">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="fff2d-1246">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="fff2d-1247">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1247">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1248">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1248">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fff2d-1249">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1249">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="fff2d-p168">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1252">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1252">Parameters</span></span>

|<span data-ttu-id="fff2d-1253">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1253">Name</span></span>|<span data-ttu-id="fff2d-1254">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1254">Type</span></span>|<span data-ttu-id="fff2d-1255">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1255">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="fff2d-1256">Строка</span><span class="sxs-lookup"><span data-stu-id="fff2d-1256">String</span></span>|<span data-ttu-id="fff2d-1257">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1257">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1258">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1258">Requirements</span></span>

|<span data-ttu-id="fff2d-1259">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1259">Requirement</span></span>|<span data-ttu-id="fff2d-1260">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1260">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-1261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1262">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-1262">1.0</span></span>|
|[<span data-ttu-id="fff2d-1263">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1263">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1264">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1265">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1265">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1266">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1266">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1267">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1267">Returns:</span></span>

<span data-ttu-id="fff2d-1268">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1268">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="fff2d-1269">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="fff2d-1269">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fff2d-1270">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="fff2d-1270">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fff2d-1271">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1271">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="fff2d-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="fff2d-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="fff2d-1273">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1273">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="fff2d-p169">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1276">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1276">Parameters</span></span>

|<span data-ttu-id="fff2d-1277">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1277">Name</span></span>|<span data-ttu-id="fff2d-1278">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1278">Type</span></span>|<span data-ttu-id="fff2d-1279">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1279">Attributes</span></span>|<span data-ttu-id="fff2d-1280">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1280">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="fff2d-1281">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fff2d-1281">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="fff2d-p170">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="fff2d-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="fff2d-1285">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1285">Object</span></span>|<span data-ttu-id="fff2d-1286">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1286">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1287">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1287">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-1288">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1288">Object</span></span>|<span data-ttu-id="fff2d-1289">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1289">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1290">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1290">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fff2d-1291">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-1291">function</span></span>||<span data-ttu-id="fff2d-1292">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1292">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fff2d-1293">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1293">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="fff2d-1294">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1294">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1295">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1295">Requirements</span></span>

|<span data-ttu-id="fff2d-1296">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1296">Requirement</span></span>|<span data-ttu-id="fff2d-1297">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1298">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1299">1.2</span><span class="sxs-lookup"><span data-stu-id="fff2d-1299">1.2</span></span>|
|[<span data-ttu-id="fff2d-1300">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1301">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1301">ReadWriteItem</span></span>|
|[<span data-ttu-id="fff2d-1302">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1303">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1303">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1304">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1304">Returns:</span></span>

<span data-ttu-id="fff2d-1305">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1305">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="fff2d-1306">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="fff2d-1306">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fff2d-1307">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-1307">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fff2d-1308">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1308">Example</span></span>

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

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="fff2d-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="fff2d-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="fff2d-1310">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1310">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="fff2d-1311">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1311">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1312">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1312">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-1313">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1313">Requirements</span></span>

|<span data-ttu-id="fff2d-1314">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1314">Requirement</span></span>|<span data-ttu-id="fff2d-1315">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1315">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1316">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1317">1.6</span><span class="sxs-lookup"><span data-stu-id="fff2d-1317">1.6</span></span>|
|[<span data-ttu-id="fff2d-1318">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1319">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1320">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1321">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1321">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1322">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1322">Returns:</span></span>

<span data-ttu-id="fff2d-1323">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="fff2d-1323">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="fff2d-1324">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1324">Example</span></span>

<span data-ttu-id="fff2d-1325">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1325">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="fff2d-1326">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="fff2d-1326">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="fff2d-p173">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="fff2d-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1329">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1329">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fff2d-p174">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="fff2d-1333">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1333">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="fff2d-1334">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1334">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="fff2d-p175">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fff2d-1338">Requirements</span><span class="sxs-lookup"><span data-stu-id="fff2d-1338">Requirements</span></span>

|<span data-ttu-id="fff2d-1339">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1339">Requirement</span></span>|<span data-ttu-id="fff2d-1340">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1340">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1341">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1342">1.6</span><span class="sxs-lookup"><span data-stu-id="fff2d-1342">1.6</span></span>|
|[<span data-ttu-id="fff2d-1343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1344">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1346">Чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1346">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fff2d-1347">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1347">Returns:</span></span>

<span data-ttu-id="fff2d-p176">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="fff2d-1350">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1350">Example</span></span>

<span data-ttu-id="fff2d-1351">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1351">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="fff2d-1352">Жетшаредпропертиесасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="fff2d-1352">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="fff2d-1353">Получает свойства выбранной встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1353">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1354">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1354">Parameters</span></span>

|<span data-ttu-id="fff2d-1355">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1355">Name</span></span>|<span data-ttu-id="fff2d-1356">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1356">Type</span></span>|<span data-ttu-id="fff2d-1357">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1357">Attributes</span></span>|<span data-ttu-id="fff2d-1358">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1358">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="fff2d-1359">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1359">Object</span></span>|<span data-ttu-id="fff2d-1360">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1360">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1361">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1361">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-1362">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1362">Object</span></span>|<span data-ttu-id="fff2d-1363">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1363">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1364">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1364">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fff2d-1365">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-1365">function</span></span>||<span data-ttu-id="fff2d-1366">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fff2d-1367">Общие свойства предоставляются в виде [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) объекта в `asyncResult.value` свойстве.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1367">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="fff2d-1368">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1368">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1369">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1369">Requirements</span></span>

|<span data-ttu-id="fff2d-1370">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1370">Requirement</span></span>|<span data-ttu-id="fff2d-1371">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1371">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1372">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1373">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fff2d-1373">Preview</span></span>|
|[<span data-ttu-id="fff2d-1374">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1374">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1375">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1376">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1376">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1377">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-1378">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1378">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="fff2d-1379">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fff2d-1379">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="fff2d-1380">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1380">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="fff2d-p178">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1384">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1384">Parameters</span></span>

|<span data-ttu-id="fff2d-1385">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1385">Name</span></span>|<span data-ttu-id="fff2d-1386">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1386">Type</span></span>|<span data-ttu-id="fff2d-1387">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1387">Attributes</span></span>|<span data-ttu-id="fff2d-1388">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1388">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="fff2d-1389">function</span><span class="sxs-lookup"><span data-stu-id="fff2d-1389">function</span></span>||<span data-ttu-id="fff2d-1390">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1390">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fff2d-1391">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1391">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="fff2d-1392">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1392">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="fff2d-1393">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1393">Object</span></span>|<span data-ttu-id="fff2d-1394">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1394">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1395">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1395">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="fff2d-1396">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1396">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1397">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1397">Requirements</span></span>

|<span data-ttu-id="fff2d-1398">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1398">Requirement</span></span>|<span data-ttu-id="fff2d-1399">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1399">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1400">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1401">1.0</span><span class="sxs-lookup"><span data-stu-id="fff2d-1401">1.0</span></span>|
|[<span data-ttu-id="fff2d-1402">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1403">ReadItem</span></span>|
|[<span data-ttu-id="fff2d-1404">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1405">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1405">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-1406">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1406">Example</span></span>

<span data-ttu-id="fff2d-p181">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

---
---

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="fff2d-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fff2d-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="fff2d-1411">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1411">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="fff2d-1412">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1412">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="fff2d-1413">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1413">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="fff2d-1414">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1414">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="fff2d-1415">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1415">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1416">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1416">Parameters</span></span>

|<span data-ttu-id="fff2d-1417">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1417">Name</span></span>|<span data-ttu-id="fff2d-1418">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1418">Type</span></span>|<span data-ttu-id="fff2d-1419">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1419">Attributes</span></span>|<span data-ttu-id="fff2d-1420">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1420">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="fff2d-1421">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-1421">String</span></span>||<span data-ttu-id="fff2d-1422">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1422">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="fff2d-1423">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1423">Object</span></span>|<span data-ttu-id="fff2d-1424">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1424">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1425">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1425">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-1426">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1426">Object</span></span>|<span data-ttu-id="fff2d-1427">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1427">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1428">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1428">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fff2d-1429">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-1429">function</span></span>|<span data-ttu-id="fff2d-1430">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1430">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1431">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1431">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fff2d-1432">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1432">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fff2d-1433">Ошибки</span><span class="sxs-lookup"><span data-stu-id="fff2d-1433">Errors</span></span>

|<span data-ttu-id="fff2d-1434">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="fff2d-1434">Error code</span></span>|<span data-ttu-id="fff2d-1435">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1435">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="fff2d-1436">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1436">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1437">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1437">Requirements</span></span>

|<span data-ttu-id="fff2d-1438">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1438">Requirement</span></span>|<span data-ttu-id="fff2d-1439">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1439">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1440">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fff2d-1440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1441">1.1</span><span class="sxs-lookup"><span data-stu-id="fff2d-1441">1.1</span></span>|
|[<span data-ttu-id="fff2d-1442">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1443">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1443">ReadWriteItem</span></span>|
|[<span data-ttu-id="fff2d-1444">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1445">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1445">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-1446">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1446">Example</span></span>

<span data-ttu-id="fff2d-1447">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="fff2d-1447">The following code removes an attachment with an identifier of '0'.</span></span>

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

---
---

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="fff2d-1448">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fff2d-1448">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="fff2d-1449">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1449">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="fff2d-1450">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1450">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1451">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1451">Parameters</span></span>

| <span data-ttu-id="fff2d-1452">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1452">Name</span></span> | <span data-ttu-id="fff2d-1453">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1453">Type</span></span> | <span data-ttu-id="fff2d-1454">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1454">Attributes</span></span> | <span data-ttu-id="fff2d-1455">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1455">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="fff2d-1456">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="fff2d-1456">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="fff2d-1457">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1457">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="fff2d-1458">Объект</span><span class="sxs-lookup"><span data-stu-id="fff2d-1458">Object</span></span> | <span data-ttu-id="fff2d-1459">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1459">&lt;optional&gt;</span></span> | <span data-ttu-id="fff2d-1460">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1460">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="fff2d-1461">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1461">Object</span></span> | <span data-ttu-id="fff2d-1462">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1462">&lt;optional&gt;</span></span> | <span data-ttu-id="fff2d-1463">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1463">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="fff2d-1464">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-1464">function</span></span>| <span data-ttu-id="fff2d-1465">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1466">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1466">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1467">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1467">Requirements</span></span>

|<span data-ttu-id="fff2d-1468">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1468">Requirement</span></span>| <span data-ttu-id="fff2d-1469">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1470">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fff2d-1471">1.7</span><span class="sxs-lookup"><span data-stu-id="fff2d-1471">1.7</span></span> |
|[<span data-ttu-id="fff2d-1472">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fff2d-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1473">ReadItem</span></span> |
|[<span data-ttu-id="fff2d-1474">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fff2d-1475">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1475">Compose or Read</span></span> |

---
---

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="fff2d-1476">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="fff2d-1476">saveAsync([options], callback)</span></span>

<span data-ttu-id="fff2d-1477">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1477">Asynchronously saves an item.</span></span>

<span data-ttu-id="fff2d-p183">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1481">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1481">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="fff2d-1482">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1482">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="fff2d-p185">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="fff2d-1486">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="fff2d-1486">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="fff2d-1487">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1487">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="fff2d-1488">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1488">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="fff2d-1489">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1489">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1490">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1490">Parameters</span></span>

|<span data-ttu-id="fff2d-1491">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1491">Name</span></span>|<span data-ttu-id="fff2d-1492">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1492">Type</span></span>|<span data-ttu-id="fff2d-1493">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1493">Attributes</span></span>|<span data-ttu-id="fff2d-1494">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1494">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="fff2d-1495">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1495">Object</span></span>|<span data-ttu-id="fff2d-1496">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1496">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1497">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1497">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-1498">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1498">Object</span></span>|<span data-ttu-id="fff2d-1499">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1499">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1500">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1500">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fff2d-1501">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-1501">function</span></span>||<span data-ttu-id="fff2d-1502">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1502">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fff2d-1503">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1503">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1504">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1504">Requirements</span></span>

|<span data-ttu-id="fff2d-1505">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1505">Requirement</span></span>|<span data-ttu-id="fff2d-1506">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1506">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1507">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1508">1.3</span><span class="sxs-lookup"><span data-stu-id="fff2d-1508">1.3</span></span>|
|[<span data-ttu-id="fff2d-1509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1510">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1510">ReadWriteItem</span></span>|
|[<span data-ttu-id="fff2d-1511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1512">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1512">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fff2d-1513">Примеры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1513">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="fff2d-p187">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="fff2d-1516">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="fff2d-1516">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="fff2d-1517">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1517">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="fff2d-p188">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fff2d-1521">Параметры</span><span class="sxs-lookup"><span data-stu-id="fff2d-1521">Parameters</span></span>

|<span data-ttu-id="fff2d-1522">Имя</span><span class="sxs-lookup"><span data-stu-id="fff2d-1522">Name</span></span>|<span data-ttu-id="fff2d-1523">Тип</span><span class="sxs-lookup"><span data-stu-id="fff2d-1523">Type</span></span>|<span data-ttu-id="fff2d-1524">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fff2d-1524">Attributes</span></span>|<span data-ttu-id="fff2d-1525">Описание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1525">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="fff2d-1526">String</span><span class="sxs-lookup"><span data-stu-id="fff2d-1526">String</span></span>||<span data-ttu-id="fff2d-p189">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="fff2d-1530">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1530">Object</span></span>|<span data-ttu-id="fff2d-1531">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1531">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1532">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fff2d-1533">Object</span><span class="sxs-lookup"><span data-stu-id="fff2d-1533">Object</span></span>|<span data-ttu-id="fff2d-1534">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1534">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-1535">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="fff2d-1536">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fff2d-1536">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="fff2d-1537">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fff2d-1537">&lt;optional&gt;</span></span>|<span data-ttu-id="fff2d-p190">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="fff2d-p191">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="fff2d-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="fff2d-1542">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="fff2d-1542">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="fff2d-1543">функция</span><span class="sxs-lookup"><span data-stu-id="fff2d-1543">function</span></span>||<span data-ttu-id="fff2d-1544">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fff2d-1544">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fff2d-1545">Требования</span><span class="sxs-lookup"><span data-stu-id="fff2d-1545">Requirements</span></span>

|<span data-ttu-id="fff2d-1546">Требование</span><span class="sxs-lookup"><span data-stu-id="fff2d-1546">Requirement</span></span>|<span data-ttu-id="fff2d-1547">Значение</span><span class="sxs-lookup"><span data-stu-id="fff2d-1547">Value</span></span>|
|---|---|
|[<span data-ttu-id="fff2d-1548">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fff2d-1548">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fff2d-1549">1.2</span><span class="sxs-lookup"><span data-stu-id="fff2d-1549">1.2</span></span>|
|[<span data-ttu-id="fff2d-1550">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fff2d-1550">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fff2d-1551">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fff2d-1551">ReadWriteItem</span></span>|
|[<span data-ttu-id="fff2d-1552">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fff2d-1552">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="fff2d-1553">Создание</span><span class="sxs-lookup"><span data-stu-id="fff2d-1553">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fff2d-1554">Пример</span><span class="sxs-lookup"><span data-stu-id="fff2d-1554">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
